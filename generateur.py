import pandas as pd
import json
import os
import re
import requests
import time
import math
import hashlib  # Pour détecter les changements de lignes

# =====================================================================
# CONFIGURATION AUTOMATIQUE
# =====================================================================
nom_fichier_excel = "00_Mes vinyles.xlsx"
nom_fichier_wanted = "01_Liste achat.xlsx"
fichier_cache = "vinyles_cache.json"  # Fichier mémoire pour accélérer le script

CONSUMER_KEY = "pTrgAPVrGOUYZbQrFbbh"
CONSUMER_SECRET = "XnsdSqnEoZQHJLjtEhVwffLFSYNJTYmV"
# =====================================================================

if not os.path.exists("pochettes"):
    os.makedirs("pochettes")

# Chargement du cache s'il existe
cache_donnees = {}
if os.path.exists(fichier_cache):
    try:
        with open(fichier_cache, 'r', encoding='utf-8') as f:
            cache_donnees = json.load(f)
        print(f"💾 Cache chargé : {len(cache_donnees)} vinyles en mémoire.")
    except Exception as e:
        print(f"⚠️ Impossible de lire le cache, il sera réinitialisé : {e}")

def generer_hash_ligne(row):
    """Crée une empreinte unique de la ligne pour détecter le moindre changement (ex: modification de prix, de titre...)"""
    chaine_complete = "".join(str(val) for val in row.values)
    return hashlib.md5(chaine_complete.encode('utf-8')).hexdigest()

def extraire_id_discogs(url):
    if pd.isna(url):
        return None
    match = re.search(r"release/(\d+)", str(url))
    return match.group(1) if match else None

def telecharger_pochette(release_id, artiste, titre):
    if not release_id:
        return "pochettes/placeholder.png"
    
    chemin_image = f"pochettes/{release_id}.jpg"
    
    if os.path.exists(chemin_image) and os.path.getsize(chemin_image) > 0:
        return chemin_image
        
    try:
        url_api = f"https://api.discogs.com/releases/{release_id}"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Authorization": f"Discogs key={CONSUMER_KEY}, secret={CONSUMER_SECRET}"
        }
        reponse = requests.get(url_api, headers=headers)
        
        if reponse.status_code == 200:
            images = reponse.json().get('images', [])
            if images:
                url_image = images[0].get('uri') or images[0].get('resource_url')
                if url_image:
                    img_req = requests.get(url_image, headers=headers)
                    if img_req.status_code == 200:
                        with open(chemin_image, 'wb') as handler:
                            handler.write(img_req.content)
                        print(f"📸 Image récupérée : {artiste} - {titre}")
                        time.sleep(1.5)
                        return chemin_image
            
            print(f"⚠️ Pas d'image trouvée sur Discogs pour l'ID {release_id} ({artiste} - {titre})")
            
        elif reponse.status_code == 429:
            print("🛑 Discogs bloque temporairement (trop de requêtes). Pause de 10 secondes...")
            time.sleep(10)
        else:
            print(f"❌ Erreur API Discogs (Code {reponse.status_code}) pour l'ID {release_id} ({artiste} - {titre})")
            
    except Exception as e:
        print(f"⚠️ Impossible de vérifier l'ID {release_id} ({artiste}) : {e}")
        
    return "pochettes/placeholder.png"

def recuperer_prix_haut(release_id):
    if not release_id:
        return None
    try:
        url_api = f"https://api.discogs.com/marketplace/price_suggestions/{release_id}"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Authorization": f"Discogs key={CONSUMER_KEY}, secret={CONSUMER_SECRET}"
        }
        reponse = requests.get(url_api, headers=headers)
        
        if reponse.status_code == 200:
            data = reponse.json()
            prix_brut = None
            if "Mint (M)" in data:
                prix_brut = data["Mint (M)"]["value"]
            elif "Near Mint (NM or M-)" in data:
                prix_brut = data["Near Mint (NM or M-)"]["value"]
            elif data.values():
                prix_brut = max(item["value"] for item in data.values() if isinstance(item, dict) and "value" in item)
            
            if prix_brut:
                return f"{math.ceil(prix_brut)}€"
        elif reponse.status_code == 429:
            time.sleep(3)
    except:
        pass
    return None

# =====================================================================
# TRAITEMENT DU FICHIER PRINCIPAL (MES VINYLES)
# =====================================================================
try:
    df_brut = pd.read_excel(nom_fichier_excel, header=None)
    valeur_a1 = df_brut.iloc[0, 0]
    total_vinyles = str(int(float(valeur_a1))) if not pd.isna(valeur_a1) else "Non spécifié"
    df = pd.read_excel(nom_fichier_excel)
except Exception as e:
    print(f"Impossible de lire le fichier principal Excel : {e}")
    exit()

colonne_type = next((c for c in df.columns if any(k in str(c).lower() for k in ['album', 'compil', 'sing'])), df.columns[2])
colonne_titre_a = next((c for c in df.columns if 'titre' in str(c).lower() and any(k in str(c).lower() for k in ['face a', 'album'])), None)
if not colonne_titre_a:
    colonne_titre_a = next((c for c in df.columns if 'titre' in str(c).lower() or 'album' in str(c).lower()), df.columns[1])
colonne_lien = next((c for c in df.columns if 'lien' in str(c).lower()), 'Lien')

colonne_genre = next((c for c in df.columns if 'genre' in str(c).lower()), None)
if not colonne_genre and len(df.columns) > 6:
    colonne_genre = df.columns[6]

collection = []
nouveau_cache = {}
liste_genres_uniques = set()
liste_types_uniques = set() 

print(f"\nAnalyse de la collection en cours... (Total global : {total_vinyles})")

compteur_api = 0

for index, row in df.iterrows():
    artiste = str(row.get('ARTISTE', '')).strip()
    if not artiste or artiste == 'nan':
        continue
    
    # Calcul de l'empreinte de la ligne
    hash_ligne = generer_hash_ligne(row)
    
    # Si la ligne existe à l'identique dans l'ancien cache, on reprend les données sans appeler l'API
    if hash_ligne in cache_donnees:
        viny_data = cache_donnees[hash_ligne]
        nouveau_cache[hash_ligne] = viny_data
        collection.append(viny_data)
        
        # On extrait quand même le genre et le type pour les filtres de la page
        if viny_data["type"]: liste_types_uniques.add(viny_data["type"])
        if viny_data["genre"] and viny_data["genre"] != "N/C": liste_genres_uniques.add(viny_data["genre"].upper())
        continue

    # SINON : C'est une nouvelle ligne ou elle a été modifiée -> Appel API
    compteur_api += 1
    titre_a = str(row.get(colonne_titre_a, '')).strip() if colonne_titre_a else ""
    if titre_a == 'nan': titre_a = ""
        
    lien_discogs = row.get(colonne_lien, '#')
    id_discogs = extraire_id_discogs(lien_discogs)
    
    chemin_pochette = telecharger_pochette(id_discogs, artiste, titre_a)
    prix_affiche = recuperer_prix_haut(id_discogs)
    
    # Si Discogs ne renvoie pas de prix ou si vous souhaitez forcer l'affichage du prix Excel :
    if not prix_affiche or prix_affiche == "":
        valeur_prix_excel = str(row.get('Prix Haut', '')).strip()
        chiffres = re.findall(r"\d+", valeur_prix_excel)
        if chiffres:
            prix_affiche = f"{chiffres[0]}€"
        else:
            prix_affiche = ""

    quantite = row.get('Qté', 1)
    try: quantite = int(quantite) if not pd.isna(quantite) else 1
    except: quantite = 1

    valeur_type = str(row.get(colonne_type, 'SINGLE')).strip().upper()
    if 'MEDLEY' in valeur_type: valeur_type = 'MEDLEY'
    elif 'JINGLE' in valeur_type: valeur_type = 'JINGLE'
    elif 'COMPIL' in valeur_type: valeur_type = 'COMPILS'
    elif valeur_type == 'NAN' or valeur_type == '': valeur_type = 'SINGLE'
    
    if valeur_type:
        liste_types_uniques.add(valeur_type)
        
    genre_vinyl = str(row.get(colonne_genre, '')).strip()
    if genre_vinyl == 'nan' or genre_vinyl == '': 
        genre_vinyl = "N/C"
    else:
        liste_genres_uniques.add(genre_vinyl.upper())

    viny_data = {
        "id": str(row.get('N°', index)),
        "type": valeur_type,
        "year": str(row.get('Année', '')),
        "country": str(row.get('Pays', '')),
        "genre": genre_vinyl,
        "artist": artiste,
        "titleA": titre_a,
        "qte": quantite,
        "durationA": str(row.get('Durée A', '')),
        "bpmA": str(row.get('Bpm A', '')),
        "titleB": str(row.get('Titre Face B', '')),
        "durationB": str(row.get('Durée B', '')),
        "bpmB": str(row.get('Bpm B', '')),
        "label": str(row.get('Préssage / Labels', '')),
        "url": str(lien_discogs),
        "pochette": chemin_pochette,
        "prix": prix_affiche,
        "comment": str(row.get('Prix Haut - Commentaires', ''))
    }
    
    # Enregistrement dans le nouveau cache et la collection
    nouveau_cache[hash_ligne] = viny_data
    collection.append(viny_data)

# Sauvegarde du fichier cache mis à jour
try:
    with open(fichier_cache, 'w', encoding='utf-8') as f:
        json.dump(nouveau_cache, f, ensure_ascii=False, indent=4)
    print(f"⚡ Cache mis à jour. Lignes lues via l'API Internet : {compteur_api} / Lignes chargées via le Cache : {len(collection) - compteur_api}")
except Exception as e:
    print(f"⚠️ Impossible de sauvegarder le cache : {e}")

json_data = json.dumps(collection, ensure_ascii=False)
genres_tries = sorted(list(liste_genres_uniques))
json_genres = json.dumps(genres_tries, ensure_ascii=False)
types_tries = sorted(list(liste_types_uniques))
json_types = json.dumps(types_tries, ensure_ascii=False)


# =====================================================================
# TRAITEMENT DU FICHIER WANTED AVEC FILTRE COLONNE C VIDE (01_LISTE ACHAT.XLSX)
# =====================================================================
wanted_collection = []
print(f"\nAnalyse du fichier Wanted en cours ({nom_fichier_wanted})...")

if os.path.exists(nom_fichier_wanted):
    try:
        df_wanted = pd.read_excel(nom_fichier_wanted)
        
        for index, row in df_wanted.iterrows():
            valeur_colonne_c = row.iloc[2] if len(row) > 2 else None
            
            if pd.isna(valeur_colonne_c) or str(valeur_colonne_c).strip() == "" or str(valeur_colonne_c).lower() == "nan":
                if len(row) > 0 and not pd.isna(row.iloc[0]):
                    chaine_artiste_titre = str(row.iloc[0]).strip()
                    
                    if " - " in chaine_artiste_titre:
                        parts = chaine_artiste_titre.split(" - ", 1)
                        artiste_w = parts[0].strip()
                        titre_w = parts[1].strip()
                    else:
                        artiste_w = chaine_artiste_titre
                        titre_w = ""
                    
                    com_w = str(row.iloc[3]).strip() if len(row) > 3 and not pd.isna(row.iloc[3]) else ""
                    lien_w = str(row.iloc[5]).strip() if len(row) > 5 and not pd.isna(row.iloc[5]) else "#"
                    
                    # Pour la liste Wanted, on utilise le lien comme clé de cache
                    hash_wanted = hashlib.md5((artiste_w + titre_w + lien_w).encode('utf-8')).hexdigest()
                    
                    if hash_wanted in cache_donnees:
                        pochette_w = cache_donnees[hash_wanted].get('pochette', 'pochettes/placeholder.png')
                    else:
                        id_discogs_w = extraire_id_discogs(lien_w)
                        pochette_w = telecharger_pochette(id_discogs_w, artiste_w, titre_w)
                    
                    wanted_collection.append({
                        "artist": artiste_w,
                        "title": titre_w,
                        "comment": com_w,
                        "url": lien_w,
                        "pochette": pochette_w
                    })
    except Exception as e:
        print(f"⚠️ Erreur lors de la lecture du fichier Wanted : {e}")
else:
    print(f"⚠️ Fichier {nom_fichier_wanted} introuvable. Page Wanted créée vide.")

json_wanted_data = json.dumps(wanted_collection, ensure_ascii=False)


# =====================================================================
# SQUELETTE HTML : PAGE PRINCIPALE (COLLECTION)
# =====================================================================
html_debut = """<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>Ma Collection de Vinyles - Discogs Style</title>
    <style>
        :root { --discogs-black: #111111; --discogs-yellow: #f5c518; --light-bg: #f8f9fa; --border-color: #e5e7eb; --text-muted: #6b7280; }
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: sans-serif; }
        body { background-color: var(--light-bg); color: var(--discogs-black); padding-bottom: 50px; }
        header { background-color: var(--discogs-black); color: white; padding: 20px; border-bottom: 4px solid var(--discogs-yellow); position: relative; display: flex; align-items: center; justify-content: center; }
        header h1 { font-size: 24px; }
        .global-counter { position: absolute; left: 20px; background: rgba(255, 255, 255, 0.1); border: 1px solid rgba(255, 255, 255, 0.2); padding: 6px 14px; border-radius: 6px; font-size: 14px; font-weight: bold; color: #ffffff; }
        .global-counter span { color: var(--discogs-yellow); font-size: 16px; margin-left: 5px; }
        .sticky-wrapper { position: -webkit-sticky; position: sticky; top: 0; z-index: 100; background-color: var(--light-bg); padding-top: 15px; padding-bottom: 10px; border-bottom: 1px solid var(--border-color); box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
        .container { max-width: 1350px; margin: 0 auto; padding: 0 15px; }
        
        .search-container { background: white; padding: 20px; border-radius: 8px; border: 1px solid var(--border-color); box-shadow: 0 2px 4px rgba(0,0,0,0.02); }
        .search-row-wrapper { display: flex; gap: 15px; align-items: center; width: 100%; }
        
        .search-box-container { position: relative; flex-grow: 1; }
        .search-box { width: 100%; padding: 12px 40px 12px 12px; font-size: 16px; border: 2px solid var(--border-color); border-radius: 6px; outline: none; }
        
        .clear-search-btn { position: absolute; right: 12px; top: 50%; transform: translateY(-50%); background: none; border: none; font-size: 16px; color: #aaa; cursor: pointer; display: none; }
        .clear-search-btn:hover { color: #555; }
        
        .wanted-btn { background-color: var(--discogs-black); color: var(--discogs-yellow); border: 2px solid var(--discogs-black); padding: 11px 24px; font-size: 15px; font-weight: bold; border-radius: 6px; cursor: pointer; text-decoration: none; text-align: center; white-space: nowrap; transition: all 0.2s ease; }
        .wanted-btn:hover { background-color: var(--discogs-yellow); color: var(--discogs-black); }
        
        .navigation-filters { display: flex; flex-direction: column; gap: 12px; padding-top: 10px; border-top: 1px dashed var(--border-color); margin-top: 15px; }
        .filter-row { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; }
        .filter-label { font-size: 13px; font-weight: bold; color: var(--text-muted); text-transform: uppercase; min-width: 90px; }
        .nav-btn { background: white; border: 1px solid var(--border-color); padding: 6px 12px; font-size: 14px; font-weight: 600; border-radius: 4px; cursor: pointer; text-transform: capitalize; }
        .nav-btn.active { background: var(--discogs-yellow); color: var(--discogs-black); border-color: var(--discogs-yellow); }
        
        .custom-dropdown { position: relative; display: inline-block; }
        .dropdown-