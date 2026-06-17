import pandas as pd
import json
import os
import re
import requests
import time

# =====================================================================
# CONFIGURATION AUTOMATIQUE
# =====================================================================
nom_fichier_excel = "00_Mes vinyles.xlsx"

CONSUMER_KEY = "pTrgAPVrGOUYZbQrFbbh"
CONSUMER_SECRET = "XnsdSqnEoZQHJLjtEhVwffLFSYNJTYmV"
# =====================================================================

if not os.path.exists("pochettes"):
    os.makedirs("pochettes")

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

try:
    df_brut = pd.read_excel(nom_fichier_excel, header=None)
    valeur_a1 = df_brut.iloc[0, 0]
    total_vinyles = str(int(float(valeur_a1))) if not pd.isna(valeur_a1) else "Non spécifié"
    df = pd.read_excel(nom_fichier_excel)
except Exception as e:
    print(f"Impossible de lire le fichier Excel : {e}")
    exit()

# Détection automatique des colonnes
colonne_type = next((c for c in df.columns if any(k in str(c).lower() for k in ['album', 'compil', 'sing'])), df.columns[2])
colonne_titre_a = next((c for c in df.columns if 'titre' in str(c).lower() and any(k in str(c).lower() for k in ['face a', 'album'])), None)
if not colonne_titre_a:
    colonne_titre_a = next((c for c in df.columns if 'titre' in str(c).lower() or 'album' in str(c).lower()), df.columns[1])
colonne_lien = next((c for c in df.columns if 'lien' in str(c).lower()), 'Lien')

# Gestion de la colonne Genre
colonne_genre = next((c for c in df.columns if 'genre' in str(c).lower()), None)
if not colonne_genre and len(df.columns) > 6:
    colonne_genre = df.columns[6]

collection = []
liste_genres_uniques = set()
print(f"\nAnalyse en cours... (Total global : {total_vinyles})")

for index, row in df.iterrows():
    artiste = str(row.get('ARTISTE', '')).strip()
    if not artiste or artiste == 'nan':
        continue
    
    titre_a = str(row.get(colonne_titre_a, '')).strip() if colonne_titre_a else ""
    if titre_a == 'nan': titre_a = ""
        
    lien_discogs = row.get(colonne_lien, '#')
    id_discogs = extraire_id_discogs(lien_discogs)
    chemin_pochette = telecharger_pochette(id_discogs, artiste, titre_a)
    
    quantite = row.get('Qté', 1)
    try: quantite = int(quantite) if not pd.isna(quantite) else 1
    except: quantite = 1

    valeur_type = str(row.get(colonne_type, 'SINGLE')).strip().upper()
    if 'MEDLEY' in valeur_type: valeur_type = 'MEDLEY'
    elif 'JINGLE' in valeur_type: valeur_type = 'JINGLE'
    elif 'COMPIL' in valeur_type: valeur_type = 'COMPILS'
    elif valeur_type == 'NAN' or valeur_type == '': valeur_type = 'SINGLE'
        
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
        "comment": str(row.get('Prix Haut - Commentaires', ''))
    }
    collection.append(viny_data)

json_data = json.dumps(collection, ensure_ascii=False)
genres_tries = sorted(list(liste_genres_uniques))
json_genres = json.dumps(genres_tries, ensure_ascii=False)

# =====================================================================
# CRÉATION DU HTML SÉCURISÉ
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
        .search-box { width: 100%; padding: 12px; font-size: 16px; border: 2px solid var(--border-color); border-radius: 6px; outline: none; }
        .navigation-filters { display: flex; flex-direction: column; gap: 12px; padding-top: 10px; border-top: 1px dashed var(--border-color); margin-top: 15px; }
        .filter-row { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; }
        .filter-label { font-size: 13px; font-weight: bold; color: var(--text-muted); text-transform: uppercase; min-width: 90px; }
        .nav-btn { background: white; border: 1px solid var(--border-color); padding: 6px 12px; font-size: 14px; font-weight: 600; border-radius: 4px; cursor: pointer; }
        .nav-btn.active { background: var(--discogs-yellow); color: var(--discogs-black); border-color: var(--discogs-yellow); }
        
        /* Style du champ de recherche de genre */
        .genre-input { background: white; border: 1px solid var(--border-color); padding: 6px 12px; font-size: 14px; font-weight: 600; border-radius: 4px; outline: none; width: 220px; }
        .genre-input:focus { border-color: var(--discogs-yellow); }
        .genre-input.active { background: var(--discogs-yellow); color: var(--discogs-black); border-color: var(--discogs-yellow); }

        .vinyl-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(190px, 1fr)); gap: 15px; margin-top: 20px; }
        .vinyl-card { background: white; border: 1px solid var(--border-color); border-radius: 6px; overflow: hidden; display: flex; flex-direction: column; box-shadow: 0 3px 5px rgba(0,0,0,0.02); position: relative; }
        .badge-qte { position: absolute; top: 8px; right: 8px; background: linear-gradient(135deg, #ffe600, #ffb300); color: #111111; font-size: 13px; font-weight: 800; width: 28px; height: 28px; border-radius: 50%; display: flex; align-items: center; justify-content: center; border: 2px solid #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.4), inset 0 1px 0 rgba(255,255,255,0.6); z-index: 10; }
        .cover-wrapper { aspect-ratio: 1; background: #222; display: flex; align-items: center; justify-content: center; position: relative; border-bottom: 1px solid var(--border-color); overflow: hidden; }
        .cover-image { width: 100%; height: 100%; object-fit: cover; }
        .cover-placeholder { color: #777; font-size: 11px; font-weight: bold; padding: 10px; text-align: center; text-transform: uppercase; }
        .vinyl-details { padding: 10px; flex-grow: 1; display: flex; flex-direction: column; justify-content: space-between; }
        .tag-type { align-self: flex-start; font-size: 9px; font-weight: 700; text-transform: uppercase; padding: 1px 4px; border-radius: 2px; background: #e2d9f3; color: #432874; margin-bottom: 6px; }
        .vinyl-artist { font-size: 13px; font-weight: 700; text-transform: uppercase; line-height: 1.2; margin-bottom: 4px; }
        .meta-info { font-size: 11px; color: var(--text-muted); margin-bottom: 3px; line-height: 1.3; }
        .tracks-block { border-top: 1px solid var(--border-color); margin-top: 6px; padding-top: 6px; }
        .track-a { font-size: 11px; color: #111111; margin-bottom: 3px; line-height: 1.2; }
        .track-b { font-size: 11px; color: #111111; line-height: 1.2; }
        .discogs-link { display: inline-block; margin-top: 10px; width: 100%; text-align: center; background-color: var(--discogs-black); color: white; text-decoration: none; padding: 6px; font-size: 11px; font-weight: 600; border-radius: 4px; }
        .discogs-link:hover { background-color: var(--discogs-yellow); color: var(--discogs-black); }
        .scroll-to-top { position: fixed; bottom: 25px; right: 25px; background-color: var(--discogs-black); color: white; border: 2px solid var(--discogs-yellow); width: 45px; height: 45px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 20px; cursor: pointer; box-shadow: 0 4px 10px rgba(0,0,0,0.3); z-index: 1000; opacity: 0; visibility: hidden; transition: all 0.3s ease; }
        .scroll-to-top.visible { opacity: 1; visibility: visible; }
        .scroll-to-top:hover { background-color: var(--discogs-yellow); color: var(--discogs-black); transform: scale(1.1); }
        @media (max-width: 768px) { header { flex-direction: column; gap: 10px; text-align: center; } .global-counter { position: static; margin-bottom: 5px; } }
    </style>
</head>
<body>
    <header>
        <div class="global-counter">Total collection : <span id="totalCounter">0</span></div>
        <h1>Ma Collection de Vinyles</h1>
    </header>
    <div class="sticky-wrapper">
        <div class="container">
            <div class="search-container">
                <input type="text" id="searchBox" class="search-box" placeholder="Rechercher un artiste, un titre, un pays...">
                <div class="navigation-filters">
                    <div class="filter-row">
                        <span class="filter-label">Type :</span>
                        <button class="nav-btn type-filter active" data-type="ALL">Tous</button>
                        <button class="nav-btn type-filter" data-type="ALBUM">Album</button>
                        <button class="nav-btn type-filter" data-type="COMPILS">Compils</button>
                        <button class="nav-btn type-filter" data-type="DOUBLE">Double</button>
                        <button class="nav-btn type-filter" data-type="INTERVIEW">Interview</button>
                        <button class="nav-btn type-filter" data-type="JINGLE">Jingle</button>
                        <button class="nav-btn type-filter" data-type="MEDLEY">Medley</button>
                        <button class="nav-btn type-filter" data-type="SINGLE">Single</button>
                    </div>
                    <div class="filter-row">
                        <span class="filter-label">Genre :</span>
                        <input type="text" id="genreInput" class="genre-input" list="genresList" placeholder="Filtrer par genre...">
                        <datalist id="genresList"></datalist>
                    </div>
                    <div class="filter-row">
                        <span class="filter-label">Alphabet :</span>
                        <button class="nav-btn alpha-filter active" data-alpha="ALL">Tous</button>
                        <div id="alphabetContainer" style="display:inline-flex; flex-wrap:wrap; gap:4px;"></div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <div class="container" style="margin-top: 20px;">
        <div style="margin-bottom:15px; font-size:14px; color:var(--text-muted);">Disques affichés : <span id="recordCount">0</span></div>
        <div class="vinyl-grid" id="vinylGrid"></div>
    </div>
    <button class="scroll-to-top" id="scrollTopBtn" title="Retour en haut">▲</button>
    <script>
"""

html_fin = """
        let currentSearch = "", currentType = "ALL", currentAlpha = "ALL", currentGenre = "ALL";
        
        // Remplissage automatique du datalist (suggestions de genres)
        const datalist = document.getElementById('genresList');
        genresAuto.forEach(genre => {
            const opt = document.createElement('option');
            opt.value = genre;
            datalist.appendChild(opt);
        });

        const container = document.getElementById('alphabetContainer');
        const alphaList = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
        alphaList.splice(alphaList.indexOf('S')+1, 0, 'The');
        alphaList.push('0 à 9');

        alphaList.forEach(item => {
            const b = document.createElement('button'); b.className = 'nav-btn alpha-filter';
            b.textContent = item; b.setAttribute('data-alpha', item); container.appendChild(b);
        });

        function renderGrid() {
            let filtered = vinylData.filter(item => {
                const sMatch = !currentSearch || 
                               (item.artist && item.artist.toLowerCase().includes(currentSearch)) || 
                               (item.titleA && item.titleA.toLowerCase().includes(currentSearch)) || 
                               (item.titleB && item.titleB.toLowerCase().includes(currentSearch)) || 
                               (item.country && item.country.toLowerCase().includes(currentSearch));
                
                const tMatch = currentType === "ALL" || item.type === currentType;
                
                const gMatch = currentGenre === "ALL" || (item.genre && item.genre.toUpperCase() === currentGenre);
                
                let aMatch = currentAlpha === "ALL";
                if(item.artist) {
                    if(currentAlpha === "The") aMatch = item.artist.toLowerCase().startsWith("the ");
                    else if(currentAlpha === "0 à 9") aMatch = /^[0-9]/.test(item.artist);
                    else if(!aMatch) aMatch = item.artist.toUpperCase().startsWith(currentAlpha);
                }
                return sMatch && tMatch && gMatch && aMatch;
            });

            document.getElementById('recordCount').textContent = filtered.length;
            document.getElementById('vinylGrid').innerHTML = filtered.map(item => {
                const badgeQte = item.qte > 1 ? '<div class="badge-qte">' + item.qte + '</div>' : '';
                const imgTag = (item.pochette && item.pochette !== "pochettes/placeholder.png")
                    ? '<img class="cover-image" src="' + item.pochette + '" alt="Pochette">'
                    : '<div class="cover-placeholder">💿 Image indisponible</div>';
                
                return '<div class="vinyl-card">' + badgeQte +
                        '<div class="cover-wrapper">' + imgTag + '</div>' +
                        '<div class="vinyl-details">' +
                            '<div>' +
                                '<div class="tag-type">' + item.type + '</div>' +
                                '<div class="vinyl-artist">' + (item.artist || '') + '</div>' +
                                '<div class="meta-info">' +
                                    '<strong>Genre :</strong> ' + (item.genre || 'N/C') + '<br>' +
                                    '<strong>Année :</strong> ' + (item.year || 'N/C') + '<br>' +
                                    '<strong>Pays :</strong> ' + ((item.country && item.country !== "nan") ? item.country : "N/C") + '<br>' +
                                    '<strong>Label :</strong> ' + (item.label || 'N/C') +
                                '</div>' +
                                '<div class="tracks-block">' +
                                    '<div class="track-a"><strong>Face A :</strong> <em>' + (item.titleA || 'Album / Inconnu') + '</em></div>' +
                                    '<div class="track-b"><strong>Face B :</strong> <em>' + (item.titleB || 'Inconnu') + '</em></div>' +
                                '</div>' +
                            '</div>' +
                            '<a href="' + (item.url || '#') + '" target="_blank" class="discogs-link">Voir sur Discogs</a>' +
                        '</div>' +
                    '</div>';
            }).join('');
        }

        const scrollTopBtn = document.getElementById('scrollTopBtn');
        window.addEventListener('scroll', () => { if (window.scrollY > 300) scrollTopBtn.classList.add('visible'); else scrollTopBtn.classList.remove('visible'); });
        scrollTopBtn.addEventListener('click', () => window.scrollTo({ top: 0, behavior: 'smooth' }));
        
        document.getElementById('searchBox').addEventListener('input', e => { currentSearch = e.target.value.toLowerCase(); renderGrid(); });
        
        document.querySelectorAll('.type-filter').forEach(btn => btn.addEventListener('click', e => { document.querySelectorAll('.type-filter').forEach(b => b.classList.remove('active')); e.target.classList.add('active'); currentType = e.target.getAttribute('data-type'); renderGrid(); }));
        
        // Écouteur pour la recherche prédictive du genre
        const gInput = document.getElementById('genreInput');
        gInput.addEventListener('input', e => {
            const val = e.target.value.trim().toUpperCase();
            
            // Si la valeur saisie correspond exactement à un genre existant ou si c'est vide
            if (val === "") {
                currentGenre = "ALL";
                gInput.classList.remove('active');
                renderGrid();
            } else if (genresAuto.includes(val)) {
                currentGenre = val;
                gInput.classList.add('active'); // Le champ devient jaune pour montrer qu'un filtre est actif
                renderGrid();
            } else {
                // Si l'utilisateur est en train de taper mais n'a pas encore validé un genre complet, 
                // on ne bloque pas le filtrage global pour laisser taper tranquillement.
            }
        });

        // Vider le champ si l'utilisateur double-clique dedans pour réinitialiser rapidement
        gInput.addEventListener('click', e => { if(e.target.value !== "") { e.target.value = ""; currentGenre = "ALL"; gInput.classList.remove('active'); renderGrid(); } });
        
        document.addEventListener('click', e => { if(e.target.classList.contains('alpha-filter')) { document.querySelectorAll('.alpha-filter').forEach(b => b.classList.remove('active')); e.target.classList.add('active'); currentAlpha = e.target.getAttribute('data-alpha'); renderGrid(); } });
        renderGrid();
    </script>
</body>
</html>
"""

# Assemblage final
with open("MaCollectionWeb.html", "w", encoding="utf-8") as f:
    f.write(html_debut)
    f.write(f"\n        const vinylData = {json_data};\n")
    f.write(f"        const genresAuto = {json_genres};\n")
    f.write(f"        document.getElementById('totalCounter').textContent = '{total_vinyles}';\n")
    f.write(html_fin)

print("\n🎉 Terminé ! Le champ de recherche prédictive pour les genres est en place.")