import os
import openpyxl
import json

# 1. Chemins des fichiers
EXCEL_VINYLES = "00_Mes vinyles.xlsx"
EXCEL_ACHAT = "01_Liste achat.xlsx"
INDEX_HTML = "index.html"
WANTED_HTML = "Wanted.html"

print("=========================================================================")
print("               GÉNÉRATION DE LA PAGE WEB ET ENVOI SUR GITHUB             ")
print("=========================================================================")
print("\n1. Génération de la collection de Vinyles...")

# ---- CHARGEMENT DE LA COLLECTION PRINCIPALE ----
if not os.path.exists(EXCEL_VINYLES):
    print(f"Erreur : Le fichier {EXCEL_VINYLES} est introuvable.")
    exit()

wb_v = openpyxl.load_workbook(EXCEL_VINYLES, data_only=True)
sheet_v = wb_v.active

# Récupération du compteur global affiché en A1 (ex: 3741)
compteur_global = sheet_v["A1"].value if sheet_v["A1"].value else 0
print(f"🎯 Compteur récupéré en A1 : {compteur_global} disques au total.")
print("Analyse des lignes de vinyles...")

liste_vinyles = []
# On commence à la ligne 2 (la ligne 1 contient les entêtes) jusqu'à la ligne 3192
for row_idx in range(2, 3193):
    # Lecture des colonnes selon la structure de votre fichier Excel
    type_album = sheet_v.cell(row=row_idx, column=3).value     # Col C: Album/Compils/Single
    annee = sheet_v.cell(row=row_idx, column=4).value          # Col D: Année
    quantite = sheet_v.cell(row=row_idx, column=5).value       # Col E: Qté
    pays = sheet_v.cell(row=row_idx, column=6).value           # Col F: Pays
    genre = sheet_v.cell(row=row_idx, column=7).value          # Col G: Genre
    commentaire = sheet_v.cell(row=row_idx, column=8).value    # Col H: Picture / Commentaire
    artiste = sheet_v.cell(row=row_idx, column=9).value        # Col I: ARTISTE
    titre_face_b = sheet_v.cell(row=row_idx, column=13).value   # Col M: Titre Face B
    duree_b = sheet_v.cell(row=row_idx, column=14).value        # Col N: Durée B

    # On ignore les lignes complètement vides (sans artiste ni genre)
    if not artiste and not genre:
        continue

    # Construction de l'objet vinyle sécurisé
    vinyle = {
        "type": str(type_album or "Unknown").strip(),
        "annee": str(annee or ""),
        "quantite": str(quantite or "1"),
        "pays": str(pays or "").strip(),
        "genre": str(genre or "Unknown").strip(),
        "commentaire": str(commentaire or "").strip(),
        "artiste": str(artiste or "Unknown").strip(),
        "titre_face_b": str(titre_face_b or "").strip(),
        "duree_b": str(duree_b or "").strip()
    }
    liste_vinyles.append(vinyle)

total_charges = len(liste_vinyles)
print(f"🎉 Terminé ! Compteur Jaune = {compteur_global} | Lignes chargées = {total_charges}")


# ---- CHARGEMENT DE LA LISTE D'ACHAT (WANTED) ----
print(f"\nAnalyse du fichier Wanted en cours ({EXCEL_ACHAT})...")
liste_wanted = []

if os.path.exists(EXCEL_ACHAT):
    wb_w = openpyxl.load_workbook(EXCEL_ACHAT, data_only=True)
    sheet_w = wb_w.active
    
    # On parcourt les lignes de la liste d'achat
    for row_idx in range(2, sheet_w.max_row + 1):
        artiste_titre = sheet_w.cell(row=row_idx, column=1).value  # Col A: Artistes / Titres
        note_c = sheet_w.cell(row=row_idx, column=3).value         # Col C: Note ("J'ai" ou vide)
        url_image = sheet_w.cell(row=row_idx, column=6).value      # Col F: Lien image

        # CRITÈRE : On affiche uniquement si la colonne C ("Note") est VIDE
        if artiste_titre and not note_c:
            wanted_item = {
                "nom": str(artiste_titre).strip(),
                "image": str(url_image or "").strip()
            }
            liste_wanted.append(wanted_item)
else:
    print(f"Attention : Le fichier {EXCEL_ACHAT} n'a pas été trouvé. Page Wanted vide.")

# ---- SÉCURISATION ET CONVERSION EN JSON POUR LE JAVASCRIPT ----
# json.dumps convertit proprement les listes Python en tableaux JavaScript valides
json_vinyles = json.dumps(liste_vinyles, ensure_ascii=False)
json_wanted = json.dumps(liste_wanted, ensure_ascii=False)


# ---- GÉNÉRATION DU FICHIER INDEX.HTML ----
# Remplacer la structure HTML par la vôtre. L'astuce est d'injecter `json_vinyles` au bon endroit.
html_index_content = f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>Ma Collection de Vinyles</title>
    </head>
<body>
    <div class="header">
        <span class="total-badge">Total collection : {compteur_global}</span>
        <h1>Ma Collection de Vinyles</h1>
    </div>

    <div id="compteur-affichage">Disques uniques affichés : <span id="count">0</span></div>
    <div id="vinylo-container"></div>

    <script>
        // Injection sécurisée des données de l'Excel
        const mesVinyles = {json_vinyles};
        
        // Exemple simple de script d'affichage et filtrage à adapter à vos fonctions existantes
        document.getElementById('count').innerText = mesVinyles.length;
        const container = document.getElementById('vinylo-container');
        
        // Votre logique d'affichage des vignettes / pochettes se met ici
        console.log("Données chargées avec succès : ", mesVinyles);
    </script>
</body>
</html>
"""

with open(INDEX_HTML, "w", encoding="utf-8") as f:
    f.write(html_index_content)
print(f"💾 Fichier {INDEX_HTML} généré avec succès !")


# ---- GÉNÉRATION DU FICHIER WANTED.HTML ----
html_wanted_content = f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>Ma Liste de Recherche (Wanted)</title>
</head>
<body>
    <h1>Ma Liste de Recherche (Wanted)</h1>
    <div>Disques recherchés : <span id="wanted-count">0</span></div>
    <div id="wanted-container"></div>

    <script>
        const listeRecherche = {json_wanted};
        document.getElementById('wanted-count').innerText = listeRecherche.length;
        
        const container = document.getElementById('wanted-container');
        listeRecherche.forEach(item => {{
            // Affichage de l'artiste, du titre et de l'image (colonne F)
            const div = document.createElement('div');
            div.className = 'wanted-card';
            div.innerHTML = `
                <p>${{item.nom}}</p>
                ${{item.image ? `<img src="${{item.image}}" alt="${{item.nom}}" width="150">` : ''}}
            `;
            container.appendChild(div);
        }});
    </script>
</body>
</html>
"""

with open(WANTED_HTML, "w", encoding="utf-8") as f:
    f.write(html_wanted_content)
print(f"💾 Fichier {WANTED_HTML} généré avec succès ! Total recherchés : {len(liste_wanted)}")

print("\n2. Envoi des mises à jour sur GitHub...")
# Votre bloc de code Git existant prend le relais ici (os.system("git add..."))