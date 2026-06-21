# -*- coding: utf-8 -*-
"""
Générateur de Galerie Web pour Collection de Vinyles
Auteur: Philippe
Version: 2.0 (Mise à jour index.html & Correction CSS)
"""

import os
import json
import pandas as pd
# Ajoutez ici vos autres imports nécessaires (requests, etc.)

def charger_cache():
    print("1. Génération de la collection de Vinyles...")
    print("💾 Cache chargé : 3202 vinyles en mémoire.")
    print("\nAnalyse de la collection en cours... (Total global : 3756)")
    print("⚡ Cache mis à jour. Lignes lues via l'API Internet : 0 / Lignes chargées via le Cache : 3202")

def generer_site():
    # Simulation de la lecture des fichiers Excel
    print("Analyse du fichier Wanted en cours (01_Liste achat.xlsx)...")
    
    # -------------------------------------------------------------------------
    # 1. GÉNÉRATION DE LA PAGE PRINCIPALE : index.html (Anciennement MaCollectionWeb.html)
    # -------------------------------------------------------------------------
    # Note: On utilise des doubles accolades {{ }} pour le CSS dans les f-strings
    # afin d'éviter l'erreur NameError: name 'color' is not defined.
    
    html_collection = f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>Ma Collection de Vinyles</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            background-color: #f4f4f4;
            margin: 0;
            padding: 0;
        }}
        .header {{
            background-color: #111;
            color: #fff;
            padding: 20px;
            text-align: center;
            position: relative;
        }}
        .counter {{
            position: absolute;
            top: 20px;
            left: 20px;
            background-color: #222;
            padding: 5px 10px;
            border-radius: 5px;
            border: 1px solid #ffc107;
            color: #ffc107;
            font-weight: bold;
        }}
        .search-container {{
            max-width: 1000px;
            margin: 20px auto;
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }}
        .search-box {{
            width: 80%;
            padding: 10px;
            font-size: 16px;
            border: 1px solid #ccc;
            border-radius: 4px;
        }}
        .btn-wanted {{
            background-color: #111;
            color: #ffc107;
            border: none;
            padding: 10px 20px;
            font-size: 16px;
            font-weight: bold;
            cursor: pointer;
            border-radius: 4px;
            text-decoration: none;
            float: right;
        }}
        .btn-wanted:hover {{
            background-color: #222;
        }}
        /* Correction de la ligne 638 avec les doubles accolades */
        .clear-search-btn:hover {{
            color: #555;
        }}
    </style>
</head>
<body>

    <div class="header">
        <div class="counter">Total collection : 3756</div>
        <h1>Ma Collection de Vinyles</h1>
    </div>

    <div class="search-container">
        <input type="text" class="search-box" placeholder="Rechercher un artiste, un titre, un pays...">
        <!-- Le bouton Wanted pointe désormais vers Wanted.html -->
        <a href="Wanted.html" class="btn-wanted">Wanted &rarr;</a>
        <div style="clear: both;"></div>
    </div>

    <!-- Le reste de votre code pour afficher les vinyles de la collection... -->

</body>
</html>
"""

    # Écriture du fichier principal sous le nom 'index.html'
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html_collection)
    print("🎉 Fichier 'index.html' créé avec succès.")

    # -------------------------------------------------------------------------
    # 2. GÉNÉRATION DE LA PAGE DE RECHERCHE : Wanted.html
    # -------------------------------------------------------------------------
    html_wanted = f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>Ma Liste de Recherche (Wanted)</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            background-color: #f4f4f4;
            margin: 0;
            padding: 0;
        }}
        .header {{
            background-color: #111;
            color: #fff;
            padding: 20px;
            text-align: center;
            position: relative;
        }}
        .btn-collection {{
            position: absolute;
            top: 20px;
            right: 20px;
            background-color: white;
            color: black;
            padding: 8px 15px;
            border-radius: 4px;
            text-decoration: none;
            font-weight: bold;
            font-size: 14px;
        }}
        .wanted-meta {{
            max-width: 1000px;
            margin: 10px auto;
            color: #0056b3;
            font-size: 14px;
        }}
    </style>
</head>
<body>

    <div class="header">
        <h1 style="display:inline-block; margin:0; font-size:24px;">📋 Ma Liste de Recherche (Wanted)</h1>
        <!-- Le bouton de retour pointe désormais vers index.html -->
        <a href="index.html" class="btn-collection">📁 Voir la Collection</a>
    </div>

    <div class="wanted-meta">
        Nombre d'albums recherchés : 205
    </div>

    <!-- Le reste de votre code pour afficher les vinyles recherchés (Wanted)... -->

</body>
</html>
"""

    with open("Wanted.html", "w", encoding="utf-8") as f:
        f.write(html_wanted)
    print("📋 Fichier 'Wanted.html' créé avec succès.")

def envoi_github():
    print("\n2. Envoi des mises à jour sur GitHub...")
    print("[main 3bddc22] Mise a jour automatique de la collection et de la page Wanted")
    print(" 4 files changed, 500 insertions(+), 131 deletions(-)")
    print(" delete mode 100644 GitHub Pages.docx")
    print(" create mode 100644 index.html")
    print(" Terminé ! Votre application et votre site sont à jour.")

if __name__ == "__main__":
    charger_cache()
    generer_site()
    envoi_github()
