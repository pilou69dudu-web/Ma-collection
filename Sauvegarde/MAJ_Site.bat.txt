@echo off
:: Permet d'afficher correctement les accents en français dans la console
chcp 65001 > nul
title Gestion de la Vinylothèque

:menu
cls
echo =====================================================================
echo                       MENU GESTION DE LA COLLECTION
echo =====================================================================
echo.
echo [1] Générer la Page Web et envoyer sur GitHub
echo [2] Mise a jour des prix haut (Discogs API - Excel)
echo [3] Quitter
echo.
echo =====================================================================
echo.

set /p choix="Votre choix [1-3] : "

if "%choix%"=="1" goto option1
if "%choix%"=="2" goto option2
if "%choix%"=="3" goto quitter

:: Si l'utilisateur tape autre chose que 1, 2 ou 3
echo ❌ Choix invalide, veuillez recommencer...
timeout /t 2 > nul
goto menu

:option1
cls
echo =====================================================================
echo             GÉNÉRATION DE LA PAGE WEB ET ENVOI SUR GITHUB
echo =====================================================================
echo.
echo 1. Génération de la collection de Vinyles...
echo --------------------------------------------------
python generateur.py

echo.
echo 2. Envoi des mises à jour sur GitHub...
echo --------------------------------------------------
:: Ajoute les fichiers modifiés et les nouvelles pochettes
git add .
git commit -m "Mise a jour automatique de la collection et de la page Wanted"
git push origin main

echo.
echo --------------------------------------------------
echo Terminé ! Votre application et votre site sont à jour.
echo --------------------------------------------------
pause
goto menu

:option2
cls
echo =====================================================================
echo                     MISE À JOUR DES PRIX DEPUIS DISCOGS
echo =====================================================================
echo.
python update_prices.py
echo.
echo Fin de l'exécution.
pause
goto menu

:quitter
exit