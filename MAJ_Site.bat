@echo off
chcp 65001 > nul
cls

:: Se place automatiquement dans le dossier principal (Mes vinyles)
cd /d "%~dp0"

:MENU
echo ============================================================
echo                MENU GESTION DE LA COLLECTION
echo ============================================================
echo.
echo [1] Générer la Page Web et envoyer sur GitHub
echo [2] Mise a jour des prix haut (Discogs API - Excel)
echo [3] Transfert Chiner vers GitHub
echo [4] Quitter
echo.
echo ============================================================
echo.
set /p choix=Votre choix [1-4] : 

if "%choix%"=="1" goto OP1
if "%choix%"=="2" goto OP2
if "%choix%"=="3" goto OP3
if "%choix%"=="4" goto FIN

echo Choix invalide, veuillez réessayer.
pause
cls
goto MENU

:OP1
echo.
echo Exécution de la génération et envoi de la collection...
cd "Ma-collection"
python generateur.py
cd /d "%~dp0"
git add .
git commit -m "Mise a jour de la collection de vinyles"
git pull --rebase origin main
git push origin main
pause
cls
goto MENU

:OP2
echo.
echo Mise à jour des prix...
cd "Ma-collection"
python update_prices.py
cd /d "%~dp0"
pause
cls
goto MENU

:OP3
echo.
echo Synchronisation globale vers GitHub...
git add .
git commit -m "Mise a jour globale de la collection et de Chiner"
git pull --rebase origin main
git push origin main
echo.
echo Transfert terminé avec succès !
pause
cls
goto MENU

:FIN
exit