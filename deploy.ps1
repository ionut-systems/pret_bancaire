Write-Host "--- INITIALISATION DU DÉPLOIEMENT ---"

# 1. Nettoyage des anciennes compilations
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue

# 2. Compilation du Backend COBOL (Nouveau)
Write-Host "Compilation du moteur COBOL..."
cobc -x -free -o calc_pret.exe calc_pret.cbl

# 3. Compilation du Frontend en mode Répertoire
Write-Host "Compilation de l'interface Python..."
pyinstaller --noconsole --name="Terminal_Conseiller" app.py

# 4. Injection du Backend COBOL dans le livrable final
Write-Host "Intégration du moteur mathématique COBOL..."
Copy-Item "calc_pret.exe" -Destination "dist\Terminal_Conseiller\"

Write-Host "--- LIVRABLE TERMINÉ DANS LE DOSSIER DIST/ ---"