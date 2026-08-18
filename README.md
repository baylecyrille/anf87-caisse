# ANF 87 — Caisse Buvette v2

## Fichiers PWA (GitHub → anf87-caisse)
- index.html
- manifest.json
- sw.js
- icon-192.png / icon-512.png

## Scripts Google Apps Script
- script_caisse.js      → Nouveau Google Sheet "ANF87 Caisse"
- script_boissons_fix.js → Remplace l'ancien script boissons offertes

## Installation
1. Google Sheet "ANF87 Caisse" → Extensions → Apps Script → coller script_caisse.js → Déployer
2. Remplacer le script boissons offertes par script_boissons_fix.js → Nouveau déploiement
3. GitHub repo anf87-caisse → uploader les 5 fichiers PWA → GitHub Pages
4. App → ⚙️ Config → saisir les 2 URLs → Initialiser les feuilles
