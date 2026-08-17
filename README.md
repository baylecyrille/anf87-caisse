# ANF 87 — Caisse Buvette

PWA de caisse pour les buvettes de l'Avenir Nord Foot 87.

## Fichiers
- `index.html` — Application principale
- `sw.js` — Service Worker (cache offline)
- `manifest.json` — Config PWA
- `icon-192.png` / `icon-512.png` — Icônes
- `script_caisse.js` — Google Apps Script à coller dans Apps Script

## Installation rapide
1. Créer un nouveau Google Sheet "ANF87 Caisse"
2. Extensions → Apps Script → coller `script_caisse.js`
3. Déployer → Nouveau déploiement → Application Web → Tout le monde
4. Créer repo GitHub `anf87-caisse` → uploader les 6 fichiers
5. Settings → Pages → main / root → Save
6. Ouvrir `https://VOTRE-NOM.github.io/anf87-caisse/`
7. Onglet Config → coller l'URL Apps Script → Initialiser les feuilles
