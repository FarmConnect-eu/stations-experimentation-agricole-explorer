# Carte des stations d'expérimentation agricole

Carte interactive et filtrable de 88 stations d'expérimentation agricole en France, consolidées depuis 5 réseaux :

- **Reseau Digifermes** (ACTA / Instituts techniques)
- **Reseau F@rmXP** (réseau national d'expérimentation)
- **Reseau IRFEL** (Institut de Recherche en Fruits Et Légumes)
- **Reseau Farm Demo Hub** (projet européen)
- **Reseau Chambres d'agriculture** (APCA)

## Démo en ligne

[Voir la carte interactive](https://USERNAME.github.io/REPO_NAME/)

> Remplace `USERNAME` et `REPO_NAME` une fois GitHub Pages activé.

## Fonctionnalités

- Carte de France avec marqueurs colorés par réseau (Leaflet + OpenStreetMap)
- Filtres multi-critères cumulables : réseaux, filières, régions, type de production
- Recherche libre (nom, espèce, sujet d'expérimentation)
- Recadrage automatique de la carte sur les stations filtrées
- Fiche détaillée au clic : contact, adresse, sujets d'expérimentation, site officiel
- Mode sombre automatique selon le système

## Déploiement sur GitHub Pages

1. **Créer un repo GitHub** (public ou privé selon ton plan)
2. **Pousser ces fichiers** à la racine du repo :
   ```bash
   git init
   git add .
   git commit -m "Carte stations experimentation agricole"
   git branch -M main
   git remote add origin https://github.com/USERNAME/REPO_NAME.git
   git push -u origin main
   ```
3. **Activer GitHub Pages** :
   - Settings → Pages
   - Source : `Deploy from a branch`
   - Branch : `main` / `/ (root)`
   - Save
4. **Attendre 1-2 minutes** que le déploiement se fasse
5. **Accéder à la page** : `https://USERNAME.github.io/REPO_NAME/`

Pour un domaine personnalisé (ex: `stations.exemple.fr`), ajouter un fichier `CNAME` à la racine contenant le domaine, puis configurer un enregistrement DNS CNAME vers `USERNAME.github.io`.

## Architecture du repo

```
.
├── index.html                          # Application autonome (HTML + CSS + JS + données inline)
├── data/
│   └── stations_experimentation_consolide.xlsx   # Source de données
├── scripts/
│   ├── build.py                        # Régénère index.html depuis le xlsx
│   ├── template.html                   # Template HTML avec placeholder __DATA__
│   └── update_data.py                  # Télécharge APCA et régénère le xlsx
├── PROMPT_MAJ.md                       # Prompt IA autonome pour régénération assistée
├── README.md
└── LICENSE
```

Le fichier `index.html` embarque toutes les données nécessaires : il fonctionne 100% côté client, aucun backend ni base de données requise. Le seul appel externe est aux tuiles OpenStreetMap (carte) et à la lib Leaflet via CDN (avec triple fallback : cdnjs → jsdelivr → unpkg).

## Mettre à jour les données

Deux options selon ton besoin.

### Option A - Script automatique (recommandé pour mise à jour de routine)

Le script `scripts/update_data.py` télécharge la dernière version de la carte APCA (Chambres d'agriculture) depuis Grist (data.gouv.fr ANCT), la fusionne avec les fermes complémentaires des autres réseaux (en dur dans le script, à mettre à jour si nouveaux ajouts) et régénère `data/stations_experimentation_consolide.xlsx`.

```bash
pip install openpyxl requests
python3 scripts/update_data.py
git add data/stations_experimentation_consolide.xlsx
git commit -m "Maj donnees stations"
git push
```

Le workflow GitHub Actions détecte le push, relance `scripts/build.py` automatiquement, régénère `index.html` et le commit. Le site est à jour en ~2 minutes.

### Option B - Régénération assistée par IA

Si la source APCA change de format, si tu veux ajouter un nouveau réseau de fermes, ou pour toute évolution structurelle : le fichier `PROMPT_MAJ.md` est un prompt complet et autonome à transmettre à n'importe quelle IA (Claude, ChatGPT, Mistral, Gemini). Il contient toutes les conventions, sources, et le schéma cible. Donne-le tel quel à l'IA en lui disant "exécute ce qui est demandé" ou "ajoute le réseau X".

## Sources

- **APCA / Chambres d'agriculture** : carte officielle des stations d'expérimentation (uMap ANCT, base Grist publique, février 2026)
- **Digifermes** : digifermes.com/blog/axes-de-travail/
- **F@rmXP** : farmxp.fr/nos-fermes/
- **IRFEL** : irfel.fr/fiche/
- **Farm Demo Hub** : farmdemo.eu

## Crédits

Consolidation et visualisation : contributeurs du repo  
Cartographie : © [OpenStreetMap](https://www.openstreetmap.org/copyright) contributors
Bibliothèque carto : [Leaflet](https://leafletjs.com/) 1.9.4 (BSD-2-Clause)

## Licence

Code : MIT (voir LICENSE)  
Données : voir conditions d'utilisation des sources (APCA, ACTA, IRFEL, FarmDemo, F@rmXP)
