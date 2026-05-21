# Prompt - Consolidation des stations d'experimentation agricole

Ce document est un **prompt complet et autonome** a transmettre tel quel a une IA (Claude, ChatGPT, Mistral, Gemini) pour relancer ou faire evoluer la consolidation de la base de donnees des stations d'experimentation agricole francaises.

---

## A copier-coller dans l'IA a partir d'ici

Je travaille sur une base de donnees consolidee des stations d'experimentation agricole francaises. J'aimerais que tu m'aides a la regenerer (ou la mettre a jour avec un nouveau reseau / une nouvelle source).

## Contexte

Je consolide une base de donnees ouverte des stations d'experimentation agricole francaises pour produire une cartographie unifiee. J'ai besoin d'une vue consolidee de **toutes les stations d'experimentation agricole en France**, croisant les principaux reseaux. Le livrable est un fichier Excel (.xlsx) avec un schema strict de 14 colonnes, et accessoirement une carte HTML interactive autonome.

## Schema de sortie (strictement obligatoire, ne pas modifier)

Le fichier xlsx final doit avoir exactement ces 14 colonnes dans cet ordre :

1. **Nom** (texte) - nom de la station
2. **Region** (texte) - region administrative francaise officielle
3. **Contact** (texte) - nom du contact principal, vide si non public
4. **Tel** (texte) - telephone, vide si non public
5. **Mail** (texte) - email, vide si non public
6. **Type de production** (texte) - une seule valeur parmi : `Conventionnel`, `Agriculture biologique`, `Mixte AB et conventionnel`
7. **Sujets d'experimentation** (texte) - liste separee par virgules
8. **Filiere** (texte) - une ou plusieurs filieres separees par virgule
9. **Especes etudiees** (texte) - liste separee par virgules
10. **Reseaux de fermes experimentales** (texte) - format obligatoire : `Reseau X` (ex: `Reseau Digifermes`). Multi-reseaux separes par un saut de ligne dans la cellule. Les stations sans appartenance a un reseau hors APCA = `Reseau Chambre d'agriculture`.
11. **Adresse postale** (texte) - adresse complete
12. **latitude** (nombre decimal, format 0.000000)
13. **longitude** (nombre decimal, format 0.000000)
14. **Site internet** (URL)

## Reseaux a integrer

| Reseau                          | Source publique                                                                                  | Volume approximatif |
|---------------------------------|--------------------------------------------------------------------------------------------------|---------------------|
| Reseau Chambre d'agriculture    | APCA Grist : https://grist.incubateur.anct.gouv.fr/api/docs/7zGLWG4fUHaJaNj9robyBS/download/xlsx?tableId=VF_Carto_Stations_Expes_25_02_2026 | 73 stations (source officielle, schema deja conforme) |
| Reseau Digifermes               | https://digifermes.com/blog/axes-de-travail/                                                     | 18 fiches           |
| Reseau F@rmXP                   | https://www.farmxp.fr/nos-fermes/                                                                | 8 fiches            |
| Reseau IRFEL                    | https://www.irfel.fr/stations/ (chaque station : https://www.irfel.fr/fiche/{slug}/)             | 17 fiches           |
| Reseau Farm Demo Hub            | https://farmdemo.eu/ - **agregateur europeen, ne pas creer de lignes propres** ; juste mentionner FARM DEMO HUB dans la col "Reseaux" pour les stations francaises qui y figurent | -                   |

## Methode de consolidation attendue

1. **Telecharger en priorite le fichier xlsx APCA** (lien direct ci-dessus). Il fait foi et son schema est deja parfaitement aligne sur les 14 colonnes. Ne PAS modifier le contenu APCA, le reprendre tel quel apres les 2 harmonisations decrites au point 4.

2. **Recuperer les fermes complementaires des 3 autres reseaux** (Digifermes, F@rmXP, IRFEL) qui ne sont PAS deja dans APCA. Source : pages publiques listees ci-dessus.

3. **Dedoublonner** : ~24 stations apparaissent dans plusieurs reseaux (ex: Trevarez dans F@rmXP + APCA, Mauron/CIRBEEF dans Digifermes + F@rmXP + APCA). Pour les doublons APCA, **toujours preferer la version APCA** (officielle, contacts complets, lat/lon precises). Pour les autres doublons, fusionner en accumulant les reseaux dans la cellule "Reseaux".

4. **Harmonisations strictes a appliquer** :
   - Type `AB` -> `Agriculture biologique` (jamais d'abreviation)
   - Reseau `Aucun` (utilise par APCA pour les stations Chambre d'agriculture sans autre affiliation) -> `Reseau Chambre d'agriculture`
   - Tirets em-dash `\u2014` et en-dash `\u2013` -> hyphen simple `-` partout
   - Espaces parasites en debut/fin de cellule supprimes
   - Coordonnees latitude/longitude converties en float numerique (pas de chaine)
   - Cellules vides reellement vides (pas la chaine "None" ni "-")

5. **Conventions de nommage** :
   - Format reseau : `Reseau Digifermes`, `Reseau F@rmXP`, `Reseau IRFEL`, `Reseau FARM DEMO HUB`, `Reseau Chambre d'agriculture` (toujours prefixe par "Reseau ")
   - Multi-reseaux : separer par un saut de ligne `\n` dans la cellule (rendu visible sur 2 lignes dans Excel)
   - Pas de point final dans les listes (sujets, especes)

6. **Contacts** : remplir Contact/Tel/Mail uniquement quand l'information est publique et sourcee sur le site officiel du reseau. **Ne jamais inventer**. Laisser vide sinon.

7. **Latitude/longitude** : utiliser l'API gratuite officielle BAN (Base Adresse Nationale) `https://api-adresse.data.gouv.fr/search/?q={adresse}` pour geocoder. Si l'adresse precise echoue, fallback sur le centroide de la commune. Precision attendue : adresse pour les stations connues, commune par defaut.

## Sortie attendue de l'IA

1. **Script Python autonome** qui :
   - Telecharge le xlsx APCA via requests
   - Genere les 14 colonnes harmonisees
   - Ajoute les fermes complementaires (a inclure en dur dans le script, schema identique)
   - Geocode via BAN si lat/lon manquantes
   - Ecrit le xlsx final avec mise en forme : police Arial, en-tete bleu `#305496` blanc bold, freeze panes A2, bordures, wrap text, format `0.000000` pour lat/lon
   - Affiche un recap final : nb de stations par reseau, par type de production

2. **Fichier `data/stations_experimentation_consolide.xlsx`** genere

3. **Statistiques de controle** :
   - Total attendu : ~85-90 lignes (selon l'etat courant des sources)
   - Repartition reseaux attendue (ordre de grandeur) : ~30 Chambre d'agriculture, ~15-20 Digifermes, ~15-17 IRFEL, ~10-15 FARM DEMO HUB, ~5-10 F@rmXP, ~5-8 multi-reseaux
   - Repartition types : majorite `Conventionnel` et `Mixte AB et conventionnel`, ~7-10 `Agriculture biologique`

## Preferences de redaction (important)

- Listes : tirets simples `-`, jamais em-dash ou en-dash
- Pas d'emoji dans le code ni les commentaires
- Commentaires Python en francais sans accents (compatibilite encoding)
- Variables et noms de fonctions en anglais standard

## Eventuelles evolutions

Si je te demande d'ajouter un nouveau reseau (ex: "ajoute le reseau Reseau Mixte Technologique X"), tu dois :
1. Trouver la page de listing publique du reseau
2. Recuperer chaque fiche station avec son nom, adresse, contact si public
3. Verifier les doublons avec les ~88 stations deja consolidees (matching par nom + code postal)
4. Pour les doublons : ajouter le nouveau reseau dans la cellule "Reseaux" existante (separateur saut de ligne)
5. Pour les nouvelles : ajouter une ligne complete au schema 14 colonnes
6. Regenerer le xlsx final et fournir un diff lisible

## Reference

Le repo de la version courante (remplace par l'URL de ton fork si different) :
https://github.com/USER/stations-experimentation-agricole-explorer

Le script Python existant qui fait deja tout ca :
`scripts/update_data.py` dans ce repo. **Lance-le en priorite si tu n'as aucune raison de tout refaire** :
```bash
pip install openpyxl requests
python3 scripts/update_data.py
```

Si quelque chose dans la source APCA a change (nouveau schema, nouvelles stations), tu peux soit adapter ce script, soit me proposer une nouvelle version.

---

## Fin du prompt

Voila. Donne ce document a une IA en lui disant simplement "execute ce qui est demande" ou "ajoute le reseau X a cette consolidation".
