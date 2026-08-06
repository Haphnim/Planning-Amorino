# Planning Amorino Besançon, contexte pour Claude Code

Application Streamlit qui remplace le fichier Excel de planning de la boutique
Amorino Besançon (BU indépendante du contexte pro habituel de Kénim chez
Kubota, ce projet est un service pour l'équipe du point de vente).

## Démarrer

```bash
pip install -r requirements.txt
streamlit run app.py
```

`planning.db` est déjà rempli avec l'historique migré depuis
`PLANNING_AMORINO_BESANCON.xlsx` (5 semaines datées : 29/06, 20/07, 27/07,
03/08, 10/08/2026). Mot de passe responsable par défaut : `amorino2026`
(voir `DEFAULT_RESPONSABLE_PASSWORD` dans `app.py`).

## Comment vérifier visuellement pendant le développement

Ce projet a été développé dans un environnement sans navigateur graphique
persistant. Utilise Playwright pour vérifier chaque changement visuel avant
de le considérer terminé, ne pas se fier uniquement à la lecture du code :

```bash
pip install playwright
playwright install chromium --with-deps   # une seule fois

streamlit run app.py --server.port 8501 --server.headless true &
sleep 5
python3 -c "
from playwright.sync_api import sync_playwright
import time
with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page(viewport={'width': 1600, 'height': 1100})
    page.goto('http://127.0.0.1:8501', timeout=30000)
    page.wait_for_selector('[data-testid=\"stAppViewContainer\"]', timeout=20000)
    time.sleep(2)
    page.screenshot(path='screenshot.png', full_page=True)
    browser.close()
"
```

Puis ouvrir `screenshot.png`. Piège rencontré plusieurs fois pendant le dev :
lancer `streamlit run ... &` puis `sleep` **dans le même appel** peut faire
perdre le process en arrière-plan selon l'environnement, si ça arrive,
relancer le serveur dans un appel séparé de celui qui attend/teste.

## Bugs déjà trouvés et corrigés (pour référence, ne pas les réintroduire)

1. **HTML rendu en bloc de code brut au lieu d'être affiché.** Cause : dans
   du HTML multi-lignes passé à `st.markdown(unsafe_allow_html=True)`, toute
   ligne indentée de 4+ espaces (l'indentation naturelle du code Python à
   l'intérieur d'une fonction) est interprétée par Markdown comme un bloc de
   code littéral. Solution : la fonction `md_html()` dans `app.py` strip
   chaque ligne individuellement avant de rendre. **Toujours passer par
   `md_html()` plutôt que `st.markdown(..., unsafe_allow_html=True)`
   directement pour du HTML multi-lignes.** Un simple `textwrap.dedent` ne
   suffit pas : dès qu'une variable interpolée (ex: `LOGO_SVG`) a ses propres
   lignes à indentation 0, dedent ne trouve plus de préfixe commun et ne
   retire rien.

2. **`st.secrets.get(...)` fait planter l'app** (`StreamlitSecretNotFoundError`)
   quand aucun fichier `secrets.toml` n'existe du tout sur la machine, au
   lieu de simplement retourner la valeur par défaut comme un `dict.get()`
   normal. Solution : `obtenir_mot_de_passe_responsable()` dans `app.py`
   protège l'accès avec un try/except. Ne jamais appeler `st.secrets.get()`
   nu ailleurs dans le code sans la même protection.

3. **Créneau traversant minuit mal daté dans le calendrier.** Dans
   `construire_evenements_calendrier()`, comparer les heures de fin/début
   sur l'heure "horloge" brute (`hf*60+mf <= hd*60+md`) pour détecter le
   passage à minuit, pas sur l'offset décalé utilisé pour le calcul
   d'effectifs (`minutes_depuis_11h`), ces deux logiques ont des usages
   différents, ne pas les confondre.

4. **Migration Excel : nombre de colonnes/lignes horaires non uniforme
   selon les feuilles.** La feuille "du 03 au 09 -08" n'a que 27 créneaux
   réels au lieu de 29 sur les autres feuilles du même format ; la colonne
   juste après la dernière heure valide contient un total d'heures encodé
   en date Excel absurde (ex: `1900-05-29 02:30`) qui, lu à tort comme un
   créneau de travail, créait un faux mini-créneau isolé "00:30→01:00"
   chaque jour pour certains salariés (SUZE notamment), gonflant ses heures
   hebdo de façon irréaliste. `migrate_from_excel.py` détecte maintenant le
   nombre réel de colonnes/lignes horaires **dynamiquement par bloc** (en
   lisant la ligne/colonne d'en-tête elle-même) plutôt que d'utiliser
   `NB_SLOTS` en dur pour la lecture. Si tu retouches ce script, ne
   réintroduis pas un nombre de créneaux codé en dur pour la boucle de
   lecture des données (NB_SLOTS reste correct comme grille canonique côté
   `app.py`, c'est uniquement la lecture Excel qui doit rester dynamique).

## Anomalie de données connue (pas un bug de code, à signaler à l'utilisateur)

Même après le correctif ci-dessus, **SUZE fait 60,5h sur la semaine du
03-09/08/2026** dans les données réelles du fichier Excel source (7 services
consécutifs de 9 à 11h30 sans repos visible). C'est ce qui est vraiment écrit
dans le fichier d'origine, ce n'est pas un artefact de migration. À signaler
à l'utilisateur, ne pas "corriger" silencieusement une donnée réelle.

## Ce qui reste à vérifier / améliorer (pistes pour la suite)

- **Interaction de glisser-déposer réelle** sur la vue journée
  (`streamlit-calendar` / FullCalendar resourceTimeline) : l'affichage
  statique est vérifié (captures d'écran OK), mais l'interaction
  drag-and-drop elle-même (glisser un créneau, le déposer sur une autre
  ligne, callback `eventChange` → `modifier_creneau_complet`) n'a été
  vérifiée que par lecture de code, pas testée en conditions réelles de
  clic-glisser. À tester en priorité.
- **Réactivité mobile** : les salariés consulteront probablement leur
  planning sur téléphone. La vue salarié (cartes empilées) devrait bien se
  comporter, mais la vue journée du responsable (calendrier large) est
  pensée desktop, vérifier sur un viewport mobile et adapter si besoin
  (le responsable gère probablement depuis un ordinateur/tablette de toute
  façon, mais à confirmer avec l'utilisateur).
- **Licence FullCalendar Scheduler** : la vue journée utilise le plugin
  Scheduler (nécessaire pour l'affichage "un salarié par ligne"), activé en
  licence gratuite non-commerciale par `streamlit-calendar`. Voir
  https://fullcalendar.io/license si Amorino en fait un jour un usage
  dépassant le cadre interne.
- **Persistance SQLite sur Streamlit Community Cloud** : le fichier
  `planning.db` est perdu à chaque redémarrage du conteneur d'hébergement.
  Pour une vraie mise en prod, migrer vers Supabase (Postgres gratuit) en
  ne touchant que la section "COUCHE DONNEES" de `app.py`.
- **Objectifs d'effectif par défaut** posés à vue de nez (2 en creux, 3 aux
  heures de repas), à ajuster avec les vrais chiffres de rush du point de
  vente dans l'onglet "Réglages effectifs".
- **REPOS affiché comme une barre pleine journée** dans la vue calendrier
  (visuellement correct mais un peu envahissant sur la timeline), pourrait
  être affiné visuellement (badge compact plutôt que barre pleine largeur).

## Architecture

- `app.py`, application Streamlit complète (une seule page multi-écrans
  pilotée par `st.session_state`). Sections marquées par des commentaires
  `# ====...====` : STYLE GLOBAL, COUCHE DONNEES (SQLite), CALCULS METIER,
  EXPORT EXCEL, UI HELPERS, PAGES, POINT D'ENTREE.
- `migrate_from_excel.py`, script one-shot pour réimporter depuis le fichier
  Excel d'origine si besoin (idempotent, ne duplique pas les créneaux déjà en
  base).
- `planning.db`, base SQLite, déjà peuplée. Schéma : `salaries`, `creneaux`,
  `objectifs_effectif`.
- `requirements.txt`, dépendances (`streamlit`, `pandas`, `openpyxl`,
  `streamlit-calendar`).

## Style de code attendu (préférences de l'utilisateur)

- Code complet et prêt à l'emploi, jamais de snippets partiels.
- Toujours utiliser des virgules, jamais de tirets cadratins (,) dans le
  texte produit pour l'utilisateur (commentaires de code exemptés).
- Expliquer le "pourquoi" (cause racine d'un bug) plutôt qu'une solution
  magique sans explication, comme dans les sections ci-dessus.
