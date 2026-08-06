# Planning Amorino Besançon

Application Streamlit qui remplace le fichier Excel de planning.

## Lancer en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

`planning.db` est déjà fourni pré-rempli avec l'historique importé depuis
`PLANNING_AMORINO_BESANCON.xlsx` (voir section Migration ci-dessous). Place-le
simplement à côté de `app.py` avant de lancer l'appli, ou relance la migration
toi-même.

## Migration depuis l'Excel

```bash
python migrate_from_excel.py chemin/vers/PLANNING_AMORINO_BESANCON.xlsx
```

5 semaines ont été importées automatiquement (dates fiables trouvées dans le
fichier) : **29/06, 20/07, 27/07, 03/08, 10/08/2026**.

3 feuilles ont été volontairement exclues, à traiter à la main si besoin :
- **Feuil2** : couvre la même semaine que "20 AU 26" (20-26 juillet), doublon
  probable, brouillon non retenu.
- **Feuil1, Feuil5, Feuil6** : aucune date nulle part dans la feuille, et tous
  les créneaux sont remplis en continu de 11h à 01h pour chaque salarié
  (amplitude de 18h), ce qui n'est pas un planning réel. Ce sont très
  probablement des **gabarits vierges** copiés-collés comme base de départ,
  pas des semaines travaillées. Si l'une d'elles correspond en fait à une
  vraie semaine, dis-moi laquelle et je fais l'import manuel correspondant.

Le script est idempotent : le relancer sur la même base ne duplique pas les
créneaux déjà importés (vérification par salarié/jour/type/horaires avant
insertion).

## Déploiement sur Streamlit Community Cloud (déjà utilisé pour DAX/SAP Toolkit BU Parts)

1. Pousser ce dossier (`app.py`, `requirements.txt`) sur un repo GitHub.
2. Sur share.streamlit.io, "New app" en pointant vers ce repo, fichier `app.py`.
3. Dans les Settings de l'app > Secrets, ajouter :
   ```toml
   RESPONSABLE_PASSWORD = "votre_mot_de_passe_ici"
   ```
   Sans ça, l'app utilise le mot de passe par défaut défini dans le code
   (`DEFAULT_RESPONSABLE_PASSWORD`), à changer avant tout partage réel.

## Limite importante : persistance des données

La base est un fichier SQLite (`planning.db`) créé automatiquement au premier
lancement. Sur Streamlit Community Cloud, ce fichier est perdu à chaque
redémarrage du conteneur (mise en veille après inactivité, redeploy, etc.).

Deux options pour une vraie persistance en production :
1. **Supabase** (Postgres gratuit) : remplacer les fonctions de la section
   "COUCHE DONNEES" dans `app.py` par des appels psycopg2/SQLAlchemy. Le reste
   de l'app (UI, calculs) ne change pas.
2. **Export régulier** : utiliser le bouton Export Excel comme sauvegarde
   manuelle en attendant une vraie base externe.

## Rôles

- **Salarié** : sélectionne son nom dans une liste, consulte uniquement son
  planning (lecture seule).
- **Responsable** : mot de passe requis, accès complet (ajout/modification/
  suppression de créneaux, gestion des salariés, export Excel).

## Types de créneaux

TRAVAIL (avec heure début/fin), REPOS, MALADIE, POSE, FORMATION.
Un salarié sans REPOS enregistré sur la semaine déclenche une alerte visuelle
dans la vue responsable (obligation légale de repos hebdomadaire).
