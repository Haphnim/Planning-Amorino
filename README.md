# Planning Amorino Besançon

Application Streamlit qui remplace le fichier Excel de planning.

## Lancer en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

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
