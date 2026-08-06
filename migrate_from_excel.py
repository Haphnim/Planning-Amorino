"""
Migration des données PLANNING_AMORINO_BESANCON.xlsx vers planning.db (SQLite).

Le fichier source contient 2 mises en page différentes :

  TYPE A (feuilles "29-06 au 05-07", "20 AU 26", "Feuil2") :
      - Date de la semaine en cellule B1 ("Semaine du : ... 2026")
      - Blocs de colonnes par jour : 1 colonne "horaires" (axe temps, ligne 4 à 32,
        11:00 à 01:00 par pas de 30 min) suivie des colonnes salariés
      - Une cellule remplie = le salarié travaille ce créneau de 30 min

  TYPE B (feuilles "du 27-07 au 02-08", "du 03 au 09 -08", "du 10 au 16-08") :
      - Date de la semaine dans le NOM de la feuille ("du DD-MM au DD-MM")
      - Blocs de LIGNES par jour : ligne d'en-tête JOURS/HORAIRES/<29 heures>,
        puis une ligne par salarié
      - Une cellule remplie = le salarié travaille ce créneau de 30 min

3 feuilles sont volontairement EXCLUES de l'import automatique :
  - "Feuil2" : couvre la même semaine (20-26 juillet 2026) que "20 AU 26" -> doublon
  - "Feuil1", "Feuil5", "Feuil6" : aucune date nulle part (ni cellule ni nom
    d'onglet), seulement 7 blocs de jours dans l'ordre Lundi->Dimanche.
    Impossible de les dater sans info supplémentaire.

Ces 3 feuilles sont quand même analysées et un résumé (salarié + jour de la
semaine + nombre de créneaux) est affiché en fin de script pour que tu puisses
me dire quelle semaine leur associer.

Usage :
    python migrate_from_excel.py chemin/vers/PLANNING_AMORINO_BESANCON.xlsx
"""

import sys
import re
import sqlite3
from datetime import datetime, date, timedelta, time as dtime

import openpyxl

DB_PATH = "planning.db"

MOIS_FR = {
    "janvier": 1, "février": 2, "fevrier": 2, "mars": 3, "avril": 4, "mai": 5,
    "juin": 6, "juillet": 7, "août": 8, "aout": 8, "septembre": 9,
    "octobre": 10, "novembre": 11, "décembre": 12, "decembre": 12,
}
JOURS_ORDRE = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI", "DIMANCHE"]

# Feuilles avec date fiable -> import automatique
FEUILLES_TYPE_A = ["29-06 au 05-07", "20 AU 26"]  # les deux ont "Semaine du" en B1
FEUILLES_TYPE_B = ["du 27-07 au 02-08", "du 03 au 09 -08", "du 10 au 16-08"]
# Feuilles exclues (doublon ou date inconnue), analysées mais pas importées
FEUILLES_EXCLUES = ["Feuil2", "Feuil1", "Feuil5", "Feuil6"]

NB_SLOTS = 29  # 11:00 -> 01:00 le lendemain, pas de 30 min


# ============================== OUTILS DATES ==============================

def parse_semaine_du(texte):
    """Gère les deux formats rencontrés dans le fichier :
    'Semaine du : 29 juin  au 05 juillet 2026' (mois répété)
    'Semaine du : 20 au 26 juillet 2026' (mois donné une seule fois, pour la fin)
    Dans les deux cas on ne dépend que de la date de FIN (dimanche), plus fiable,
    puis on recule de 6 jours pour obtenir le lundi.
    """
    m = re.search(r"au\s*(\d{1,2})\s+(\w+)\s+(\d{4})", texte, re.IGNORECASE)
    if not m:
        return None
    jour_fin, mois_fin_txt, annee = m.groups()
    mois_fin = MOIS_FR.get(mois_fin_txt.lower())
    if mois_fin is None:
        return None
    dimanche = date(int(annee), mois_fin, int(jour_fin))
    return dimanche - timedelta(days=6)


def parse_nom_feuille_type_b(nom_feuille, annee=2026):
    """'du 27-07 au 02-08' ou 'du 03 au 09 -08' -> date du lundi."""
    nom = nom_feuille.replace(" ", "")
    m = re.search(r"du(\d{1,2})(?:-(\d{1,2}))?au(\d{1,2})-(\d{1,2})", nom, re.IGNORECASE)
    if not m:
        return None
    jour_debut, mois_debut, jour_fin, mois_fin = m.groups()
    mois_fin = int(mois_fin)
    mois_debut = int(mois_debut) if mois_debut else mois_fin
    lundi = date(annee, mois_debut, int(jour_debut))
    return lundi


def slot_index_to_time(idx):
    """0 -> 11:00, 1 -> 11:30, ..., jusqu'à 01:00 (idx=28)."""
    total_minutes = 11 * 60 + idx * 30
    h = (total_minutes // 60) % 24
    m = total_minutes % 60
    return dtime(h, m)


def fusionner_creneaux(slots_travailles):
    """[0,1,2,5,6] -> [(0,3), (5,7)] c'est-à-dire liste de (debut_idx, fin_idx_exclu)."""
    if not slots_travailles:
        return []
    slots = sorted(slots_travailles)
    plages = []
    debut = slots[0]
    prec = slots[0]
    for s in slots[1:]:
        if s == prec + 1:
            prec = s
            continue
        plages.append((debut, prec + 1))
        debut = s
        prec = s
    plages.append((debut, prec + 1))
    return plages


# ============================== PARSEUR TYPE A ==============================

def parse_feuille_type_a(ws):
    """Retourne une liste de dicts {salarie, jour_semaine (0=lundi), plages:[(debut,fin)]}."""
    lundi = parse_semaine_du(ws.cell(row=1, column=2).value or "")
    if lundi is None:
        print(f"  ! Impossible de lire la date sur {ws.title}, feuille ignorée.")
        return None, []

    # Détecter les colonnes "axe temps" : cellule ligne 4 = 11:00 pile
    colonnes_temps = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=4, column=c).value
        h = getattr(v, "hour", None)
        mi = getattr(v, "minute", None)
        if h == 11 and mi == 0:
            colonnes_temps.append(c)

    resultats = []
    for i, col_temps in enumerate(colonnes_temps):
        col_fin_bloc = colonnes_temps[i + 1] if i + 1 < len(colonnes_temps) else ws.max_column + 1
        colonnes_salaries = [c for c in range(col_temps + 1, col_fin_bloc)]

        for col in colonnes_salaries:
            nom = ws.cell(row=3, column=col).value
            if not nom or str(nom).strip().upper() in ("", "HORAIRES"):
                continue
            nom = str(nom).strip().upper()

            slots = []
            for r in range(4, 4 + NB_SLOTS):
                if ws.cell(row=r, column=col).value is not None:
                    slots.append(r - 4)

            plages = fusionner_creneaux(slots)
            resultats.append({"salarie": nom, "jour_semaine": i, "plages": plages})

    return lundi, resultats


# ============================== PARSEUR TYPE B ==============================

def parse_feuille_type_b(ws):
    lundi = parse_nom_feuille_type_b(ws.title)
    if lundi is None:
        print(f"  ! Impossible de lire la date sur {ws.title}, feuille ignorée.")
        return None, []

    lignes_entete = [r for r in range(1, ws.max_row + 1) if ws.cell(row=r, column=1).value == "JOURS"]

    resultats = []
    for i, r_entete in enumerate(lignes_entete):
        r_fin_bloc = lignes_entete[i + 1] if i + 1 < len(lignes_entete) else ws.max_row + 1

        for r in range(r_entete + 1, r_fin_bloc):
            nom = ws.cell(row=r, column=2).value
            if not nom:
                continue
            nom = str(nom).strip().upper()
            if nom in ("HORAIRES", "JOURS"):
                continue

            slots = []
            for idx in range(NB_SLOTS):
                col = 3 + idx  # colonne C = premier créneau
                if ws.cell(row=r, column=col).value is not None:
                    slots.append(idx)

            plages = fusionner_creneaux(slots)
            resultats.append({"salarie": nom, "jour_semaine": i, "plages": plages})

    return lundi, resultats


# ============================== ANALYSE DES FEUILLES EXCLUES ==============================

def analyser_feuille_exclue(ws):
    """Renvoie un résumé texte (salarié / jour / nb créneaux) sans rien importer,
    pour que l'utilisateur puisse identifier la semaine concernée."""
    lignes_entete = [r for r in range(1, ws.max_row + 1) if ws.cell(row=r, column=1).value == "JOURS"]
    if not lignes_entete:
        return []
    resume = []
    for i, r_entete in enumerate(lignes_entete):
        r_fin_bloc = lignes_entete[i + 1] if i + 1 < len(lignes_entete) else ws.max_row + 1
        jour = JOURS_ORDRE[i] if i < len(JOURS_ORDRE) else f"Bloc {i}"
        for r in range(r_entete + 1, r_fin_bloc):
            nom = ws.cell(row=r, column=2).value
            if not nom or str(nom).strip().upper() in ("HORAIRES", "JOURS"):
                continue
            nb_slots = sum(
                1 for idx in range(NB_SLOTS)
                if ws.cell(row=r, column=3 + idx).value is not None
            )
            if nb_slots > 0:
                resume.append((jour, str(nom).strip().upper(), nb_slots))
    return resume


# ============================== ECRITURE EN BASE ==============================

def get_or_create_salarie(conn, nom):
    row = conn.execute("SELECT id FROM salaries WHERE nom = ?", (nom,)).fetchone()
    if row:
        return row[0]
    cur = conn.execute("INSERT INTO salaries (nom, actif) VALUES (?, 1)", (nom,))
    return cur.lastrowid


def importer_resultats(conn, lundi, resultats, source_feuille):
    n_travail, n_repos = 0, 0
    jours_avec_salaries = set()

    for entree in resultats:
        salarie_id = get_or_create_salarie(conn, entree["salarie"])
        jour = lundi + timedelta(days=entree["jour_semaine"])
        jours_avec_salaries.add((salarie_id, entree["jour_semaine"]))

        if not entree["plages"]:
            # Aucun créneau travaillé ce jour-là pour ce salarié -> repos présumé
            existe = conn.execute(
                "SELECT 1 FROM creneaux WHERE salarie_id=? AND jour=? AND type='REPOS'",
                (salarie_id, jour.isoformat()),
            ).fetchone()
            if not existe:
                conn.execute(
                    "INSERT INTO creneaux (salarie_id, jour, type, commentaire) VALUES (?, ?, 'REPOS', ?)",
                    (salarie_id, jour.isoformat(), f"Importé de {source_feuille}"),
                )
                n_repos += 1
            continue

        for debut_idx, fin_idx in entree["plages"]:
            heure_debut = slot_index_to_time(debut_idx)
            heure_fin = slot_index_to_time(fin_idx)  # fin exclue = fin du dernier créneau + 30min
            existe = conn.execute(
                """SELECT 1 FROM creneaux WHERE salarie_id=? AND jour=? AND type='TRAVAIL'
                   AND heure_debut=? AND heure_fin=?""",
                (salarie_id, jour.isoformat(), heure_debut.strftime("%H:%M"), heure_fin.strftime("%H:%M")),
            ).fetchone()
            if existe:
                continue
            conn.execute(
                """INSERT INTO creneaux (salarie_id, jour, type, heure_debut, heure_fin, commentaire)
                   VALUES (?, ?, 'TRAVAIL', ?, ?, ?)""",
                (salarie_id, jour.isoformat(), heure_debut.strftime("%H:%M"),
                 heure_fin.strftime("%H:%M"), f"Importé de {source_feuille}"),
            )
            n_travail += 1

    return n_travail, n_repos


# ============================== POINT D'ENTREE ==============================

def main():
    if len(sys.argv) < 2:
        print("Usage : python migrate_from_excel.py chemin/vers/fichier.xlsx")
        sys.exit(1)

    chemin_excel = sys.argv[1]
    wb = openpyxl.load_workbook(chemin_excel, data_only=True)

    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS salaries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nom TEXT NOT NULL UNIQUE,
            actif INTEGER NOT NULL DEFAULT 1,
            heures_contrat REAL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS creneaux (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            salarie_id INTEGER NOT NULL REFERENCES salaries(id) ON DELETE CASCADE,
            jour DATE NOT NULL,
            type TEXT NOT NULL DEFAULT 'TRAVAIL',
            heure_debut TEXT,
            heure_fin TEXT,
            commentaire TEXT
        )
    """)

    print("=" * 60)
    print("IMPORT DES FEUILLES DATEES")
    print("=" * 60)

    for nom_feuille in FEUILLES_TYPE_A:
        if nom_feuille not in wb.sheetnames:
            continue
        ws = wb[nom_feuille]
        lundi, resultats = parse_feuille_type_a(ws)
        if lundi is None:
            continue
        n_t, n_r = importer_resultats(conn, lundi, resultats, nom_feuille)
        print(f"[{nom_feuille}] semaine du {lundi.strftime('%d/%m/%Y')} "
              f"-> {n_t} créneaux TRAVAIL, {n_r} jours REPOS importés")

    for nom_feuille in FEUILLES_TYPE_B:
        if nom_feuille not in wb.sheetnames:
            continue
        ws = wb[nom_feuille]
        lundi, resultats = parse_feuille_type_b(ws)
        if lundi is None:
            continue
        n_t, n_r = importer_resultats(conn, lundi, resultats, nom_feuille)
        print(f"[{nom_feuille}] semaine du {lundi.strftime('%d/%m/%Y')} "
              f"-> {n_t} créneaux TRAVAIL, {n_r} jours REPOS importés")

    conn.commit()

    print()
    print("=" * 60)
    print("FEUILLES EXCLUES (nécessitent ta confirmation, rien importé)")
    print("=" * 60)
    for nom_feuille in FEUILLES_EXCLUES:
        if nom_feuille not in wb.sheetnames:
            continue
        ws = wb[nom_feuille]
        resume = analyser_feuille_exclue(ws)
        print(f"\n--- {nom_feuille} ---")
        if nom_feuille == "Feuil2":
            print("  Contient une date explicite : 'Semaine du 20 au 26 juillet 2026'")
            print("  -> Identique à la semaine de la feuille '20 AU 26'. Doublon probable.")
        if not resume:
            print("  (vide ou illisible)")
            continue
        par_jour = {}
        for jour, nom, nb in resume:
            par_jour.setdefault(jour, []).append(f"{nom} ({nb} créneaux)")
        for jour in JOURS_ORDRE:
            if jour in par_jour:
                print(f"  {jour}: {', '.join(par_jour[jour])}")

    conn.close()
    print()
    print(f"Base mise à jour : {DB_PATH}")
    print("Relance l'appli Streamlit, les vues semaine ne seront plus vides")
    print("pour les 5 semaines datées (29/06, 20/07, 27/07, 03/08, 10/08).")


if __name__ == "__main__":
    main()
