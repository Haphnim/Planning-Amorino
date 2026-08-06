"""
Planning Amorino Besançon, application Streamlit.

Remplace le fichier Excel PLANNING_AMORINO_BESANCON.xlsx par une application web
avec 3 niveaux d'accès :
  - Salarié : consultation de son propre planning uniquement
  - Responsable : création/modification/suppression des créneaux, vue globale
  - Vous (super admin) : responsable + gestion des salariés et export

Stockage : SQLite (fichier planning.db), suffisant pour ce volume de données.
Si l'app est déployée sur Streamlit Community Cloud, le fichier SQLite est
réinitialisé à chaque redéploiement/restart du conteneur. Pour une persistance
garantie en production, migrer vers une base externe (Supabase/Postgres gratuit)
en remplaçant uniquement les fonctions de la section "COUCHE DONNEES" ci-dessous.
"""

import sqlite3
import io
from datetime import datetime, date, timedelta, time as dtime
from contextlib import contextmanager

import streamlit as st
import pandas as pd

DB_PATH = "planning.db"
TYPES_ABSENCE = ["TRAVAIL", "REPOS", "MALADIE", "POSE", "FORMATION"]
COULEURS_TYPE = {
    "TRAVAIL": "#2E86AB",
    "REPOS": "#B0B0B0",
    "MALADIE": "#E76F51",
    "POSE": "#F4A261",
    "FORMATION": "#8E44AD",
}
JOURS_FR = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]

# Mot de passe responsable, à changer avant mise en prod réelle.
# Idéalement stocké dans .streamlit/secrets.toml sous la clé RESPONSABLE_PASSWORD.
DEFAULT_RESPONSABLE_PASSWORD = "amorino2026"


# ============================== COUCHE DONNEES ==============================

@contextmanager
def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db():
    with get_conn() as conn:
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
        # Salariés observés dans le fichier Excel d'origine, pré-remplissage pratique
        noms_initiaux = ["ANAE", "CLOE", "LOUISE", "JEREMY", "LUCAS", "CELINE",
                          "LILOU", "ROMANE", "SUZE", "SINEM"]
        for nom in noms_initiaux:
            conn.execute("INSERT OR IGNORE INTO salaries (nom) VALUES (?)", (nom,))


def liste_salaries(actifs_seulement=True):
    with get_conn() as conn:
        q = "SELECT * FROM salaries"
        if actifs_seulement:
            q += " WHERE actif = 1"
        q += " ORDER BY nom"
        return pd.DataFrame([dict(r) for r in conn.execute(q).fetchall()])


def ajouter_salarie(nom, heures_contrat=None):
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO salaries (nom, heures_contrat) VALUES (?, ?)",
            (nom.strip().upper(), heures_contrat),
        )


def desactiver_salarie(salarie_id):
    with get_conn() as conn:
        conn.execute("UPDATE salaries SET actif = 0 WHERE id = ?", (salarie_id,))


def reactiver_salarie(salarie_id):
    with get_conn() as conn:
        conn.execute("UPDATE salaries SET actif = 1 WHERE id = ?", (salarie_id,))


def ajouter_creneau(salarie_id, jour, type_, heure_debut, heure_fin, commentaire=""):
    with get_conn() as conn:
        conn.execute(
            """INSERT INTO creneaux (salarie_id, jour, type, heure_debut, heure_fin, commentaire)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (salarie_id, jour.isoformat(), type_,
             heure_debut.strftime("%H:%M") if heure_debut else None,
             heure_fin.strftime("%H:%M") if heure_fin else None,
             commentaire),
        )


def modifier_creneau(creneau_id, type_, heure_debut, heure_fin, commentaire=""):
    with get_conn() as conn:
        conn.execute(
            """UPDATE creneaux SET type = ?, heure_debut = ?, heure_fin = ?, commentaire = ?
               WHERE id = ?""",
            (type_,
             heure_debut.strftime("%H:%M") if heure_debut else None,
             heure_fin.strftime("%H:%M") if heure_fin else None,
             commentaire, creneau_id),
        )


def supprimer_creneau(creneau_id):
    with get_conn() as conn:
        conn.execute("DELETE FROM creneaux WHERE id = ?", (creneau_id,))


def creneaux_semaine(lundi: date):
    dimanche = lundi + timedelta(days=6)
    with get_conn() as conn:
        rows = conn.execute(
            """SELECT c.*, s.nom AS salarie_nom
               FROM creneaux c JOIN salaries s ON s.id = c.salarie_id
               WHERE c.jour BETWEEN ? AND ?
               ORDER BY c.jour, s.nom, c.heure_debut""",
            (lundi.isoformat(), dimanche.isoformat()),
        ).fetchall()
        return pd.DataFrame([dict(r) for r in rows])


def creneaux_salarie(salarie_id, lundi: date):
    dimanche = lundi + timedelta(days=6)
    with get_conn() as conn:
        rows = conn.execute(
            """SELECT c.*, s.nom AS salarie_nom
               FROM creneaux c JOIN salaries s ON s.id = c.salarie_id
               WHERE c.salarie_id = ? AND c.jour BETWEEN ? AND ?
               ORDER BY c.jour, c.heure_debut""",
            (salarie_id, lundi.isoformat(), dimanche.isoformat()),
        ).fetchall()
        return pd.DataFrame([dict(r) for r in rows])


# ============================== CALCULS METIER ==============================

def duree_heures(heure_debut, heure_fin):
    """Retourne la durée en heures décimales entre deux horaires 'HH:MM'."""
    if not heure_debut or not heure_fin:
        return 0.0
    fmt = "%H:%M"
    t1 = datetime.strptime(heure_debut, fmt)
    t2 = datetime.strptime(heure_fin, fmt)
    delta = (t2 - t1).total_seconds() / 3600
    if delta < 0:
        delta += 24  # créneau passant minuit
    return round(delta, 2)


def total_heures_par_salarie(df_creneaux):
    if df_creneaux.empty:
        return pd.DataFrame(columns=["salarie_nom", "heures_travaillees", "jours_repos"])
    df = df_creneaux.copy()
    df["heures"] = df.apply(
        lambda r: duree_heures(r["heure_debut"], r["heure_fin"]) if r["type"] == "TRAVAIL" else 0.0,
        axis=1,
    )
    resume = df.groupby("salarie_nom").agg(
        heures_travaillees=("heures", "sum"),
        jours_repos=("type", lambda s: (s == "REPOS").sum()),
    ).reset_index()
    return resume


def alertes_repos(df_creneaux, seuil_jours_consecutifs=6):
    """Signale les salariés sans jour REPOS sur la semaine (obligation légale)."""
    if df_creneaux.empty:
        return []
    alertes = []
    for nom, grp in df_creneaux.groupby("salarie_nom"):
        if not (grp["type"] == "REPOS").any():
            alertes.append(f"{nom} n'a aucun jour de repos enregistré cette semaine.")
    return alertes


# ============================== EXPORT EXCEL ==============================

def export_excel(df_semaine: pd.DataFrame, lundi: date) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        if df_semaine.empty:
            pd.DataFrame({"info": ["Aucun créneau sur cette semaine"]}).to_excel(
                writer, index=False, sheet_name="Planning"
            )
        else:
            export = df_semaine[["jour", "salarie_nom", "type", "heure_debut", "heure_fin", "commentaire"]].copy()
            export.columns = ["Jour", "Salarié", "Type", "Début", "Fin", "Commentaire"]
            export.to_excel(writer, index=False, sheet_name="Planning")
            resume = total_heures_par_salarie(df_semaine)
            resume.columns = ["Salarié", "Heures travaillées", "Jours de repos"]
            resume.to_excel(writer, index=False, sheet_name="Résumé heures")
    buffer.seek(0)
    return buffer.read()


# ============================== UI HELPERS ==============================

def lundi_de_la_semaine(d: date) -> date:
    return d - timedelta(days=d.weekday())


def selecteur_semaine(key_prefix):
    if f"{key_prefix}_lundi" not in st.session_state:
        st.session_state[f"{key_prefix}_lundi"] = lundi_de_la_semaine(date.today())

    c1, c2, c3, c4 = st.columns([1, 1, 3, 1])
    with c1:
        if st.button("◀ Semaine préc.", key=f"{key_prefix}_prev"):
            st.session_state[f"{key_prefix}_lundi"] -= timedelta(days=7)
    with c2:
        if st.button("Semaine suiv. ▶", key=f"{key_prefix}_next"):
            st.session_state[f"{key_prefix}_lundi"] += timedelta(days=7)
    with c4:
        if st.button("Aujourd'hui", key=f"{key_prefix}_today"):
            st.session_state[f"{key_prefix}_lundi"] = lundi_de_la_semaine(date.today())

    lundi = st.session_state[f"{key_prefix}_lundi"]
    dimanche = lundi + timedelta(days=6)
    with c3:
        st.markdown(
            f"### Semaine du {lundi.strftime('%d/%m/%Y')} au {dimanche.strftime('%d/%m/%Y')}"
        )
    return lundi


def afficher_grille_semaine(lundi, df_creneaux, editable_ids=None):
    """Affiche la grille jour x salarié avec les créneaux, dans un style proche
    du planning Excel d'origine mais lisible."""
    dimanche = lundi + timedelta(days=6)
    jours = [lundi + timedelta(days=i) for i in range(7)]

    if df_creneaux.empty:
        st.info("Aucun créneau enregistré pour cette semaine.")
        return

    for i, jour in enumerate(jours):
        jour_creneaux = df_creneaux[df_creneaux["jour"] == jour.isoformat()]
        if jour_creneaux.empty:
            continue
        st.markdown(f"**{JOURS_FR[i]} {jour.strftime('%d/%m')}**")
        cols = st.columns(len(jour_creneaux)) if len(jour_creneaux) <= 6 else st.columns(6)
        for idx, (_, row) in enumerate(jour_creneaux.iterrows()):
            col = cols[idx % len(cols)]
            couleur = COULEURS_TYPE.get(row["type"], "#999999")
            horaire = (f"{row['heure_debut']} - {row['heure_fin']}"
                       if row["type"] == "TRAVAIL" and row["heure_debut"] else row["type"])
            with col:
                st.markdown(
                    f"""<div style="background:{couleur}22;border-left:4px solid {couleur};
                    padding:6px 10px;border-radius:4px;margin-bottom:6px;">
                    <b>{row['salarie_nom']}</b><br><span style="font-size:0.85em">{horaire}</span>
                    </div>""",
                    unsafe_allow_html=True,
                )


# ============================== PAGES ==============================

def page_connexion():
    st.title("🍦 Planning Amorino Besançon")
    st.write("Sélectionnez votre profil pour continuer.")

    role = st.radio("Profil", ["Salarié", "Responsable"], horizontal=True)

    if role == "Salarié":
        df_sal = liste_salaries()
        if df_sal.empty:
            st.warning("Aucun salarié enregistré pour le moment.")
            return
        nom = st.selectbox("Votre nom", df_sal["nom"].tolist())
        if st.button("Voir mon planning"):
            st.session_state["role"] = "salarie"
            st.session_state["salarie_id"] = int(df_sal[df_sal["nom"] == nom]["id"].iloc[0])
            st.session_state["salarie_nom"] = nom
            st.rerun()
    else:
        mdp = st.text_input("Mot de passe responsable", type="password")
        vrai_mdp = st.secrets.get("RESPONSABLE_PASSWORD", DEFAULT_RESPONSABLE_PASSWORD)
        if st.button("Connexion"):
            if mdp == vrai_mdp:
                st.session_state["role"] = "responsable"
                st.rerun()
            else:
                st.error("Mot de passe incorrect.")


def page_salarie():
    st.title(f"🍦 Mon planning — {st.session_state['salarie_nom']}")
    if st.sidebar.button("← Se déconnecter"):
        for k in ["role", "salarie_id", "salarie_nom"]:
            st.session_state.pop(k, None)
        st.rerun()

    lundi = selecteur_semaine("salarie")
    df = creneaux_salarie(st.session_state["salarie_id"], lundi)

    afficher_grille_semaine(lundi, df)

    if not df.empty:
        resume = total_heures_par_salarie(df)
        if not resume.empty:
            st.metric("Heures travaillées cette semaine", f"{resume['heures_travaillees'].iloc[0]:.1f} h")


def page_responsable():
    st.title("🍦 Planning Amorino Besançon — Gestion")
    if st.sidebar.button("← Se déconnecter"):
        st.session_state.pop("role", None)
        st.rerun()

    onglet = st.sidebar.radio(
        "Navigation", ["Vue semaine", "Ajouter un créneau", "Modifier / supprimer", "Salariés", "Export"]
    )

    if onglet == "Vue semaine":
        lundi = selecteur_semaine("resp")
        df = creneaux_semaine(lundi)
        for a in alertes_repos(df):
            st.warning(f"⚠️ {a}")
        afficher_grille_semaine(lundi, df)
        if not df.empty:
            st.markdown("---")
            st.subheader("Total heures par salarié")
            st.dataframe(total_heures_par_salarie(df), use_container_width=True, hide_index=True)

    elif onglet == "Ajouter un créneau":
        st.subheader("Ajouter un créneau")
        df_sal = liste_salaries()
        if df_sal.empty:
            st.warning("Ajoutez d'abord un salarié dans l'onglet Salariés.")
            return
        with st.form("form_ajout"):
            nom = st.selectbox("Salarié", df_sal["nom"].tolist())
            jour = st.date_input("Jour", value=date.today())
            type_ = st.selectbox("Type", TYPES_ABSENCE)
            c1, c2 = st.columns(2)
            with c1:
                heure_debut = st.time_input("Heure de début", value=dtime(11, 0)) if type_ == "TRAVAIL" else None
            with c2:
                heure_fin = st.time_input("Heure de fin", value=dtime(18, 0)) if type_ == "TRAVAIL" else None
            commentaire = st.text_input("Commentaire (optionnel)")
            submit = st.form_submit_button("Ajouter")
            if submit:
                salarie_id = int(df_sal[df_sal["nom"] == nom]["id"].iloc[0])
                ajouter_creneau(salarie_id, jour, type_, heure_debut, heure_fin, commentaire)
                st.success(f"Créneau ajouté pour {nom} le {jour.strftime('%d/%m/%Y')}.")
                st.rerun()

    elif onglet == "Modifier / supprimer":
        st.subheader("Modifier ou supprimer un créneau")
        lundi = selecteur_semaine("edit")
        df = creneaux_semaine(lundi)
        if df.empty:
            st.info("Aucun créneau cette semaine.")
            return
        df_aff = df.copy()
        df_aff["label"] = df_aff.apply(
            lambda r: f"{r['jour']} — {r['salarie_nom']} — {r['type']}"
                      + (f" ({r['heure_debut']}-{r['heure_fin']})" if r['type'] == 'TRAVAIL' else ""),
            axis=1,
        )
        choix = st.selectbox("Créneau", df_aff["label"].tolist())
        row = df_aff[df_aff["label"] == choix].iloc[0]

        with st.form("form_edit"):
            type_ = st.selectbox("Type", TYPES_ABSENCE, index=TYPES_ABSENCE.index(row["type"]))
            c1, c2 = st.columns(2)
            hd_defaut = datetime.strptime(row["heure_debut"], "%H:%M").time() if row["heure_debut"] else dtime(11, 0)
            hf_defaut = datetime.strptime(row["heure_fin"], "%H:%M").time() if row["heure_fin"] else dtime(18, 0)
            with c1:
                heure_debut = st.time_input("Heure de début", value=hd_defaut) if type_ == "TRAVAIL" else None
            with c2:
                heure_fin = st.time_input("Heure de fin", value=hf_defaut) if type_ == "TRAVAIL" else None
            commentaire = st.text_input("Commentaire", value=row["commentaire"] or "")
            c3, c4 = st.columns(2)
            with c3:
                maj = st.form_submit_button("Enregistrer les modifications")
            with c4:
                suppr = st.form_submit_button("🗑️ Supprimer ce créneau")
            if maj:
                modifier_creneau(int(row["id"]), type_, heure_debut, heure_fin, commentaire)
                st.success("Créneau modifié.")
                st.rerun()
            if suppr:
                supprimer_creneau(int(row["id"]))
                st.success("Créneau supprimé.")
                st.rerun()

    elif onglet == "Salariés":
        st.subheader("Gestion des salariés")
        with st.form("form_nouveau_salarie"):
            nom = st.text_input("Nom du nouveau salarié")
            heures = st.number_input("Heures contractuelles / semaine (optionnel)", min_value=0.0, step=0.5)
            if st.form_submit_button("Ajouter le salarié"):
                if nom.strip():
                    try:
                        ajouter_salarie(nom, heures or None)
                        st.success(f"{nom.strip().upper()} ajouté(e).")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Ce nom existe déjà.")
                else:
                    st.error("Le nom ne peut pas être vide.")

        st.markdown("---")
        st.write("**Salariés actifs**")
        df_actifs = liste_salaries(actifs_seulement=True)
        for _, r in df_actifs.iterrows():
            c1, c2 = st.columns([4, 1])
            c1.write(r["nom"])
            if c2.button("Désactiver", key=f"desact_{r['id']}"):
                desactiver_salarie(int(r["id"]))
                st.rerun()

        df_inactifs = liste_salaries(actifs_seulement=False)
        df_inactifs = df_inactifs[df_inactifs["actif"] == 0]
        if not df_inactifs.empty:
            st.write("**Salariés désactivés**")
            for _, r in df_inactifs.iterrows():
                c1, c2 = st.columns([4, 1])
                c1.write(r["nom"])
                if c2.button("Réactiver", key=f"react_{r['id']}"):
                    reactiver_salarie(int(r["id"]))
                    st.rerun()

    elif onglet == "Export":
        st.subheader("Export Excel de la semaine")
        lundi = selecteur_semaine("export")
        df = creneaux_semaine(lundi)
        excel_bytes = export_excel(df, lundi)
        st.download_button(
            "📥 Télécharger le planning de la semaine (.xlsx)",
            data=excel_bytes,
            file_name=f"planning_{lundi.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ============================== POINT D'ENTREE ==============================

def main():
    st.set_page_config(page_title="Planning Amorino Besançon", page_icon="🍦", layout="wide")
    init_db()

    if "role" not in st.session_state:
        page_connexion()
    elif st.session_state["role"] == "salarie":
        page_salarie()
    elif st.session_state["role"] == "responsable":
        page_responsable()


if __name__ == "__main__":
    main()
