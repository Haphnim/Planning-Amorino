"""
Planning Amorino Besançon, application Streamlit (v2 — vue journée + effectifs + drag & drop).

3 niveaux d'accès :
  - Salarié   : son planning en 2 clics (lecture seule)
  - Responsable : vue journée façon planning mural, drag & drop des créneaux,
                  jauge d'effectifs par tranche horaire (repérer les trous en
                  heures de rush), gestion salariés, export Excel

Stockage : SQLite (planning.db). Voir README pour la limite de persistance
sur Streamlit Community Cloud et la migration vers une base externe.
"""

import sqlite3
import io
from datetime import datetime, date, timedelta, time as dtime
from contextlib import contextmanager

import streamlit as st
import pandas as pd
from streamlit_calendar import calendar

DB_PATH = "planning.db"
TYPES_ABSENCE = ["TRAVAIL", "REPOS", "MALADIE", "POSE", "FORMATION"]

# Palette Amorino (source : identité de marque officielle, terracotta/crème/brun)
C_TERRACOTTA = "#C97F3B"
C_TERRACOTTA_DARK = "#A8632A"
C_CREME = "#FDE8DC"
C_CREME_CLAIR = "#FFF9F4"
C_BRUN = "#6F2100"
C_BRUN_CLAIR = "#8C3A12"
C_ROSE = "#E8A0AE"
C_OR = "#D4A937"
C_ROUGE = "#C0524A"
C_VIOLET = "#8B6F9E"
C_VERT = "#6E9B5E"

COULEURS_TYPE = {
    "TRAVAIL": C_TERRACOTTA,
    "REPOS": "#C9B8A8",
    "MALADIE": C_ROUGE,
    "POSE": C_OR,
    "FORMATION": C_VIOLET,
}
JOURS_FR = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]

NB_SLOTS = 29  # créneaux de 30 min de 11:00 à 01:00 le lendemain
DEFAULT_RESPONSABLE_PASSWORD = "amorino2026"


# ============================== LOGO (SVG maison, style rose Amorino) ==============================
# Pas de logo officiel embarqué (droits d'usage) : ce SVG évoque la rose de gelato
# signature d'Amorino avec la palette de marque. Remplace-le par le vrai fichier
# logo Amorino (PNG/SVG) si tu en as un, en éditant LOGO_SVG ci-dessous ou en
# passant à st.image("logo_amorino.png") dans afficher_entete().
LOGO_SVG = f"""
<svg width="46" height="46" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
  <circle cx="50" cy="50" r="48" fill="{C_CREME}" stroke="{C_TERRACOTTA}" stroke-width="2"/>
  <g transform="translate(50,50)">
    <path d="M0,-30 C12,-22 12,-8 0,0 C-12,-8 -12,-22 0,-30 Z" fill="{C_TERRACOTTA}"/>
    <path d="M0,-30 C12,-22 12,-8 0,0 C-12,-8 -12,-22 0,-30 Z" fill="{C_TERRACOTTA_DARK}" transform="rotate(72)"/>
    <path d="M0,-30 C12,-22 12,-8 0,0 C-12,-8 -12,-22 0,-30 Z" fill="{C_TERRACOTTA}" transform="rotate(144)"/>
    <path d="M0,-30 C12,-22 12,-8 0,0 C-12,-8 -12,-22 0,-30 Z" fill="{C_TERRACOTTA_DARK}" transform="rotate(216)"/>
    <path d="M0,-30 C12,-22 12,-8 0,0 C-12,-8 -12,-22 0,-30 Z" fill="{C_TERRACOTTA}" transform="rotate(288)"/>
    <circle cx="0" cy="0" r="9" fill="{C_BRUN}"/>
  </g>
</svg>
"""


# ============================== STYLE GLOBAL ==============================

def injecter_css():
    st.markdown(f"""
    <style>
        .stApp {{
            background-color: {C_CREME_CLAIR};
        }}
        h1, h2, h3 {{
            color: {C_BRUN} !important;
            font-family: 'Georgia', 'Times New Roman', serif;
        }}
        [data-testid="stSidebar"] {{
            background-color: {C_CREME};
            border-right: 1px solid {C_TERRACOTTA}33;
        }}
        .stButton > button {{
            background-color: {C_TERRACOTTA};
            color: white;
            border: none;
            border-radius: 8px;
            font-weight: 600;
            transition: all 0.15s ease;
        }}
        .stButton > button:hover {{
            background-color: {C_TERRACOTTA_DARK};
            transform: translateY(-1px);
            box-shadow: 0 3px 8px {C_TERRACOTTA}55;
        }}
        .stDownloadButton > button {{
            background-color: {C_BRUN};
            color: white;
            border-radius: 8px;
        }}
        div[data-testid="stMetric"] {{
            background-color: white;
            border: 1px solid {C_TERRACOTTA}33;
            border-radius: 12px;
            padding: 12px 16px;
        }}
        .amorino-header {{
            display: flex;
            align-items: center;
            gap: 14px;
            padding: 10px 0 18px 0;
            border-bottom: 2px solid {C_TERRACOTTA}55;
            margin-bottom: 18px;
        }}
        .amorino-header h1 {{
            margin: 0;
            font-size: 26px;
        }}
        .amorino-header .sous-titre {{
            color: {C_BRUN_CLAIR};
            font-size: 13px;
            margin-top: 2px;
        }}
        .carte-profil {{
            background: white;
            border: 2px solid {C_TERRACOTTA}44;
            border-radius: 16px;
            padding: 28px 20px;
            text-align: center;
            transition: all 0.15s ease;
        }}
        .carte-profil:hover {{
            border-color: {C_TERRACOTTA};
            box-shadow: 0 6px 16px {C_TERRACOTTA}33;
        }}
        .badge-type {{
            display: inline-block;
            padding: 2px 10px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: 600;
            color: white;
        }}
        .effectif-cell {{
            display: inline-flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            width: 30px;
            height: 46px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 700;
            color: white;
            margin: 1px;
        }}
        .effectif-heure {{
            font-size: 9px;
            color: {C_BRUN_CLAIR};
            text-align: center;
            width: 30px;
            margin: 1px;
        }}
    </style>
    """, unsafe_allow_html=True)


def afficher_entete(sous_titre=""):
    st.markdown(f"""
    <div class="amorino-header">
        {LOGO_SVG}
        <div>
            <h1>Amorino Besançon</h1>
            <div class="sous-titre">{sous_titre}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


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
        conn.execute("""
            CREATE TABLE IF NOT EXISTS objectifs_effectif (
                slot_idx INTEGER PRIMARY KEY,
                effectif_cible INTEGER NOT NULL
            )
        """)
        noms_initiaux = ["ANAE", "CLOE", "LOUISE", "JEREMY", "LUCAS", "CELINE",
                          "LILOU", "ROMANE", "SUZE", "SINEM"]
        for nom in noms_initiaux:
            conn.execute("INSERT OR IGNORE INTO salaries (nom) VALUES (?)", (nom,))

        # Objectifs d'effectif par défaut : léger creux hors repas, pics déjeuner/soir.
        # A AJUSTER par le responsable dans l'onglet Réglages selon le vrai rush du point de vente.
        nb_deja = conn.execute("SELECT COUNT(*) FROM objectifs_effectif").fetchone()[0]
        if nb_deja == 0:
            for idx in range(NB_SLOTS):
                h = slot_index_to_time(idx)
                if (12 <= h.hour < 14) or (18 <= h.hour or h.hour < 1):
                    cible = 3
                else:
                    cible = 2
                conn.execute(
                    "INSERT INTO objectifs_effectif (slot_idx, effectif_cible) VALUES (?, ?)",
                    (idx, cible),
                )


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


def modifier_creneau_complet(creneau_id, salarie_id, jour, type_, heure_debut, heure_fin, commentaire=None):
    """Met à jour tous les champs, y compris salarié et jour (utile pour le drag & drop
    qui peut déplacer un créneau vers un autre salarié ou un autre horaire)."""
    with get_conn() as conn:
        maj = {
            "salarie_id": salarie_id,
            "jour": jour.isoformat() if isinstance(jour, date) else jour,
            "type": type_,
            "heure_debut": heure_debut.strftime("%H:%M") if isinstance(heure_debut, dtime) else heure_debut,
            "heure_fin": heure_fin.strftime("%H:%M") if isinstance(heure_fin, dtime) else heure_fin,
        }
        if commentaire is not None:
            maj["commentaire"] = commentaire
        set_clause = ", ".join(f"{k} = ?" for k in maj)
        conn.execute(f"UPDATE creneaux SET {set_clause} WHERE id = ?", (*maj.values(), creneau_id))


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


def creneaux_jour(jour: date):
    with get_conn() as conn:
        rows = conn.execute(
            """SELECT c.*, s.nom AS salarie_nom
               FROM creneaux c JOIN salaries s ON s.id = c.salarie_id
               WHERE c.jour = ?
               ORDER BY s.nom, c.heure_debut""",
            (jour.isoformat(),),
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


def get_objectifs():
    with get_conn() as conn:
        rows = conn.execute("SELECT * FROM objectifs_effectif ORDER BY slot_idx").fetchall()
        return {r["slot_idx"]: r["effectif_cible"] for r in rows}


def set_objectif(slot_idx, valeur):
    with get_conn() as conn:
        conn.execute(
            """INSERT INTO objectifs_effectif (slot_idx, effectif_cible) VALUES (?, ?)
               ON CONFLICT(slot_idx) DO UPDATE SET effectif_cible = excluded.effectif_cible""",
            (slot_idx, valeur),
        )


# ============================== CALCULS METIER ==============================

def slot_index_to_time(idx):
    total_minutes = 11 * 60 + idx * 30
    h = (total_minutes // 60) % 24
    m = total_minutes % 60
    return dtime(h, m)


def minutes_depuis_11h(h, m):
    raw = h * 60 + m
    if raw < 11 * 60:
        raw += 24 * 60
    return raw - 11 * 60


def duree_heures(heure_debut, heure_fin):
    if not heure_debut or not heure_fin:
        return 0.0
    fmt = "%H:%M"
    t1 = datetime.strptime(heure_debut, fmt)
    t2 = datetime.strptime(heure_fin, fmt)
    delta = (t2 - t1).total_seconds() / 3600
    if delta < 0:
        delta += 24
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


def alertes_repos(df_creneaux):
    if df_creneaux.empty:
        return []
    alertes = []
    for nom, grp in df_creneaux.groupby("salarie_nom"):
        if not (grp["type"] == "REPOS").any():
            alertes.append(f"{nom} n'a aucun jour de repos enregistré cette semaine.")
    return alertes


def effectif_par_slot(df_jour):
    """Retourne une liste de 29 entiers : nombre de salariés en TRAVAIL sur chaque créneau."""
    compte = [0] * NB_SLOTS
    if df_jour.empty:
        return compte
    travail = df_jour[df_jour["type"] == "TRAVAIL"]
    for _, r in travail.iterrows():
        if not r["heure_debut"] or not r["heure_fin"]:
            continue
        hd, md = map(int, r["heure_debut"].split(":"))
        hf, mf = map(int, r["heure_fin"].split(":"))
        deb = minutes_depuis_11h(hd, md)
        fin = minutes_depuis_11h(hf, mf)
        if fin <= deb:
            fin += 24 * 60
        for idx in range(NB_SLOTS):
            slot_min = idx * 30
            if deb <= slot_min < fin:
                compte[idx] += 1
    return compte


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
        st.markdown(f"### Semaine du {lundi.strftime('%d/%m/%Y')} au {dimanche.strftime('%d/%m/%Y')}")
    return lundi


def afficher_jauge_effectifs(jour: date):
    """Bande horizontale colorée : effectif réel vs cible, par créneau de 30 min."""
    df_jour = creneaux_jour(jour)
    effectifs = effectif_par_slot(df_jour)
    objectifs = get_objectifs()

    st.markdown("**Effectifs par créneau (repérer les trous sur les heures de rush)**")
    html = '<div style="display:flex;flex-wrap:wrap;align-items:flex-end;">'
    for idx in range(NB_SLOTS):
        heure = slot_index_to_time(idx)
        reel = effectifs[idx]
        cible = objectifs.get(idx, 2)
        if cible == 0:
            couleur = "#DDD3C8"
        elif reel >= cible:
            couleur = C_VERT
        elif reel >= max(cible - 1, 1):
            couleur = C_OR
        else:
            couleur = C_ROUGE
        label_heure = heure.strftime("%Hh%M") if heure.minute else heure.strftime("%Hh")
        html += (
            f'<div style="text-align:center;">'
            f'<div class="effectif-cell" style="background:{couleur};" title="{label_heure} : {reel}/{cible}">{reel}</div>'
            f'<div class="effectif-heure">{label_heure if idx % 2 == 0 else ""}</div>'
            f'</div>'
        )
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)

    trous = [
        slot_index_to_time(i).strftime("%H:%M")
        for i in range(NB_SLOTS)
        if objectifs.get(i, 2) > 0 and effectifs[i] < objectifs.get(i, 2)
    ]
    if trous:
        st.warning(f"⚠️ Sous-effectif à : {', '.join(trous)}")
    else:
        st.success("✅ Effectifs conformes aux objectifs sur toute la journée.")


def construire_evenements_calendrier(df_jour):
    events = []
    for _, r in df_jour.iterrows():
        if r["type"] == "TRAVAIL" and r["heure_debut"] and r["heure_fin"]:
            start = f"{r['jour']}T{r['heure_debut']}:00"
            hf, mf = map(int, r["heure_fin"].split(":"))
            hd, md = map(int, r["heure_debut"].split(":"))
            jour_fin = r["jour"]
            # Comparaison sur l'heure réelle (pas l'offset décalé) : si l'heure de fin
            # "sur l'horloge" est antérieure ou égale à l'heure de début, le créneau
            # traverse minuit (ex: 18:00 -> 01:00) et se termine donc le lendemain.
            if (hf * 60 + mf) <= (hd * 60 + md):
                jour_fin = (date.fromisoformat(r["jour"]) + timedelta(days=1)).isoformat()
            end = f"{jour_fin}T{r['heure_fin']}:00"
            titre = f"{r['heure_debut']}-{r['heure_fin']}"
        else:
            start = f"{r['jour']}T00:00:00"
            end = f"{r['jour']}T23:59:00"
            titre = r["type"]
        events.append({
            "id": str(r["id"]),
            "resourceId": str(r["salarie_id"]),
            "title": titre,
            "start": start,
            "end": end,
            "backgroundColor": COULEURS_TYPE.get(r["type"], "#999999"),
            "borderColor": COULEURS_TYPE.get(r["type"], "#999999"),
        })
    return events


# ============================== PAGES ==============================

def page_connexion():
    injecter_css()
    afficher_entete("Gestion du planning d'équipe")
    st.write("")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"""<div class="carte-profil"><div style="font-size:38px;">🧑‍🍳</div>
        <h3>Je suis salarié</h3><p style="color:{C_BRUN_CLAIR};font-size:13px;">
        Voir mon planning de la semaine</p></div>""", unsafe_allow_html=True)
        if st.button("Continuer en tant que salarié", key="btn_role_salarie", use_container_width=True):
            st.session_state["ecran"] = "choix_salarie"
            st.rerun()
    with c2:
        st.markdown(f"""<div class="carte-profil"><div style="font-size:38px;">📋</div>
        <h3>Je suis responsable</h3><p style="color:{C_BRUN_CLAIR};font-size:13px;">
        Gérer le planning de l'équipe</p></div>""", unsafe_allow_html=True)
        if st.button("Continuer en tant que responsable", key="btn_role_resp", use_container_width=True):
            st.session_state["ecran"] = "mdp_responsable"
            st.rerun()

    st.write("")
    if st.session_state.get("ecran") == "choix_salarie":
        st.markdown("---")
        st.subheader("Quel est ton prénom ?")
        df_sal = liste_salaries()
        if df_sal.empty:
            st.warning("Aucun salarié enregistré pour le moment.")
        else:
            cols = st.columns(4)
            for i, nom in enumerate(df_sal["nom"].tolist()):
                with cols[i % 4]:
                    if st.button(nom, key=f"tuile_{nom}", use_container_width=True):
                        st.session_state["role"] = "salarie"
                        st.session_state["salarie_id"] = int(df_sal[df_sal["nom"] == nom]["id"].iloc[0])
                        st.session_state["salarie_nom"] = nom
                        st.rerun()

    if st.session_state.get("ecran") == "mdp_responsable":
        st.markdown("---")
        st.subheader("Connexion responsable")
        with st.form("form_mdp"):
            mdp = st.text_input("Mot de passe", type="password")
            ok = st.form_submit_button("Connexion")
            if ok:
                vrai_mdp = st.secrets.get("RESPONSABLE_PASSWORD", DEFAULT_RESPONSABLE_PASSWORD)
                if mdp == vrai_mdp:
                    st.session_state["role"] = "responsable"
                    st.rerun()
                else:
                    st.error("Mot de passe incorrect.")


def page_salarie():
    injecter_css()
    afficher_entete(f"Planning de {st.session_state['salarie_nom']}")
    if st.sidebar.button("← Changer de profil"):
        for k in ["role", "salarie_id", "salarie_nom", "ecran"]:
            st.session_state.pop(k, None)
        st.rerun()

    lundi = selecteur_semaine("salarie")
    df = creneaux_salarie(st.session_state["salarie_id"], lundi)

    if df.empty:
        st.info("Aucun créneau enregistré pour cette semaine.")
        return

    resume = total_heures_par_salarie(df)
    if not resume.empty:
        st.metric("Heures travaillées cette semaine", f"{resume['heures_travaillees'].iloc[0]:.1f} h")

    dimanche = lundi + timedelta(days=6)
    jours = [lundi + timedelta(days=i) for i in range(7)]
    for i, jour in enumerate(jours):
        jour_creneaux = df[df["jour"] == jour.isoformat()]
        est_aujourdhui = jour == date.today()
        if jour_creneaux.empty:
            continue
        entete = f"**{'🔵 ' if est_aujourdhui else ''}{JOURS_FR[i]} {jour.strftime('%d/%m')}**"
        st.markdown(entete)
        for _, row in jour_creneaux.iterrows():
            couleur = COULEURS_TYPE.get(row["type"], "#999999")
            horaire = (f"{row['heure_debut']} → {row['heure_fin']}"
                       if row["type"] == "TRAVAIL" and row["heure_debut"] else row["type"])
            st.markdown(
                f'<div style="background:{couleur}22;border-left:4px solid {couleur};'
                f'padding:8px 14px;border-radius:6px;margin-bottom:8px;">{horaire}</div>',
                unsafe_allow_html=True,
            )


def page_responsable():
    injecter_css()
    afficher_entete("Espace responsable")
    if st.sidebar.button("← Se déconnecter"):
        st.session_state.pop("role", None)
        st.rerun()

    onglet = st.sidebar.radio(
        "Navigation",
        ["Vue journée", "Vue semaine", "Ajouter un créneau", "Salariés", "Réglages effectifs", "Export"],
    )

    if onglet == "Vue journée":
        page_vue_journee()
    elif onglet == "Vue semaine":
        page_vue_semaine()
    elif onglet == "Ajouter un créneau":
        page_ajouter()
    elif onglet == "Salariés":
        page_salaries_admin()
    elif onglet == "Réglages effectifs":
        page_reglages_effectifs()
    elif onglet == "Export":
        page_export()


def page_vue_journee():
    st.subheader("Vue journée — glisse-dépose pour réorganiser")
    if "vj_jour" not in st.session_state:
        st.session_state["vj_jour"] = date.today()

    c1, c2, c3 = st.columns([1, 2, 1])
    with c1:
        if st.button("◀ Jour préc."):
            st.session_state["vj_jour"] -= timedelta(days=1)
    with c3:
        if st.button("Jour suiv. ▶"):
            st.session_state["vj_jour"] += timedelta(days=1)
    with c2:
        st.session_state["vj_jour"] = st.date_input(
            "Jour affiché", value=st.session_state["vj_jour"], label_visibility="collapsed"
        )

    jour = st.session_state["vj_jour"]
    st.markdown(f"#### {JOURS_FR[jour.weekday()]} {jour.strftime('%d/%m/%Y')}")

    afficher_jauge_effectifs(jour)
    st.markdown("---")

    df_sal = liste_salaries()
    if df_sal.empty:
        st.warning("Ajoute d'abord des salariés dans l'onglet Salariés.")
        return

    resources = [{"id": str(r["id"]), "title": r["nom"]} for _, r in df_sal.iterrows()]
    df_jour = creneaux_jour(jour)
    events = construire_evenements_calendrier(df_jour)

    calendar_options = {
        "initialView": "resourceTimelineDay",
        "initialDate": jour.isoformat(),
        "headerToolbar": False,
        "slotMinTime": "11:00:00",
        "slotMaxTime": "25:00:00",
        "slotDuration": "00:30:00",
        "slotLabelFormat": {"hour": "2-digit", "minute": "2-digit", "hour12": False},
        "resourceAreaHeaderContent": "Salarié",
        "resourceAreaWidth": "150px",
        "resources": resources,
        "editable": True,
        "eventStartEditable": True,
        "eventDurationEditable": True,
        "eventResourceEditable": True,
        "selectable": True,
        "selectMirror": True,
        "nowIndicator": True,
        "height": 480,
        "locale": "fr",
    }
    custom_css = f"""
        .fc-timeline-event {{ border-radius: 6px; font-size: 12px; padding: 2px 4px; }}
        .fc-datagrid-cell-cushion {{ font-weight: 600; color: {C_BRUN}; }}
        .fc-timeline-slot-frame {{ font-size: 11px; }}
        .fc-scrollgrid {{ border-radius: 8px; overflow: hidden; }}
    """

    cal = calendar(
        events=events,
        options=calendar_options,
        custom_css=custom_css,
        callbacks=["eventChange", "eventClick", "select"],
        key=f"cal_{jour.isoformat()}",
    )

    if cal:
        callback = cal.get("callback")

        if callback == "eventChange":
            ev = cal["eventChange"]["event"]
            creneau_id = int(ev["id"])
            resource_id = ev.get("resourceId")
            if not resource_id and ev.get("resourceIds"):
                resource_id = ev["resourceIds"][0]
            start_dt = datetime.fromisoformat(ev["start"].replace("Z", ""))
            end_dt = datetime.fromisoformat(ev["end"].replace("Z", "")) if ev.get("end") else start_dt
            modifier_creneau_complet(
                creneau_id,
                int(resource_id),
                start_dt.date(),
                "TRAVAIL",
                start_dt.time(),
                end_dt.time(),
            )
            st.toast("Créneau déplacé.")
            st.rerun()

        elif callback == "eventClick":
            creneau_id = int(cal["eventClick"]["event"]["id"])
            st.session_state["creneau_selectionne"] = creneau_id

        elif callback == "select":
            sel = cal["select"]
            resource = sel.get("resource", {})
            st.session_state["preremplissage_creneau"] = {
                "salarie_id": resource.get("id"),
                "jour": jour,
                "debut": sel.get("start"),
                "fin": sel.get("end"),
            }
            st.info("Plage sélectionnée. Complète les détails ci-dessous puis valide.")

    st.caption(
        "Glisse un créneau pour changer son horaire, dépose-le sur une autre ligne pour le "
        "réattribuer à un autre salarié. Clique sur un créneau pour le modifier ou le supprimer. "
        "Clique-glisse sur une zone vide pour créer un nouveau créneau."
    )

    if st.session_state.get("creneau_selectionne"):
        st.markdown("---")
        afficher_edition_creneau(st.session_state["creneau_selectionne"])

    if st.session_state.get("preremplissage_creneau"):
        st.markdown("---")
        afficher_creation_rapide(st.session_state["preremplissage_creneau"], df_sal)


def afficher_edition_creneau(creneau_id):
    with get_conn() as conn:
        row = conn.execute(
            """SELECT c.*, s.nom AS salarie_nom FROM creneaux c
               JOIN salaries s ON s.id = c.salarie_id WHERE c.id = ?""",
            (creneau_id,),
        ).fetchone()
    if not row:
        st.session_state.pop("creneau_selectionne", None)
        return

    st.subheader(f"Modifier le créneau — {row['salarie_nom']} le {row['jour']}")
    with st.form("form_edit_journee"):
        type_ = st.selectbox("Type", TYPES_ABSENCE, index=TYPES_ABSENCE.index(row["type"]))
        c1, c2 = st.columns(2)
        hd_defaut = datetime.strptime(row["heure_debut"], "%H:%M").time() if row["heure_debut"] else dtime(11, 0)
        hf_defaut = datetime.strptime(row["heure_fin"], "%H:%M").time() if row["heure_fin"] else dtime(18, 0)
        with c1:
            heure_debut = st.time_input("Heure de début", value=hd_defaut) if type_ == "TRAVAIL" else None
        with c2:
            heure_fin = st.time_input("Heure de fin", value=hf_defaut) if type_ == "TRAVAIL" else None
        commentaire = st.text_input("Commentaire", value=row["commentaire"] or "")
        c3, c4, c5 = st.columns(3)
        with c3:
            maj = st.form_submit_button("Enregistrer")
        with c4:
            suppr = st.form_submit_button("🗑️ Supprimer")
        with c5:
            annuler = st.form_submit_button("Annuler")
        if maj:
            modifier_creneau_complet(
                creneau_id, row["salarie_id"], date.fromisoformat(row["jour"]),
                type_, heure_debut, heure_fin, commentaire,
            )
            st.session_state.pop("creneau_selectionne", None)
            st.rerun()
        if suppr:
            supprimer_creneau(creneau_id)
            st.session_state.pop("creneau_selectionne", None)
            st.rerun()
        if annuler:
            st.session_state.pop("creneau_selectionne", None)
            st.rerun()


def afficher_creation_rapide(prefill, df_sal):
    st.subheader("Nouveau créneau")
    noms = df_sal["nom"].tolist()
    ids = df_sal["id"].tolist()
    salarie_defaut = 0
    if prefill.get("salarie_id"):
        try:
            salarie_defaut = ids.index(int(prefill["salarie_id"]))
        except (ValueError, TypeError):
            salarie_defaut = 0

    try:
        hd_defaut = datetime.fromisoformat(prefill["debut"].replace("Z", "")).time()
        hf_defaut = datetime.fromisoformat(prefill["fin"].replace("Z", "")).time()
    except Exception:
        hd_defaut, hf_defaut = dtime(11, 0), dtime(18, 0)

    with st.form("form_creation_rapide"):
        nom = st.selectbox("Salarié", noms, index=salarie_defaut)
        type_ = st.selectbox("Type", TYPES_ABSENCE)
        c1, c2 = st.columns(2)
        with c1:
            heure_debut = st.time_input("Début", value=hd_defaut) if type_ == "TRAVAIL" else None
        with c2:
            heure_fin = st.time_input("Fin", value=hf_defaut) if type_ == "TRAVAIL" else None
        c3, c4 = st.columns(2)
        with c3:
            valider = st.form_submit_button("Créer le créneau")
        with c4:
            annuler = st.form_submit_button("Annuler")
        if valider:
            salarie_id = int(df_sal[df_sal["nom"] == nom]["id"].iloc[0])
            ajouter_creneau(salarie_id, prefill["jour"], type_, heure_debut, heure_fin)
            st.session_state.pop("preremplissage_creneau", None)
            st.rerun()
        if annuler:
            st.session_state.pop("preremplissage_creneau", None)
            st.rerun()


def page_vue_semaine():
    st.subheader("Vue semaine")
    lundi = selecteur_semaine("resp")
    df = creneaux_semaine(lundi)
    for a in alertes_repos(df):
        st.warning(f"⚠️ {a}")
    if df.empty:
        st.info("Aucun créneau cette semaine.")
        return
    jours = [lundi + timedelta(days=i) for i in range(7)]
    for i, jour in enumerate(jours):
        jour_creneaux = df[df["jour"] == jour.isoformat()]
        if jour_creneaux.empty:
            continue
        st.markdown(f"**{JOURS_FR[i]} {jour.strftime('%d/%m')}**")
        cols = st.columns(min(len(jour_creneaux), 6))
        for idx, (_, row) in enumerate(jour_creneaux.iterrows()):
            col = cols[idx % len(cols)]
            couleur = COULEURS_TYPE.get(row["type"], "#999999")
            horaire = (f"{row['heure_debut']}-{row['heure_fin']}"
                       if row["type"] == "TRAVAIL" and row["heure_debut"] else row["type"])
            with col:
                st.markdown(
                    f'<div style="background:{couleur}22;border-left:4px solid {couleur};'
                    f'padding:6px 10px;border-radius:4px;margin-bottom:6px;">'
                    f'<b>{row["salarie_nom"]}</b><br><span style="font-size:0.85em">{horaire}</span></div>',
                    unsafe_allow_html=True,
                )
    st.markdown("---")
    st.subheader("Total heures par salarié")
    st.dataframe(total_heures_par_salarie(df), use_container_width=True, hide_index=True)


def page_ajouter():
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


def page_salaries_admin():
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


def page_reglages_effectifs():
    st.subheader("Objectifs d'effectif par créneau")
    st.caption(
        "Définis combien de salariés tu veux au minimum sur chaque tranche horaire. "
        "La vue journée signalera en rouge/orange les créneaux en dessous de l'objectif."
    )
    objectifs = get_objectifs()
    with st.form("form_objectifs"):
        cols = st.columns(6)
        nouvelles_valeurs = {}
        for idx in range(NB_SLOTS):
            heure = slot_index_to_time(idx)
            with cols[idx % 6]:
                nouvelles_valeurs[idx] = st.number_input(
                    heure.strftime("%H:%M"), min_value=0, max_value=20,
                    value=objectifs.get(idx, 2), step=1, key=f"obj_{idx}",
                )
        if st.form_submit_button("Enregistrer les objectifs"):
            for idx, val in nouvelles_valeurs.items():
                set_objectif(idx, val)
            st.success("Objectifs mis à jour.")
            st.rerun()


def page_export():
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
