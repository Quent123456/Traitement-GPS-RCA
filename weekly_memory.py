"""
==============================================================================
MODULE DE MÉMOIRE SÉLECTIVE & RAPPORTS HEBDOMADAIRES — RUGBY CLUB AMIÉNOIS (RCA)
==============================================================================
Module 100% indépendant gérant :
1. L'archivage manuel et ciblé de séances validées par semaine ISO.
2. Le stockage persistant des séances dans data/weekly_memory/.
3. La fusion et l'agrégation statistique multi-séances par joueur :
   - Somme pour le volume (Distance totale, Haute intensité, VHSR/Sprint, Accel/Decel, Temps)
   - Maximum pour les pics (Top speed km/h)
   - Moyenne / ratio pour les intensités relatives (m/min, % marche)
   - Comptage des séances participées
4. Le comparatif de charge journalière (Mardi vs Mercredi vs Vendredi vs Dimanche).
5. L'export PDF multi-pages officiel du Rapport Hebdomadaire pour les coachs.
==============================================================================
"""

from __future__ import annotations

import datetime
import io
import json
import os
import re
from typing import Any, Dict, List, Optional, Tuple, Union

from fpdf import FPDF
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

import photo_manager

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
WEEKLY_MEMORY_DIR = os.path.join(PROJECT_ROOT, "data", "weekly_memory")
os.makedirs(WEEKLY_MEMORY_DIR, exist_ok=True)

JOURS_SEANCE_STANDARDS = [
    "Mardi (Entraînement)",
    "Mercredi (Entraînement)",
    "Vendredi (Entraînement)",
    "Dimanche (Match)",
    "Lundi (Récup / Séance)",
    "Jeudi (Entraînement)",
    "Samedi (Match / Mise en place)",
    "Autre séance",
]


def get_current_iso_week() -> Tuple[int, int]:
    """Retourne l'année et le numéro de semaine ISO actuels (ex: (2026, 43))."""
    today = datetime.date.today()
    iso = today.isocalendar()
    return int(iso[0]), int(iso[1])


def get_week_file_path(year: int, week_num: int) -> str:
    """Retourne le chemin absolu du fichier JSON pour une semaine donnée."""
    os.makedirs(WEEKLY_MEMORY_DIR, exist_ok=True)
    return os.path.join(WEEKLY_MEMORY_DIR, f"week_{year}_W{week_num:02d}.json")


def list_saved_weeks() -> List[Dict[str, Any]]:
    """
    Scanne data/weekly_memory/ et retourne la liste triée des semaines archivées.
    Chaque élément : { 'key': '2026-W43', 'year': 2026, 'week_num': 43, 'label': 'Semaine 43 - 2026', 'nb_sessions': 3 }
    """
    os.makedirs(WEEKLY_MEMORY_DIR, exist_ok=True)
    res = []
    pattern = re.compile(r"^week_(\d{4})_W(\d{2})\.json$")
    for f in os.listdir(WEEKLY_MEMORY_DIR):
        m = pattern.match(f)
        if m:
            y = int(m.group(1))
            w = int(m.group(2))
            fp = os.path.join(WEEKLY_MEMORY_DIR, f)
            try:
                with open(fp, "r", encoding="utf-8") as jf:
                    data = json.load(jf)
                    sessions = data.get("sessions", {})
                    nb_s = len(sessions)
            except Exception:
                nb_s = 0
            res.append({
                "key": f"{y}-W{w:02d}",
                "year": y,
                "week_num": w,
                "label": f"Semaine {w:02d} — {y} ({nb_s} séance{'s' if nb_s > 1 else ''})",
                "nb_sessions": nb_s,
                "file_path": fp,
            })
    res.sort(key=lambda x: (x["year"], x["week_num"]), reverse=True)
    return res


def load_week_data(year: int, week_num: int) -> Dict[str, Any]:
    """Charge les données d'une semaine enregistrée."""
    fp = get_week_file_path(year, week_num)
    if not os.path.exists(fp):
        return {
            "key": f"{year}-W{week_num:02d}",
            "year": year,
            "week_num": week_num,
            "sessions": {},
        }
    try:
        with open(fp, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {
            "key": f"{year}-W{week_num:02d}",
            "year": year,
            "week_num": week_num,
            "sessions": {},
        }


def save_session_to_memory(
    df_indiv: pd.DataFrame,
    year: int,
    week_num: int,
    day_type: str,
    date_str: Optional[str] = None,
    notes: str = "",
    overwrite: bool = True,
) -> Tuple[bool, str]:
    """
    Sauvegarde manuellement une séance validée dans la mémoire hebdomadaire.
    Gère la persistance sous format JSON structuré avec typage nettoyé.
    """
    if df_indiv is None or df_indiv.empty:
        return False, "Le tableau des données joueurs est vide."

    clean_day = day_type.strip()
    session_id = clean_day.lower().replace(" ", "_").replace("(", "").replace(")", "").replace("'", "")
    clean_date = date_str.strip() if date_str else datetime.date.today().strftime("%d/%m/%Y")

    week_data = load_week_data(year, week_num)
    existing_sessions = week_data.setdefault("sessions", {})

    if session_id in existing_sessions and not overwrite:
        return False, f"Une séance pour « {clean_day} » existe déjà pour la semaine {week_num}."

    # Sérialisation propre des enregistrements joueurs
    serialized_players: List[Dict[str, Any]] = []
    col_player = "Nom ; Prénom" if "Nom ; Prénom" in df_indiv.columns else df_indiv.columns[0]

    for _, row in df_indiv.iterrows():
        p_dict: Dict[str, Any] = {}
        for col in df_indiv.columns:
            val = row[col]
            if pd.isna(val) or val is None:
                p_dict[col] = None
            elif isinstance(val, (np.floating, float)):
                p_dict[col] = round(float(val), 2)
            elif isinstance(val, (np.integer, int)):
                p_dict[col] = int(val)
            else:
                p_dict[col] = str(val).strip()
        serialized_players.append(p_dict)

    # Métadonnées calculées de la séance
    total_dist_km = 0.0
    if "Distance (km)" in df_indiv.columns:
        total_dist_km = float(pd.to_numeric(df_indiv["Distance (km)"], errors="coerce").fillna(0).sum())

    total_hi_m = 0.0
    hi_col = "Haute intensité Podium (m)" if "Haute intensité Podium (m)" in df_indiv.columns else " Distance haute intensité (m)"
    if hi_col in df_indiv.columns:
        total_hi_m = float(pd.to_numeric(df_indiv[hi_col], errors="coerce").fillna(0).sum())

    avg_dpm = 0.0
    if "Distance Per Min (m/min)" in df_indiv.columns:
        avg_dpm = float(pd.to_numeric(df_indiv["Distance Per Min (m/min)"], errors="coerce").fillna(0).mean())

    existing_sessions[session_id] = {
        "session_id": session_id,
        "day_type": clean_day,
        "date": clean_date,
        "recorded_at": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "notes": notes,
        "num_players": len(df_indiv),
        "total_distance_km": round(total_dist_km, 2),
        "total_hi_m": round(total_hi_m, 1),
        "avg_dpm": round(avg_dpm, 1),
        "players_data": serialized_players,
    }

    week_data["last_updated"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    fp = get_week_file_path(year, week_num)
    try:
        with open(fp, "w", encoding="utf-8") as f:
            json.dump(week_data, f, ensure_ascii=False, indent=2)
        return True, f"Séance « {clean_day} » archivée avec succès en Semaine {week_num} ({len(df_indiv)} joueurs)."
    except Exception as e:
        return False, f"Erreur lors de la sauvegarde : {str(e)}"


def delete_session_from_memory(year: int, week_num: int, session_id: str) -> Tuple[bool, str]:
    """Supprime une séance spécifique d'une semaine archivée."""
    week_data = load_week_data(year, week_num)
    sessions = week_data.get("sessions", {})
    if session_id in sessions:
        day_label = sessions[session_id].get("day_type", session_id)
        del sessions[session_id]
        fp = get_week_file_path(year, week_num)
        try:
            if not sessions:
                # Si plus aucune séance dans la semaine, supprimer le fichier
                if os.path.exists(fp):
                    os.remove(fp)
                return True, f"Séance « {day_label} » supprimée. Semaine {week_num} vide nettoyée."
            else:
                week_data["last_updated"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                with open(fp, "w", encoding="utf-8") as f:
                    json.dump(week_data, f, ensure_ascii=False, indent=2)
                return True, f"Séance « {day_label} » supprimée de la Semaine {week_num}."
        except Exception as e:
            return False, f"Erreur lors de la suppression : {str(e)}"
    return False, "Séance introuvable dans cette semaine."


def delete_player_records(player_name: str) -> int:
    """
    Supprime toutes les traces d'un joueur personnalisé dans la mémoire hebdomadaire
    lorsque son profil est supprimé de l'application (respect de la règle sans orphelin).
    """
    p_clean = player_name.strip().lower()
    total_cleaned = 0
    weeks = list_saved_weeks()
    for w_info in weeks:
        fp = w_info["file_path"]
        try:
            with open(fp, "r", encoding="utf-8") as f:
                data = json.load(f)
            changed = False
            for s_id, s_data in data.get("sessions", {}).items():
                p_list = s_data.get("players_data", [])
                new_list = [p for p in p_list if str(p.get("Nom ; Prénom", "")).strip().lower() != p_clean]
                if len(new_list) != len(p_list):
                    s_data["players_data"] = new_list
                    s_data["num_players"] = len(new_list)
                    changed = True
                    total_cleaned += 1
            if changed:
                with open(fp, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    return total_cleaned


# ==============================================================================
# MOTEUR D'AGRÉGATION HEBDOMADAIRE MULTI-SÉANCES
# ==============================================================================

def aggregate_weekly_sessions(
    week_data: Dict[str, Any],
    selected_session_ids: Optional[List[str]] = None,
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Agrège les séances sélectionnées sur la semaine selon les règles athlétiques du rugby :
    - Somme (Cumul) : Distance totale, Haute intensité, VHSR/Sprint, Accélérations/Décélérations, Temps de jeu.
    - Maximum (Pic) : Top Speed (km/h).
    - Moyenne pondérée / ratio : Distance/min (m/min), % de marche.
    - Métrique de présence : Nombre de séances participées.
    Retourne (df_hebdo, synthese_jours).
    """
    sessions = week_data.get("sessions", {})
    if not sessions:
        return pd.DataFrame(), {}

    if selected_session_ids is None:
        target_sessions = list(sessions.values())
    else:
        target_sessions = [s for s_id, s in sessions.items() if s_id in selected_session_ids]

    if not target_sessions:
        return pd.DataFrame(), {}

    # 1. Comparatif journalier de la charge
    synthese_jours: Dict[str, Dict[str, Any]] = {}
    for s in target_sessions:
        s_id = s.get("session_id", "")
        synthese_jours[s_id] = {
            "day_type": s.get("day_type", ""),
            "date": s.get("date", ""),
            "num_players": s.get("num_players", 0),
            "total_distance_km": s.get("total_distance_km", 0.0),
            "total_hi_m": s.get("total_hi_m", 0.0),
            "avg_dpm": s.get("avg_dpm", 0.0),
        }

    # 2. Consolidation individuelle par joueur
    all_player_rows: List[Dict[str, Any]] = []
    for s in target_sessions:
        day_name = s.get("day_type", "")
        p_list = s.get("players_data", [])
        for p in p_list:
            row_copy = dict(p)
            row_copy["_session_day"] = day_name
            all_player_rows.append(row_copy)

    if not all_player_rows:
        return pd.DataFrame(), synthese_jours

    df_all = pd.DataFrame(all_player_rows)
    p_col = "Nom ; Prénom"

    # Conversion numérique robuste
    numeric_cols = [
        "Temps de jeu ", "Distance (km)", "Distance_m",
        "Haute intensité Podium (m)", " Distance haute intensité (m)",
        "HSR Z3 (m)", "HSR_m", "VHSR (m)", "Top Speed (km/h)",
        "Distance Per Min (m/min)", "%Marche ", "%course", "%HSR ", "%VHSR",
        "Accel/Dec", "Accelerations Zone Count: 3 - 4 m/s/s", "Accelerations Zone Count: > 4 m/s/s",
        "Deceleration Zone Count: 3 - 4 m/s/s", "Deceleration Zone Count: > 4 m/s/s",
    ]
    for c in numeric_cols:
        if c in df_all.columns:
            df_all[c] = pd.to_numeric(df_all[c], errors="coerce").fillna(0.0)

    # Colonne standardisée pour la distance en mètres et haute intensité
    if "Distance_m" not in df_all.columns and "Distance (km)" in df_all.columns:
        df_all["Distance_m"] = df_all["Distance (km)"] * 1000.0
    if "Distance (km)" not in df_all.columns and "Distance_m" in df_all.columns:
        df_all["Distance (km)"] = df_all["Distance_m"] / 1000.0

    hi_col_target = "Haute intensité Podium (m)"
    if hi_col_target not in df_all.columns and " Distance haute intensité (m)" in df_all.columns:
        df_all[hi_col_target] = df_all[" Distance haute intensité (m)"]

    grouped = df_all.groupby(p_col)

    agg_records = []
    for name, group in grouped:
        nb_seances = len(group)
        jours_presence = ", ".join(group["_session_day"].unique().tolist())

        # Groupe et poste de référence (mode / dernier)
        poste_det = group["Poste_Detaille"].iloc[-1] if "Poste_Detaille" in group.columns else "Inconnu"
        groupe_rca = group["Groupe"].iloc[-1] if "Groupe" in group.columns else "Trois-quarts"

        # CUMULS (SOMMES HEBDOMADAIRES)
        tot_time = float(group["Temps de jeu "].sum()) if "Temps de jeu " in group.columns else 0.0
        tot_dist_km = float(group["Distance (km)"].sum()) if "Distance (km)" in group.columns else 0.0
        tot_dist_m = tot_dist_km * 1000.0

        tot_hi_m = float(group[hi_col_target].sum()) if hi_col_target in group.columns else 0.0
        tot_hsr_z3 = float(group["HSR Z3 (m)"].sum()) if "HSR Z3 (m)" in group.columns else (float(group["HSR_m"].sum()) if "HSR_m" in group.columns else 0.0)
        tot_vhsr_m = float(group["VHSR (m)"].sum()) if "VHSR (m)" in group.columns else 0.0
        tot_acc_dec = float(group["Accel/Dec"].sum()) if "Accel/Dec" in group.columns else 0.0

        tot_acc_34 = float(group["Accelerations Zone Count: 3 - 4 m/s/s"].sum()) if "Accelerations Zone Count: 3 - 4 m/s/s" in group.columns else 0.0
        tot_acc_gt4 = float(group["Accelerations Zone Count: > 4 m/s/s"].sum()) if "Accelerations Zone Count: > 4 m/s/s" in group.columns else 0.0
        tot_dec_34 = float(group["Deceleration Zone Count: 3 - 4 m/s/s"].sum()) if "Deceleration Zone Count: 3 - 4 m/s/s" in group.columns else 0.0
        tot_dec_gt4 = float(group["Deceleration Zone Count: > 4 m/s/s"].sum()) if "Deceleration Zone Count: > 4 m/s/s" in group.columns else 0.0

        # PICS (MAXIMUMS HEBDOMADAIRES)
        top_speed = float(group["Top Speed (km/h)"].max()) if "Top Speed (km/h)" in group.columns else 0.0

        # INTENSITÉS RELATIVES & MOYENNES
        dpm = (tot_dist_m / tot_time) if tot_time > 0 else 0.0
        hsr_pm = (tot_hi_m / tot_time) if tot_time > 0 else 0.0
        pct_marche = float(group["%Marche "].mean()) if "%Marche " in group.columns else 0.0

        agg_records.append({
            "Nom ; Prénom": name,
            "Poste_Detaille": poste_det,
            "Groupe": groupe_rca,
            "Nombre de séances": nb_seances,
            "Séances": jours_presence,
            "Temps total (min)": round(tot_time, 1),
            "Distance Totale (km)": round(tot_dist_km, 2),
            "Distance Totale (m)": round(tot_dist_m, 1),
            "Haute intensité (m)": round(tot_hi_m, 0),
            "Haute intensité Podium (m)": round(tot_hi_m, 0),
            " Distance haute intensité (m)": round(tot_hi_m, 0),
            "HSR Z3 (m)": round(tot_hsr_z3, 0),
            "Distance Sprint VHSR (m)": round(tot_vhsr_m, 0),
            "Top Speed (km/h)": round(top_speed, 1),
            "Distance/min (m/min)": round(dpm, 1),
            "HSR m/min": round(hsr_pm, 1),
            "Accel/Dec total": int(round(tot_acc_dec)),
            "Accélérations >3m/s²": int(round(tot_acc_34 + tot_acc_gt4)),
            "Décélérations >3m/s²": int(round(tot_dec_34 + tot_dec_gt4)),
            "% de marche moyen": round(pct_marche, 1),
            "%Marche ": round(pct_marche, 1),
        })

    df_hebdo = pd.DataFrame(agg_records)
    df_hebdo.sort_values(by="Distance Totale (km)", ascending=False, inplace=True)
    return df_hebdo, synthese_jours


# ==============================================================================
# GÉNÉRATION DES PODIUMS HEBDOMADAIRES (AVEC PHOTOS & HTML NETTOYÉ)
# ==============================================================================

def generer_podiums_hebdomadaires(df_hebdo: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
    """Génère les podiums de la semaine : Top Speed, % Marche, Haute Intensité, et Distance Totale."""
    def _make_player_entry(rank: int, name: str, stat_str: str, poste: str = "") -> Dict[str, Any]:
        p_name = name.strip() if name else ""
        return {
            "rank": rank,
            "name": p_name,
            "poste": poste if poste else "RCA",
            "stat": stat_str,
            "avatar_src": photo_manager.get_player_image_b64(p_name, size=80 if rank == 1 else 65),
        }

    def _placeholder(rank: int) -> Dict[str, Any]:
        return {
            "rank": rank,
            "name": "-",
            "poste": "",
            "stat": "-",
            "avatar_src": photo_manager.get_player_image_b64("", size=80 if rank == 1 else 65),
        }

    if df_hebdo.empty:
        return {
            "Distance Totale (km)": {"icon": "🏃", "players": [_placeholder(1), _placeholder(2), _placeholder(3)]},
            "Top Speed (km/h)": {"icon": "⚡", "players": [_placeholder(1), _placeholder(2), _placeholder(3)]},
            "Haute Intensité (m)": {"icon": "🔥", "players": [_placeholder(1), _placeholder(2), _placeholder(3)]},
        }

    podiums: Dict[str, Dict[str, Any]] = {}

    # 1. DISTANCE TOTALE
    df_s_dist = df_hebdo.sort_values(by="Distance Totale (km)", ascending=False)
    p_dist = []
    for r, (_, row) in enumerate(df_s_dist.head(3).iterrows(), start=1):
        p_dist.append(_make_player_entry(r, row["Nom ; Prénom"], f"{row['Distance Totale (km)']:.2f} km", row.get("Poste_Detaille", "")))
    while len(p_dist) < 3:
        p_dist.append(_placeholder(len(p_dist) + 1))
    podiums["Distance Totale (km)"] = {"icon": "🏃", "players": p_dist}

    # 2. TOP SPEED
    df_s_ts = df_hebdo.sort_values(by="Top Speed (km/h)", ascending=False)
    p_ts = []
    for r, (_, row) in enumerate(df_s_ts.head(3).iterrows(), start=1):
        p_ts.append(_make_player_entry(r, row["Nom ; Prénom"], f"{row['Top Speed (km/h)']:.1f} km/h", row.get("Poste_Detaille", "")))
    while len(p_ts) < 3:
        p_ts.append(_placeholder(len(p_ts) + 1))
    podiums["Top Speed (km/h)"] = {"icon": "⚡", "players": p_ts}

    # 3. HAUTE INTENSITÉ
    df_s_hi = df_hebdo.sort_values(by="Haute intensité (m)", ascending=False)
    p_hi = []
    for r, (_, row) in enumerate(df_s_hi.head(3).iterrows(), start=1):
        p_hi.append(_make_player_entry(r, row["Nom ; Prénom"], f"{int(round(row['Haute intensité (m)'])):,d} m".replace(",", " "), row.get("Poste_Detaille", "")))
    while len(p_hi) < 3:
        p_hi.append(_placeholder(len(p_hi) + 1))
    podiums["Haute Intensité (m)"] = {"icon": "🔥", "players": p_hi}

    return podiums


# ==============================================================================
# EXPORT PDF MULTI-PAGES RAPPORT HEBDOMADAIRE (FORMAT PAYSAGE COACHS)
# ==============================================================================

def safe_latin1(s: Any) -> str:
    """Assure qu'une chaîne est compatible avec l'encodage Latin-1 (Helvetica standard FPDF)."""
    if s is None:
        return ""
    txt = str(s)
    replacements = {
        "\u2019": "'",
        "\u2018": "'",
        "\u201c": '"',
        "\u201d": '"',
        "\u201e": '"',
        "\u2013": "-",
        "\u2014": "-",
        "\u2022": "-",
        "\u2026": "...",
        "\u0153": "oe",
        "\u0152": "OE",
        "\u20ac": "EUR",
        "\xa0": " ",
        "\u200b": "",
        "\ufeff": "",
    }
    for old, new in replacements.items():
        txt = txt.replace(old, new)
    try:
        return txt.encode("latin-1", errors="replace").decode("latin-1")
    except Exception:
        return "".join([c if ord(c) < 256 else "?" for c in txt])


class RCAWeeklyPDFReport(FPDF):
    """Classe FPDF Landscape dédiée au rapport hebdomadaire du RCA."""
    def __init__(self, week_label: str, dates_label: str):
        super().__init__(orientation="L", unit="mm", format="A4")
        self.week_label = week_label
        self.dates_label = dates_label
        self.set_auto_page_break(auto=True, margin=10)

    def header(self):
        # Bandeau officiel RCA
        self.set_fill_color(11, 25, 44)  # #0B192C
        self.rect(0, 0, 297, 18, "F")
        self.set_fill_color(212, 175, 55)  # Or #D4AF37
        self.rect(0, 18, 297, 1.2, "F")

        # Titre
        self.set_font("Helvetica", "B", 13)
        self.set_text_color(255, 255, 255)
        self.set_xy(12, 4)
        self.cell(160, 6, safe_latin1("RUGBY CLUB AMIÉNOIS  |  RAPPORT HEBDOMADAIRE DE CHARGE GPS"), align="L")

        self.set_font("Helvetica", "B", 10)
        self.set_text_color(212, 175, 55)
        self.set_xy(180, 4)
        dates_disp = f" - {self.dates_label}" if self.dates_label else ""
        self.cell(105, 6, safe_latin1(f"{self.week_label.upper()}{dates_disp}"), align="R")

    def footer(self):
        self.set_y(-9)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(100, 116, 139)
        self.cell(140, 6, safe_latin1("Cellule Performance & Préparation Physique RCA Amiens - Confidentiel Staff"), align="L")
        self.cell(147, 6, safe_latin1(f"Page {self.page_no()}"), align="R")


def generate_weekly_pdf_report(
    df_hebdo: Optional[pd.DataFrame] = None,
    year: int = 2026,
    week_num: int = 1,
    synthese_jours: Optional[Union[Dict[str, Any], List[Dict[str, Any]]]] = None,
    notes_coach: str = "",
    **kwargs: Any,
) -> bytes:
    """Génère un rapport PDF multi-pages consolidé de la semaine pour le staff."""
    if df_hebdo is None:
        df_hebdo = kwargs.get("df_agg", pd.DataFrame())

    if synthese_jours is None:
        synthese_jours = kwargs.get("selected_sessions_meta", {})

    if isinstance(synthese_jours, list):
        synthese_dict = {f"s_{i}": s for i, s in enumerate(synthese_jours) if isinstance(s, dict)}
    elif isinstance(synthese_jours, dict):
        synthese_dict = synthese_jours
    else:
        synthese_dict = {}

    notes_final = notes_coach or kwargs.get("notes", "")

    week_label = f"Semaine {week_num:02d} - {year}"
    dates = [s.get("date", "") for s in synthese_dict.values() if isinstance(s, dict) and s.get("date")]
    dates_str = f"Du {dates[0]} au {dates[-1]}" if len(dates) >= 2 else (dates[0] if dates else "")

    pdf = RCAWeeklyPDFReport(week_label=week_label, dates_label=dates_str)

    # --------------------------------------------------------------------------
    # PAGE 1 : SYNTHÈSE GLOBALE DE LA SEMAINE & CHARGE JOURNALIÈRE
    # --------------------------------------------------------------------------
    pdf.add_page()
    pdf.set_y(24)

    # KPIs Globaux
    tot_km = df_hebdo["Distance Totale (km)"].sum() if not df_hebdo.empty else 0.0
    tot_hi = df_hebdo["Haute intensité (m)"].sum() if not df_hebdo.empty else 0.0
    nb_joueurs_tot = len(df_hebdo)
    max_speed = df_hebdo["Top Speed (km/h)"].max() if not df_hebdo.empty else 0.0
    nb_seances = len(synthese_dict)

    kpis = [
        ("Volume Cumulé Équipe", f"{tot_km:.1f} km"),
        ("Haute Intensité Cumulée", f"{int(round(tot_hi)):,d} m".replace(",", " ")),
        ("Joueurs Actifs Semaine", f"{nb_joueurs_tot} joueurs"),
        ("Vitesse Max Semaine", f"{max_speed:.1f} km/h"),
        ("Séances Fusionnées", f"{nb_seances} séances"),
    ]

    col_w = 54.0
    for idx, (label, val) in enumerate(kpis):
        x = 12 + idx * (col_w + 2.5)
        pdf.set_fill_color(248, 250, 252)
        pdf.set_draw_color(226, 232, 240)
        pdf.rect(x, 24, col_w, 20, "FD")
        pdf.set_xy(x, 26)
        pdf.set_font("Helvetica", "B", 8)
        pdf.set_text_color(100, 116, 139)
        pdf.cell(col_w, 4, label.upper(), align="C")
        pdf.set_xy(x, 31)
        pdf.set_font("Helvetica", "B", 13)
        pdf.set_text_color(31, 78, 121)
        pdf.cell(col_w, 8, val, align="C")

    # Tableau comparatif journalier
    pdf.set_xy(12, 48)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(11, 25, 44)
    pdf.cell(100, 6, "1. RÉPARTITION DE LA CHARGE PAR SÉANCE DANS LA SEMAINE :", align="L")

    y_t = 56
    pdf.set_xy(12, y_t)
    pdf.set_fill_color(31, 78, 121)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(65, 7, "Séance / Type", 1, 0, "L", fill=True)
    pdf.cell(35, 7, "Date", 1, 0, "C", fill=True)
    pdf.cell(40, 7, "Effectif Présent", 1, 0, "C", fill=True)
    pdf.cell(45, 7, "Distance Totale (km)", 1, 0, "C", fill=True)
    pdf.cell(45, 7, "Haute Intensité (m)", 1, 0, "C", fill=True)
    pdf.cell(43, 7, "Moyenne m/min", 1, 1, "C", fill=True)

    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(15, 23, 42)
    for s_id, s in synthese_dict.items():
        pdf.set_x(12)
        pdf.cell(65, 7, safe_latin1(s.get("day_type", "")), 1, 0, "L")
        pdf.cell(35, 7, safe_latin1(s.get("date", "")), 1, 0, "C")
        pdf.cell(40, 7, safe_latin1(f"{s.get('num_players', 0)} joueurs"), 1, 0, "C")
        pdf.cell(45, 7, safe_latin1(f"{s.get('total_distance_km', 0.0):.2f} km"), 1, 0, "C")
        pdf.cell(45, 7, safe_latin1(f"{int(round(s.get('total_hi_m', 0.0))):,d} m".replace(",", " ")), 1, 0, "C")
        pdf.cell(43, 7, safe_latin1(f"{s.get('avg_dpm', 0.0):.1f} m/min"), 1, 1, "C")

    # Podiums Hebdomadaires
    pdf.set_xy(12, 94)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(11, 25, 44)
    pdf.cell(100, 6, "2. PODIUMS DE LA SEMAINE (TOUS POSTES CONFONDUS) :", align="L")

    podiums = generer_podiums_hebdomadaires(df_hebdo)
    cat_keys = ["Distance Totale (km)", "Top Speed (km/h)", "Haute Intensité (m)"]
    p_col_w = 89.0

    for idx, c_key in enumerate(cat_keys):
        c_data = podiums.get(c_key, {})
        players = c_data.get("players", [])
        px = 12 + idx * (p_col_w + 3.0)
        py = 102

        # Card header
        pdf.set_xy(px, py)
        pdf.set_fill_color(11, 25, 44)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(p_col_w, 7, safe_latin1(f"{c_data.get('icon', '')} {c_key.upper()}"), 1, 1, "C", fill=True)

        for p in players:
            rank = p.get("rank", 1)
            p_name = p.get("name", "-")
            p_stat = p.get("stat", "-")
            p_pos = p.get("poste", "")

            # Médaille
            bg_color = (255, 243, 199) if rank == 1 else ((241, 245, 249) if rank == 2 else (255, 237, 213))
            pdf.set_fill_color(*bg_color)
            pdf.set_draw_color(226, 232, 240)
            cur_y = pdf.get_y()
            pdf.set_x(px)
            pdf.rect(px, cur_y, p_col_w, 16, "FD")

            # Photo du joueur
            photo_path = photo_manager.get_player_photo_path(p_name)
            has_photo = bool(photo_path and os.path.isfile(photo_path) and os.path.basename(photo_path) != "avatar_par_defaut.png")
            if has_photo:
                try:
                    pdf.image(photo_path, x=px + 2, y=cur_y + 1.5, w=13, h=13)
                except Exception:
                    has_photo = False

            text_x = px + (17 if has_photo else 4)
            pdf.set_xy(text_x, cur_y + 2)
            pdf.set_font("Helvetica", "B", 9)
            pdf.set_text_color(15, 23, 42)
            pdf.cell(p_col_w - 20, 4, safe_latin1(f"#{rank} {p_name}"), align="L")

            pdf.set_xy(text_x, cur_y + 6.5)
            pdf.set_font("Helvetica", "I", 7.5)
            pdf.set_text_color(100, 116, 139)
            pdf.cell(p_col_w - 20, 3.5, safe_latin1(f"Poste : {p_pos}"), align="L")

            pdf.set_xy(text_x, cur_y + 10.5)
            pdf.set_font("Helvetica", "B", 9.5)
            pdf.set_text_color(163, 0, 0 if rank == 1 else 31)
            pdf.cell(p_col_w - 20, 4, safe_latin1(p_stat), align="L")

            pdf.set_y(cur_y + 16)

    # --------------------------------------------------------------------------
    # PAGES SUIVANTES : TABLEAUX PAR CATÉGORIE DE POSTE (5 de devant, 3e Ligne, 3/4)
    # --------------------------------------------------------------------------
    groupes_rca = ["5 de devant", "3ème Ligne", "Trois-quarts"]

    for grp in groupes_rca:
        df_sub = df_hebdo[df_hebdo["Groupe"] == grp]
        if df_sub.empty:
            continue

        pdf.add_page()
        pdf.set_y(23)
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(11, 25, 44)
        pdf.cell(200, 7, safe_latin1(f"BILAN HEBDOMADAIRE - {grp.upper()} ({len(df_sub)} joueurs)"), align="L")

        # Table header
        pdf.set_xy(12, 32)
        pdf.set_fill_color(31, 78, 121)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 8)

        cols_def = [
            ("Joueur", 46, "L"),
            ("Poste", 26, "C"),
            ("Présences", 32, "C"),
            ("Temps (min)", 22, "C"),
            ("Dist. (km)", 24, "C"),
            ("H.I. (m)", 26, "C"),
            ("Sprint (m)", 22, "C"),
            ("Vmax (km/h)", 24, "C"),
            ("m/min", 22, "C"),
            ("Acc/Dec", 25, "C"),
        ]

        for title, w, align in cols_def:
            pdf.cell(w, 7, safe_latin1(title), 1, 0, align, fill=True)
        pdf.ln(7)

        pdf.set_font("Helvetica", "", 8)
        pdf.set_text_color(15, 23, 42)
        fill_row = False

        for _, row in df_sub.iterrows():
            pdf.set_x(12)
            bg = (248, 250, 252) if fill_row else (255, 255, 255)
            pdf.set_fill_color(*bg)

            p_nom = str(row["Nom ; Prénom"])[:22]
            p_pos = str(row["Poste_Detaille"])[:13]
            p_pres = f"{row['Nombre de séances']} séance{'s' if row['Nombre de séances'] > 1 else ''}"
            p_tps = f"{row['Temps total (min)']:.0f}"
            p_dist = f"{row['Distance Totale (km)']:.2f}"
            p_hi = f"{int(round(row['Haute intensité (m)'])):,d}".replace(",", " ")
            p_vhsr = f"{int(round(row['Distance Sprint VHSR (m)'])):,d}".replace(",", " ")
            p_vmax = f"{row['Top Speed (km/h)']:.1f}"
            p_dpm = f"{row['Distance/min (m/min)']:.1f}"
            p_acc = f"{row['Accel/Dec total']}"

            pdf.cell(46, 6.5, safe_latin1(p_nom), 1, 0, "L", fill=True)
            pdf.cell(26, 6.5, safe_latin1(p_pos), 1, 0, "C", fill=True)
            pdf.cell(32, 6.5, safe_latin1(p_pres), 1, 0, "C", fill=True)
            pdf.cell(22, 6.5, safe_latin1(p_tps), 1, 0, "C", fill=True)
            pdf.cell(24, 6.5, safe_latin1(p_dist), 1, 0, "C", fill=True)
            pdf.cell(26, 6.5, safe_latin1(p_hi), 1, 0, "C", fill=True)
            pdf.cell(22, 6.5, safe_latin1(p_vhsr), 1, 0, "C", fill=True)
            pdf.cell(24, 6.5, safe_latin1(p_vmax), 1, 0, "C", fill=True)
            pdf.cell(22, 6.5, safe_latin1(p_dpm), 1, 0, "C", fill=True)
            pdf.cell(25, 6.5, safe_latin1(p_acc), 1, 1, "C", fill=True)
            fill_row = not fill_row

    return bytes(pdf.output())
