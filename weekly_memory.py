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
import pathlib
import re
from typing import Any, Dict, List, Optional, Tuple, Union
import zipfile

from fpdf import FPDF
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

import photo_manager

PROJECT_ROOT = pathlib.Path(__file__).parent.resolve()
DATA_DIR = PROJECT_ROOT / "data"
WEEKLY_MEMORY_DIR = str(DATA_DIR / "weekly_memory")
WEEKLY_PDFS_DIR = str(DATA_DIR / "weekly_pdfs")
TARGETS_CONFIG_FILE = str(DATA_DIR / "targets_config.json")
os.makedirs(WEEKLY_MEMORY_DIR, exist_ok=True)
os.makedirs(WEEKLY_PDFS_DIR, exist_ok=True)


def format_readable_date(val: Any) -> str:
    """
    Convertit n'importe quelle valeur de date (numéro de série Excel ex: 46315, 46316, 46318,
    timestamp, chaîne ISO, ou datetime) en date lisible au format français JJ/MM/AAAA.
    """
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)):
        return val.strftime("%d/%m/%Y")
    
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ["nan", "none", "nat"]:
        return ""

    # Test si c'est un numéro de série Excel (ex: 46315 ou "46315.0")
    try:
        val_clean_num = val_str.replace(",", ".")
        val_float = float(val_clean_num)
        if 30000 <= val_float <= 80000:
            dt = pd.to_datetime(val_float, unit="D", origin="1899-12-30")
            return dt.strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        pass

    # Si c'est déjà JJ/MM/AAAA
    if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", val_str):
        parts = val_str.split("/")
        return f"{int(parts[0]):02d}/{int(parts[1]):02d}/{parts[2]}"

    # Si format ISO YYYY-MM-DD
    if re.match(r"^\d{4}-\d{2}-\d{2}", val_str):
        try:
            dt = pd.to_datetime(val_str)
            return dt.strftime("%d/%m/%Y")
        except Exception:
            pass

    try:
        dt = pd.to_datetime(val_str, errors="coerce", dayfirst=True)
        if pd.notna(dt):
            return dt.strftime("%d/%m/%Y")
    except Exception:
        pass

    return val_str


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


def get_session_pdf_path(year: int, week_num: int, session_id: str) -> str:
    """Retourne le chemin absolu du fichier PDF archivé pour une séance spécifique."""
    os.makedirs(WEEKLY_PDFS_DIR, exist_ok=True)
    return os.path.join(WEEKLY_PDFS_DIR, f"week_{year}_W{week_num:02d}_{session_id}.pdf")


def get_session_pdf_bytes(
    year: int,
    week_num: int,
    session_id: str,
    generate_if_missing: bool = True,
) -> Optional[bytes]:
    """
    Récupère les octets (bytes) du PDF archivé de la séance.
    Si le fichier physique n'est pas encore présent sur disque mais que la séance existe en mémoire,
    il régénère le PDF à la volée à partir des données de joueurs sauvegardées.
    """
    pdf_path = get_session_pdf_path(year, week_num, session_id)
    if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000:
        try:
            with open(pdf_path, "rb") as f:
                return f.read()
        except Exception:
            pass

    if not generate_if_missing:
        return None

    # Régénération dynamique à la volée
    week_data = load_week_data(year, week_num)
    sessions = week_data.get("sessions", {})
    if session_id not in sessions:
        return None

    s_info = sessions[session_id]
    p_data = s_info.get("players_data", [])
    if not p_data:
        return None

    df_indiv = pd.DataFrame(p_data)
    date_val = format_readable_date(s_info.get("date", ""))
    day_val = s_info.get("day_type", "")

    try:
        import app
        pdf_bytes = app.generer_rapport_pdf_coachs(
            df_indiv=df_indiv,
            date_seance=date_val or day_val,
            synthese_note=s_info.get("notes", ""),
            moyennes_groupes=s_info.get("group_stats"),
        )
        if pdf_bytes and len(pdf_bytes) > 1000:
            try:
                os.makedirs(WEEKLY_PDFS_DIR, exist_ok=True)
                with open(pdf_path, "wb") as f_out:
                    f_out.write(pdf_bytes)
                s_info["has_pdf"] = True
                s_info["pdf_filename"] = os.path.basename(pdf_path)
                fp = get_week_file_path(year, week_num)
                with open(fp, "w", encoding="utf-8") as f_w:
                    json.dump(week_data, f_w, ensure_ascii=False, indent=2)
            except Exception:
                pass
            return pdf_bytes
    except Exception as e:
        print(f"Erreur lors de la régénération du PDF journalier pour {session_id} : {e}")

    return None


def save_session_to_memory(
    df_indiv: pd.DataFrame,
    year: int,
    week_num: int,
    day_type: str,
    date_str: Optional[str] = None,
    notes: str = "",
    overwrite: bool = True,
    pdf_bytes: Optional[bytes] = None,
    group_stats: Optional[Dict[str, Any]] = None,
) -> Tuple[bool, str]:
    """
    Sauvegarde manuellement une séance validée dans la mémoire hebdomadaire.
    Archivage simultané des données structurées et du fichier PDF journalier dans data/weekly_pdfs/.
    """
    if df_indiv is None or df_indiv.empty:
        return False, "Le tableau des données joueurs est vide."

    clean_day = day_type.strip()
    session_id = clean_day.lower().replace(" ", "_").replace("(", "").replace(")", "").replace("'", "")
    clean_date = format_readable_date(date_str) if date_str else datetime.date.today().strftime("%d/%m/%Y")
    if not clean_date:
        clean_date = datetime.date.today().strftime("%d/%m/%Y")

    week_data = load_week_data(year, week_num)
    existing_sessions = week_data.setdefault("sessions", {})

    if session_id in existing_sessions and not overwrite:
        return False, f"Une séance pour « {clean_day} » existe déjà pour la semaine {week_num}."

    # Sérialisation propre des enregistrements joueurs
    serialized_players: List[Dict[str, Any]] = []
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

    # Archivage automatique du fichier PDF si fourni
    pdf_filename = f"week_{year}_W{week_num:02d}_{session_id}.pdf"
    pdf_path = get_session_pdf_path(year, week_num, session_id)
    has_pdf = False
    if pdf_bytes and len(pdf_bytes) > 1000:
        try:
            os.makedirs(WEEKLY_PDFS_DIR, exist_ok=True)
            with open(pdf_path, "wb") as f_pdf:
                f_pdf.write(pdf_bytes)
            has_pdf = True
        except Exception as e:
            print(f"Avertissement : Impossible d'écrire le PDF journalier sur disque : {e}")
    elif os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000:
        has_pdf = True

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
        "has_pdf": has_pdf,
        "pdf_filename": pdf_filename,
        "group_stats": group_stats or {},
        "players_data": serialized_players,
    }

    week_data["last_updated"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    fp = get_week_file_path(year, week_num)
    try:
        with open(fp, "w", encoding="utf-8") as f:
            json.dump(week_data, f, ensure_ascii=False, indent=2)

        try:
            import github_sync
            github_sync.push_file_to_github(
                fp,
                commit_message=f"data: enregistrement séance {clean_day} (Semaine {week_num:02d} — {year})"
            )
        except Exception:
            pass

        return True, f"Séance « {clean_day} » archivée avec succès en Semaine {week_num} ({len(df_indiv)} joueurs)."
    except Exception as e:
        return False, f"Erreur lors de la sauvegarde : {str(e)}"


def delete_session_from_memory(year: int, week_num: int, session_id: str) -> Tuple[bool, str]:
    """Supprime une séance spécifique d'une semaine archivée ainsi que son PDF associé."""
    week_data = load_week_data(year, week_num)
    sessions = week_data.get("sessions", {})
    if session_id in sessions:
        day_label = sessions[session_id].get("day_type", session_id)
        del sessions[session_id]

        # Suppression du fichier PDF archivé s'il existe
        pdf_path = get_session_pdf_path(year, week_num, session_id)
        if os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except Exception:
                pass

        fp = get_week_file_path(year, week_num)
        try:
            if not sessions:
                if os.path.exists(fp):
                    os.remove(fp)
                try:
                    import github_sync
                    github_sync.delete_file_from_github(
                        fp,
                        commit_message=f"data: suppression de la semaine {week_num:02d} — {year} (vide)"
                    )
                except Exception:
                    pass
                return True, f"Séance « {day_label} » supprimée. Semaine {week_num} vide nettoyée."
            else:
                week_data["last_updated"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                with open(fp, "w", encoding="utf-8") as f:
                    json.dump(week_data, f, ensure_ascii=False, indent=2)
                try:
                    import github_sync
                    github_sync.push_file_to_github(
                        fp,
                        commit_message=f"data: suppression séance {day_label} (Semaine {week_num:02d} — {year})"
                    )
                except Exception:
                    pass
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
                try:
                    import github_sync
                    github_sync.push_file_to_github(
                        fp,
                        commit_message=f"data: nettoyage données joueur {player_name} (Semaine {w_info.get('week_num', '')})"
                    )
                except Exception:
                    pass
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
# GESTION DES OBJECTIFS DE CHARGE ATHLÉTIQUE (PAR GROUPE DE POSTE)
# ==============================================================================

PREREGLAGES_OBJECTIFS: Dict[str, Dict[str, Dict[str, float]]] = {
    "Standard / Référence RCA": {
        "5 de devant": {"Objectif Distance (m)": 4500.0, "Obj. HSR": 250.0, "Obj. VHSR": 50.0, "Obj. ACC/DEC": 35.0},
        "3ème Ligne": {"Objectif Distance (m)": 5500.0, "Obj. HSR": 450.0, "Obj. VHSR": 100.0, "Obj. ACC/DEC": 45.0},
        "Trois-quarts": {"Objectif Distance (m)": 6200.0, "Obj. HSR": 700.0, "Obj. VHSR": 250.0, "Obj. ACC/DEC": 50.0},
    },
    "Mardi (Aérobie & Volume)": {
        "5 de devant": {"Objectif Distance (m)": 4800.0, "Obj. HSR": 220.0, "Obj. VHSR": 40.0, "Obj. ACC/DEC": 32.0},
        "3ème Ligne": {"Objectif Distance (m)": 5600.0, "Obj. HSR": 380.0, "Obj. VHSR": 80.0, "Obj. ACC/DEC": 40.0},
        "Trois-quarts": {"Objectif Distance (m)": 6200.0, "Obj. HSR": 550.0, "Obj. VHSR": 180.0, "Obj. ACC/DEC": 45.0},
    },
    "Mercredi (Intensité & Vitesse)": {
        "5 de devant": {"Objectif Distance (m)": 4200.0, "Obj. HSR": 280.0, "Obj. VHSR": 60.0, "Obj. ACC/DEC": 38.0},
        "3ème Ligne": {"Objectif Distance (m)": 5200.0, "Obj. HSR": 500.0, "Obj. VHSR": 120.0, "Obj. ACC/DEC": 48.0},
        "Trois-quarts": {"Objectif Distance (m)": 5800.0, "Obj. HSR": 750.0, "Obj. VHSR": 300.0, "Obj. ACC/DEC": 55.0},
    },
    "Vendredi (Mise en place / Affûtage)": {
        "5 de devant": {"Objectif Distance (m)": 3200.0, "Obj. HSR": 140.0, "Obj. VHSR": 30.0, "Obj. ACC/DEC": 22.0},
        "3ème Ligne": {"Objectif Distance (m)": 3800.0, "Obj. HSR": 250.0, "Obj. VHSR": 60.0, "Obj. ACC/DEC": 28.0},
        "Trois-quarts": {"Objectif Distance (m)": 4200.0, "Obj. HSR": 350.0, "Obj. VHSR": 120.0, "Obj. ACC/DEC": 32.0},
    },
    "Dimanche (Match)": {
        "5 de devant": {"Objectif Distance (m)": 5500.0, "Obj. HSR": 350.0, "Obj. VHSR": 80.0, "Obj. ACC/DEC": 45.0},
        "3ème Ligne": {"Objectif Distance (m)": 6800.0, "Obj. HSR": 650.0, "Obj. VHSR": 180.0, "Obj. ACC/DEC": 55.0},
        "Trois-quarts": {"Objectif Distance (m)": 7500.0, "Obj. HSR": 900.0, "Obj. VHSR": 400.0, "Obj. ACC/DEC": 60.0},
    },
}

DEFAULT_OBJECTIFS_CONFIG: Dict[str, Dict[str, float]] = {
    "5 de devant": {"Objectif Distance (m)": 4500.0, "Obj. HSR": 250.0, "Obj. VHSR": 50.0, "Obj. ACC/DEC": 35.0},
    "3ème Ligne": {"Objectif Distance (m)": 5500.0, "Obj. HSR": 450.0, "Obj. VHSR": 100.0, "Obj. ACC/DEC": 45.0},
    "Trois-quarts": {"Objectif Distance (m)": 6200.0, "Obj. HSR": 700.0, "Obj. VHSR": 250.0, "Obj. ACC/DEC": 50.0},
    "Tous les groupes": {"Objectif Distance (m)": 5400.0, "Obj. HSR": 460.0, "Obj. VHSR": 130.0, "Obj. ACC/DEC": 43.0},
}


def _calculer_moyenne_tous_groupes(cfg: Dict[str, Dict[str, float]]) -> Dict[str, float]:
    """Calcule la moyenne pondérée / équiprobable de référence pour 'Tous les groupes'."""
    keys = ["Objectif Distance (m)", "Obj. HSR", "Obj. VHSR", "Obj. ACC/DEC"]
    grps = ["5 de devant", "3ème Ligne", "Trois-quarts"]
    res = {}
    for k in keys:
        vals = [float(cfg.get(g, {}).get(k, 0.0)) for g in grps if g in cfg]
        res[k] = round(sum(vals) / max(len(vals), 1), 1)
    return res


def charger_objectifs_config() -> Dict[str, Dict[str, float]]:
    """
    Charge les objectifs configurés depuis data/targets_config.json.
    Si le fichier n'existe pas encore, il est automatiquement initialisé avec les standards RCA.
    """
    if os.path.exists(TARGETS_CONFIG_FILE):
        try:
            with open(TARGETS_CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict) and "5 de devant" in data:
                    if "Tous les groupes" not in data:
                        data["Tous les groupes"] = _calculer_moyenne_tous_groupes(data)
                    return data
        except Exception as e:
            print(f"Avertissement targets_config.json : {e}")

    cfg = dict(DEFAULT_OBJECTIFS_CONFIG)
    sauvegarder_objectifs_config(cfg)
    return cfg


def sauvegarder_objectifs_config(cfg: Dict[str, Dict[str, float]]) -> bool:
    """Enregistre les objectifs configurés dans data/targets_config.json."""
    try:
        os.makedirs(os.path.dirname(TARGETS_CONFIG_FILE), exist_ok=True)
        cfg_to_save = dict(cfg)
        cfg_to_save["Tous les groupes"] = _calculer_moyenne_tous_groupes(cfg_to_save)
        with open(TARGETS_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg_to_save, f, ensure_ascii=False, indent=2)

        try:
            import github_sync
            github_sync.push_file_to_github(
                TARGETS_CONFIG_FILE,
                commit_message="config: mise à jour des objectifs de performance par poste"
            )
        except Exception:
            pass

        return True
    except Exception as e:
        print(f"Erreur enregistrement targets_config.json : {e}")
        return False


# ==============================================================================
# SÉCURITÉ CLOUD : SAUVEGARDE & RESTAURATION (ARCHIVE ZIP & JSON)
# ==============================================================================

def creer_archive_backup_zip() -> bytes:
    """
    Génère une archive ZIP complète de toute la mémoire hebdomadaire :
    - Tous les JSON de data/weekly_memory/
    - Tous les PDFs de data/weekly_pdfs/
    - La configuration des objectifs data/targets_config.json
    - Les mappings persistants config_rca_gps.json & config_rca_joueurs_custom.json
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        if os.path.exists(WEEKLY_MEMORY_DIR):
            for fname in os.listdir(WEEKLY_MEMORY_DIR):
                if fname.endswith(".json"):
                    fpath = os.path.join(WEEKLY_MEMORY_DIR, fname)
                    zf.write(fpath, arcname=f"weekly_memory/{fname}")

        if os.path.exists(WEEKLY_PDFS_DIR):
            for fname in os.listdir(WEEKLY_PDFS_DIR):
                if fname.endswith(".pdf"):
                    fpath = os.path.join(WEEKLY_PDFS_DIR, fname)
                    zf.write(fpath, arcname=f"weekly_pdfs/{fname}")

        if os.path.exists(TARGETS_CONFIG_FILE):
            zf.write(TARGETS_CONFIG_FILE, arcname="targets_config.json")

        for cfg_name in ["config_rca_gps.json", "config_rca_joueurs_custom.json"]:
            cp = os.path.join(str(PROJECT_ROOT), cfg_name)
            if os.path.exists(cp):
                zf.write(cp, arcname=cfg_name)

    buf.seek(0)
    return buf.getvalue()


def restaurer_archive_backup(uploaded_file: Any) -> Tuple[bool, str]:
    """
    Restaure une sauvegarde à partir d'un fichier ZIP ou JSON uploadé par l'utilisateur.
    Garantit une persistance absolue en cas de réinitialisation d'un conteneur Streamlit Cloud.
    """
    if uploaded_file is None:
        return False, "Aucun fichier fourni."

    filename = getattr(uploaded_file, "name", "").lower()
    file_bytes = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()

    if not file_bytes:
        return False, "Le fichier de sauvegarde est vide."

    os.makedirs(WEEKLY_MEMORY_DIR, exist_ok=True)
    os.makedirs(WEEKLY_PDFS_DIR, exist_ok=True)

    # 1. Format ZIP
    if filename.endswith(".zip") or zipfile.is_zipfile(io.BytesIO(file_bytes)):
        try:
            nb_json = 0
            nb_pdf = 0
            restored_paths: List[str] = []
            with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as zf:
                for member in zf.infolist():
                    m_name = member.filename.replace("\\", "/")
                    if m_name.endswith("/") or ".." in m_name:
                        continue
                    base = os.path.basename(m_name)

                    if m_name.startswith("weekly_memory/") or (base.startswith("week_") and base.endswith(".json")):
                        target_path = os.path.join(WEEKLY_MEMORY_DIR, base)
                        with open(target_path, "wb") as f_out:
                            f_out.write(zf.read(member))
                        restored_paths.append(target_path)
                        nb_json += 1
                    elif m_name.startswith("weekly_pdfs/") or base.endswith(".pdf"):
                        target_path = os.path.join(WEEKLY_PDFS_DIR, base)
                        with open(target_path, "wb") as f_out:
                            f_out.write(zf.read(member))
                        nb_pdf += 1
                    elif base == "targets_config.json":
                        with open(TARGETS_CONFIG_FILE, "wb") as f_out:
                            f_out.write(zf.read(member))
                        restored_paths.append(TARGETS_CONFIG_FILE)
                    elif base in ["config_rca_gps.json", "config_rca_joueurs_custom.json"]:
                        target_path = os.path.join(str(PROJECT_ROOT), base)
                        with open(target_path, "wb") as f_out:
                            f_out.write(zf.read(member))
                        restored_paths.append(target_path)

            try:
                import github_sync
                for rp in restored_paths:
                    github_sync.push_file_to_github(
                        rp,
                        commit_message=f"backup: restauration {os.path.basename(rp)}"
                    )
            except Exception:
                pass

            return True, f"Restauration réussie ! {nb_json} semaine(s) et {nb_pdf} PDF(s) restaurés."
        except Exception as e:
            return False, f"Erreur lors de la décompression : {str(e)}"

    # 2. Format JSON
    elif filename.endswith(".json"):
        try:
            raw_text = file_bytes.decode("utf-8")
            data = json.loads(raw_text)
            restored_path = None
            if isinstance(data, dict):
                if "sessions" in data:
                    y = data.get("year", 2026)
                    w = data.get("week_num", 1)
                    target_path = get_week_file_path(y, w)
                    with open(target_path, "w", encoding="utf-8") as f_out:
                        json.dump(data, f_out, ensure_ascii=False, indent=2)
                    restored_path = target_path
                    res_msg = f"Semaine {w} ({y}) restaurée avec succès ({len(data.get('sessions', {}))} séances)."
                elif "5 de devant" in data and "Trois-quarts" in data:
                    with open(TARGETS_CONFIG_FILE, "w", encoding="utf-8") as f_out:
                        json.dump(data, f_out, ensure_ascii=False, indent=2)
                    restored_path = TARGETS_CONFIG_FILE
                    res_msg = "Configuration des objectifs restaurée avec succès."
                else:
                    target_path = os.path.join(WEEKLY_MEMORY_DIR, os.path.basename(filename))
                    with open(target_path, "w", encoding="utf-8") as f_out:
                        json.dump(data, f_out, ensure_ascii=False, indent=2)
                    restored_path = target_path
                    res_msg = f"Fichier {filename} importé dans la mémoire hebdomadaire."

                if restored_path:
                    try:
                        import github_sync
                        github_sync.push_file_to_github(
                            restored_path,
                            commit_message=f"backup: restauration {os.path.basename(restored_path)}"
                        )
                    except Exception:
                        pass

                return True, res_msg
        except Exception as e:
            return False, f"Erreur lecture JSON : {str(e)}"

    return False, "Format non reconnu (.zip ou .json attendu)."


# ==============================================================================
# GRAPHIQUES MATPLOTLIB POUR LE RAPPORT HEBDOMADAIRE (PDF & STREAMLIT)
# ==============================================================================


def generer_graphique_hebdo_groupes_matplotlib(df_agg: pd.DataFrame) -> io.BytesIO:
    """
    Génère le diagramme combiné à barres comparatif de la semaine pour la Page 1 du PDF et le Dashboard :
    - Distance Moyenne (m) par Groupe (barres bleu marine #1F4E79)
    - HSR Z3 Moyen (m) par Groupe (barres rouge #A30000)
    sur les 3 groupes officiels du RCA : 5 de devant, 3ème Ligne, Trois-quarts.
    """
    fig, ax1 = plt.subplots(figsize=(6.2, 3.6), dpi=180)
    fig.patch.set_facecolor("#FFFFFF")
    ax1.set_facecolor("#FFFFFF")

    groupes = ["5 de devant", "3ème Ligne", "Trois-quarts"]
    dist_vals: List[float] = []
    hsr_vals: List[float] = []

    for g in groupes:
        sub = df_agg[df_agg["Groupe"] == g] if not df_agg.empty and "Groupe" in df_agg.columns else pd.DataFrame()
        if not sub.empty:
            d_val = float(sub["Distance Totale (m)"].mean()) if "Distance Totale (m)" in sub.columns else (float(sub["Distance Totale (km)"].mean()) * 1000.0 if "Distance Totale (km)" in sub.columns else 0.0)
            h_val = float(sub["HSR Z3 (m)"].mean()) if "HSR Z3 (m)" in sub.columns else (float(sub["Haute intensité (m)"].mean()) if "Haute intensité (m)" in sub.columns else 0.0)
        else:
            d_val = 0.0
            h_val = 0.0
        dist_vals.append(d_val)
        hsr_vals.append(h_val)

    x = np.arange(len(groupes))
    width = 0.35

    rects1 = ax1.bar(x - width/2, dist_vals, width, label="Distance Moyenne (m)", color="#1F4E79", edgecolor="none", zorder=3)
    ax2 = ax1.twinx()
    rects2 = ax2.bar(x + width/2, hsr_vals, width, label="HSR Z3 Moyen (m)", color="#A30000", edgecolor="none", zorder=3)

    ax1.set_ylabel("Distance Moyenne (m)", color="#1F4E79", fontweight="bold", fontsize=8.5)
    ax2.set_ylabel("HSR Z3 Moyen (m)", color="#A30000", fontweight="bold", fontsize=8.5)
    ax1.set_xticks(x)
    ax1.set_xticklabels(groupes, fontweight="bold", fontsize=9.0, color="#1E293B")
    ax1.set_title("Charge Moyenne Hebdomadaire par Ligne", fontweight="bold", color="#1F4E79", fontsize=10.0, pad=8)

    # Annotations au-dessus des barres
    for rect in rects1:
        h = rect.get_height()
        if h > 0:
            ax1.annotate(f"{int(round(h)):,d} m".replace(",", " "),
                         xy=(rect.get_x() + rect.get_width() / 2, h),
                         xytext=(0, 2), textcoords="offset points",
                         ha="center", va="bottom", fontsize=7.5, fontweight="bold", color="#1F4E79")

    for rect in rects2:
        h = rect.get_height()
        if h > 0:
            ax2.annotate(f"{int(round(h)):,d} m".replace(",", " "),
                         xy=(rect.get_x() + rect.get_width() / 2, h),
                         xytext=(0, 2), textcoords="offset points",
                         ha="center", va="bottom", fontsize=7.5, fontweight="bold", color="#A30000")

    max_dist = max(dist_vals) if dist_vals and max(dist_vals) > 0 else 1000.0
    max_hsr = max(hsr_vals) if hsr_vals and max(hsr_vals) > 0 else 100.0
    ax1.set_ylim(0, max_dist * 1.25)
    ax2.set_ylim(0, max_hsr * 1.25)
    ax1.grid(axis="y", linestyle="--", alpha=0.3, zorder=1)
    ax1.tick_params(axis="y", labelcolor="#1F4E79", labelsize=8)
    ax2.tick_params(axis="y", labelcolor="#A30000", labelsize=8)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper center", bbox_to_anchor=(0.5, -0.12), ncol=2, frameon=False, fontsize=8.0)

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def generer_graphique_hebdo_repartition_jours_matplotlib(synthese_jours: Dict[str, Any]) -> io.BytesIO:
    """
    Génère le graphique comparatif de la répartition de la charge par séance (Mardi vs Mercredi vs Vendredi...)
    pour la Page 1 du PDF et le Dashboard.
    """
    fig, ax1 = plt.subplots(figsize=(6.2, 2.6), dpi=180)
    fig.patch.set_facecolor("#FFFFFF")
    ax1.set_facecolor("#FFFFFF")

    days_labels: List[str] = []
    dist_vals: List[float] = []
    hi_vals: List[float] = []

    for s_id, s in synthese_jours.items():
        if not isinstance(s, dict):
            continue
        raw_day = s.get("day_type", s_id)
        short_day = raw_day.split("(")[0].strip()
        date_clean = format_readable_date(s.get("date", ""))
        label = f"{short_day} ({date_clean})" if date_clean else short_day
        days_labels.append(label)
        dist_vals.append(float(s.get("total_distance_km", 0.0)))
        hi_vals.append(float(s.get("total_hi_m", 0.0)))

    if not days_labels:
        days_labels = ["Séance 1"]
        dist_vals = [0.0]
        hi_vals = [0.0]

    x = np.arange(len(days_labels))
    width = 0.35

    rects1 = ax1.bar(x - width/2, dist_vals, width, label="Distance Totale (km)", color="#1F4E79", edgecolor="none", zorder=3)
    ax2 = ax1.twinx()
    rects2 = ax2.bar(x + width/2, hi_vals, width, label="Haute Intensité (m)", color="#A30000", edgecolor="none", zorder=3)

    ax1.set_ylabel("Distance (km)", color="#1F4E79", fontweight="bold", fontsize=8.2)
    ax2.set_ylabel("H.I. (m)", color="#A30000", fontweight="bold", fontsize=8.2)
    ax1.set_xticks(x)
    ax1.set_xticklabels(days_labels, fontweight="bold", fontsize=8.2, color="#1E293B")
    ax1.set_title("Volume Journalier par Séance", fontweight="bold", color="#1F4E79", fontsize=9.5, pad=6)

    for rect in rects1:
        h = rect.get_height()
        if h > 0:
            ax1.annotate(f"{h:.1f} km",
                         xy=(rect.get_x() + rect.get_width() / 2, h),
                         xytext=(0, 2), textcoords="offset points",
                         ha="center", va="bottom", fontsize=7.2, fontweight="bold", color="#1F4E79")

    for rect in rects2:
        h = rect.get_height()
        if h > 0:
            ax2.annotate(f"{int(round(h)):,d} m".replace(",", " "),
                         xy=(rect.get_x() + rect.get_width() / 2, h),
                         xytext=(0, 2), textcoords="offset points",
                         ha="center", va="bottom", fontsize=7.2, fontweight="bold", color="#A30000")

    max_dist = max(dist_vals) if dist_vals and max(dist_vals) > 0 else 5.0
    max_hi = max(hi_vals) if hi_vals and max(hi_vals) > 0 else 500.0
    ax1.set_ylim(0, max_dist * 1.25)
    ax2.set_ylim(0, max_hi * 1.25)
    ax1.grid(axis="y", linestyle="--", alpha=0.3, zorder=1)
    ax1.tick_params(axis="y", labelcolor="#1F4E79", labelsize=7.5)
    ax2.tick_params(axis="y", labelcolor="#A30000", labelsize=7.5)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper center", bbox_to_anchor=(0.5, -0.16), ncol=2, frameon=False, fontsize=7.5)

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def generer_graphique_hebdo_jauges_matplotlib(
    group_name: str,
    df_grp: pd.DataFrame,
    nb_seances: int = 3,
    objectifs_config: Optional[Dict[str, Dict[str, float]]] = None,
) -> io.BytesIO:
    """
    Génère une figure 2x2 propre regroupant les 4 duels de barres 'Objectif Hebdomadaire vs Réalisé Cumulé'
    pour une ligne de poste RCA (5 de devant, 3ème Ligne, Trois-quarts) :
    1. Distance Totale (km ou m)
    2. HSR Z3 (m)
    3. VHSR Sprint (m)
    4. ACC / DEC (> 3 m/s²)
    """
    active_cfg = objectifs_config or charger_objectifs_config()
    targets_daily = active_cfg.get(group_name, active_cfg.get("Tous les groupes", DEFAULT_OBJECTIFS_CONFIG["Tous les groupes"]))

    multiplier = max(nb_seances, 1)

    # Calcul des réalisés moyens par joueur du groupe sur la semaine
    if not df_grp.empty:
        val_dist = float(df_grp["Distance Totale (m)"].mean()) if "Distance Totale (m)" in df_grp.columns else (float(df_grp["Distance Totale (km)"].mean()) * 1000.0 if "Distance Totale (km)" in df_grp.columns else 0.0)
        val_hsr = float(df_grp["HSR Z3 (m)"].mean()) if "HSR Z3 (m)" in df_grp.columns else (float(df_grp["Haute intensité (m)"].mean()) if "Haute intensité (m)" in df_grp.columns else 0.0)
        val_vhsr = float(df_grp["Distance Sprint VHSR (m)"].mean()) if "Distance Sprint VHSR (m)" in df_grp.columns else (float(df_grp["VHSR (m)"].mean()) if "VHSR (m)" in df_grp.columns else 0.0)
        val_acc = float(df_grp["Accel/Dec total"].mean()) if "Accel/Dec total" in df_grp.columns else (float(df_grp["Accel/Dec"].mean()) if "Accel/Dec" in df_grp.columns else 0.0)
    else:
        val_dist = 0.0
        val_hsr = 0.0
        val_vhsr = 0.0
        val_acc = 0.0

    cards_def = [
        ("Distance Totale (m)", targets_daily["Objectif Distance (m)"] * multiplier, val_dist, "m"),
        ("HSR Z3 (m)", targets_daily["Obj. HSR"] * multiplier, val_hsr, "m"),
        ("VHSR Sprint (m)", targets_daily["Obj. VHSR"] * multiplier, val_vhsr, "m"),
        ("ACC / DEC (> 3 m/s²)", targets_daily["Obj. ACC/DEC"] * multiplier, val_acc, ""),
    ]

    fig, axes = plt.subplots(1, 4, figsize=(11.5, 2.5), dpi=180)
    fig.patch.set_facecolor("#FFFFFF")

    for idx, (title, obj_val, real_val, unit) in enumerate(cards_def):
        ax = axes[idx]
        ax.set_facecolor("#F8FAFC")
        for spine in ax.spines.values():
            spine.set_color("#CBD5E1")
            spine.set_linewidth(0.8)

        pct = (real_val / obj_val * 100.0) if obj_val > 0 else 0.0

        x = np.array([0, 1])
        width = 0.44

        ax.bar(
            x,
            [obj_val, real_val],
            width=width,
            color=["#1F4E79", "#A30000"],
            edgecolor="none",
            zorder=3,
        )

        ax.set_title(title, fontsize=9.2, fontweight="bold", color="#1F4E79", pad=6)
        ax.set_xticks(x)
        ax.set_xticklabels(["Obj.", "Réalisé"], fontsize=8.5, fontweight="bold", color="#334155")

        max_v = max(obj_val, real_val, 1.0)
        ax.set_ylim(0, max_v * 1.35)
        ax.grid(axis="y", linestyle="--", alpha=0.35, color="#94A3B8", zorder=1)
        ax.set_yticklabels([])
        ax.tick_params(axis="both", which="both", length=0)

        def _fmt(v: float, u: str) -> str:
            if abs(v - round(v)) < 0.05:
                s = f"{int(round(v)):,d}".replace(",", " ")
            else:
                s = f"{v:.1f}"
            return f"{s} {u}".strip()

        ax.text(
            0, obj_val + (max_v * 0.03),
            _fmt(obj_val, unit),
            ha="center", va="bottom",
            fontsize=8.0, fontweight="bold", color="#1F4E79"
        )

        ax.text(
            1, real_val + (max_v * 0.03),
            _fmt(real_val, unit),
            ha="center", va="bottom",
            fontsize=8.0, fontweight="bold", color="#A30000"
        )

        # Badge % d'atteinte
        badge_bg = "#DCFCE7" if pct >= 95.0 else ("#FEF3C7" if pct >= 75.0 else "#FEE2E2")
        badge_fg = "#166534" if pct >= 95.0 else ("#92400E" if pct >= 75.0 else "#991B1B")

        ax.text(
            0.5, 0.88,
            f"{pct:.0f}%",
            transform=ax.transAxes,
            ha="center", va="center",
            fontsize=8.0, fontweight="bold",
            color=badge_fg,
            bbox=dict(boxstyle="round,pad=0.25", facecolor=badge_bg, edgecolor="none", alpha=0.95),
            zorder=4,
        )

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


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
        self.current_subtitle = ""
        self.set_auto_page_break(auto=False)
        self.set_margins(left=12, top=20, right=12)

    def header(self):
        # Bandeau officiel RCA
        self.set_fill_color(11, 25, 44)  # #0B192C
        self.rect(0, 0, 297, 18, "F")
        self.set_fill_color(212, 175, 55)  # Or #D4AF37
        self.rect(0, 18, 297, 1.2, "F")

        # Titre
        self.set_font("Helvetica", "B", 12)
        self.set_text_color(255, 255, 255)
        self.set_xy(12, 3)
        self.cell(160, 6, safe_latin1("RUGBY CLUB AMIÉNOIS  |  RAPPORT HEBDOMADAIRE DE CHARGE GPS"), align="L")

        if self.current_subtitle:
            self.set_xy(12, 9)
            self.set_font("Helvetica", "B", 9)
            self.set_text_color(212, 175, 55)
            self.cell(160, 5, safe_latin1(self.current_subtitle), align="L")

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
        self.cell(133, 6, safe_latin1(f"Page {self.page_no()}/4"), align="R")


def generate_weekly_pdf_report(
    df_hebdo: Optional[pd.DataFrame] = None,
    year: int = 2026,
    week_num: int = 1,
    synthese_jours: Optional[Union[Dict[str, Any], List[Dict[str, Any]]]] = None,
    notes_coach: str = "",
    objectifs_config: Optional[Dict[str, Dict[str, float]]] = None,
    **kwargs: Any,
) -> bytes:
    """Génère un rapport PDF multi-pages enrichi de graphiques de performance sur 4 pages au format Paysage."""
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

    week_label = f"Semaine {week_num:02d} - {year}"
    raw_dates = [format_readable_date(s.get("date", "")) for s in synthese_dict.values() if isinstance(s, dict) and s.get("date")]
    raw_dates = [d for d in raw_dates if d]
    dates_str = f"Du {raw_dates[0]} au {raw_dates[-1]}" if len(raw_dates) >= 2 else (raw_dates[0] if raw_dates else "")

    pdf = RCAWeeklyPDFReport(week_label=week_label, dates_label=dates_str)

    # --------------------------------------------------------------------------
    # PAGE 1 : SYNTHÈSE GLOBALE DE LA SEMAINE & GRAPHIQUES DE CHARGE
    # --------------------------------------------------------------------------
    pdf.current_subtitle = "SYNTHÈSE GLOBALE & RÉPARTITION DE LA CHARGE"
    pdf.add_page()
    pdf.set_y(22)

    # KPIs Globaux
    tot_km = float(df_hebdo["Distance Totale (km)"].sum()) if not df_hebdo.empty and "Distance Totale (km)" in df_hebdo.columns else 0.0
    tot_hi = float(df_hebdo["Haute intensité (m)"].sum()) if not df_hebdo.empty and "Haute intensité (m)" in df_hebdo.columns else 0.0
    nb_joueurs_tot = len(df_hebdo)
    max_speed = float(df_hebdo["Top Speed (km/h)"].max()) if not df_hebdo.empty and "Top Speed (km/h)" in df_hebdo.columns else 0.0
    nb_seances = len(synthese_dict)

    kpis = [
        ("Volume Cumulé Équipe", f"{tot_km:.1f} km"),
        ("Haute Intensité Cumulée", f"{int(round(tot_hi)):,d} m".replace(",", " ")),
        ("Joueurs Actifs Semaine", f"{nb_joueurs_tot} joueurs"),
        ("Vitesse Max Semaine", f"{max_speed:.1f} km/h"),
        ("Séances Fusionnées", f"{nb_seances} séances"),
    ]

    col_w = 52.8
    for idx, (label, val) in enumerate(kpis):
        x = 12 + idx * (col_w + 2.2)
        pdf.set_fill_color(248, 250, 252)
        pdf.set_draw_color(226, 232, 240)
        pdf.rect(x, 22, col_w, 17, "FD")
        pdf.set_xy(x, 23.5)
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.set_text_color(100, 116, 139)
        pdf.cell(col_w, 4, label.upper(), align="C")
        pdf.set_xy(x, 28.5)
        pdf.set_font("Helvetica", "B", 11.5)
        pdf.set_text_color(31, 78, 121)
        pdf.cell(col_w, 8, val, align="C")

    # Section intermédiaire (X gauche = 12, X droite = 151)
    y_mid = 42.0

    # Colonne Gauche : Tableau des séances + Graphique Répartition de charge
    pdf.set_xy(12, y_mid)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.set_text_color(11, 25, 44)
    pdf.cell(133, 5, "1. SÉANCES ENREGISTRÉES & RÉPARTITION DE CHARGE :", align="L")

    # Petit tableau des séances (width = 133)
    y_table = y_mid + 6.0
    pdf.set_xy(12, y_table)
    pdf.set_fill_color(31, 78, 121)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 8)
    pdf.cell(45, 5.5, "Séance / Type", 1, 0, "L", fill=True)
    pdf.cell(24, 5.5, "Date", 1, 0, "C", fill=True)
    pdf.cell(20, 5.5, "Effectif", 1, 0, "C", fill=True)
    pdf.cell(22, 5.5, "Dist. (km)", 1, 0, "C", fill=True)
    pdf.cell(22, 5.5, "H.I. (m)", 1, 1, "C", fill=True)

    pdf.set_font("Helvetica", "", 7.5)
    pdf.set_text_color(15, 23, 42)
    for s_id, s in synthese_dict.items():
        pdf.set_x(12)
        pdf.cell(45, 5.0, safe_latin1(str(s.get("day_type", ""))[:20]), 1, 0, "L")
        pdf.cell(24, 5.0, safe_latin1(format_readable_date(s.get("date", ""))), 1, 0, "C")
        pdf.cell(20, 5.0, safe_latin1(f"{s.get('num_players', 0)}"), 1, 0, "C")
        pdf.cell(22, 5.0, safe_latin1(f"{float(s.get('total_distance_km', 0.0)):.1f} km"), 1, 0, "C")
        pdf.cell(22, 5.0, safe_latin1(f"{int(round(float(s.get('total_hi_m', 0.0)))):,d}".replace(",", " ")), 1, 1, "C")

    # Graphique Répartition de charge par séance sous le tableau
    y_chart_days = pdf.get_y() + 2.0
    try:
        buf_days = generer_graphique_hebdo_repartition_jours_matplotlib(synthese_dict)
        chart_days_h = min(40.0, 118.0 - y_chart_days)
        if chart_days_h > 20:
            pdf.image(buf_days, x=12, y=y_chart_days, w=133, h=chart_days_h)
    except Exception as e:
        print(f"Erreur graphique repartition jours PDF : {e}")

    # Colonne Droite : Graphique Charge moyenne par Ligne
    pdf.set_xy(151, y_mid)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.set_text_color(11, 25, 44)
    pdf.cell(133, 5, "2. CHARGE MOYENNE PAR LIGNE DE POSTE (SEMAINE) :", align="L")

    try:
        buf_groups = generer_graphique_hebdo_groupes_matplotlib(df_hebdo)
        pdf.image(buf_groups, x=151, y=y_mid + 6.0, w=134, h=66.0)
    except Exception as e:
        print(f"Erreur graphique groupes hebdo PDF : {e}")

    # 3. Podiums Hebdomadaires en bas de Page 1
    y_pod = 120.0
    pdf.set_xy(12, y_pod)
    pdf.set_font("Helvetica", "B", 10.0)
    pdf.set_text_color(11, 25, 44)
    pdf.cell(273, 5, "3. PODIUMS HEBDOMADAIRES (CUMULS DE LA SEMAINE) :", align="L")

    podiums = generer_podiums_hebdomadaires(df_hebdo)
    cat_keys = ["Distance Totale (km)", "Top Speed (km/h)", "Haute Intensité (m)"]
    p_col_w = 88.5

    for idx, c_key in enumerate(cat_keys):
        c_data = podiums.get(c_key, {})
        players = c_data.get("players", [])
        px = 12 + idx * (p_col_w + 3.5)
        py = y_pod + 6.5

        # Card header
        pdf.set_xy(px, py)
        pdf.set_fill_color(11, 25, 44)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(p_col_w, 6.0, safe_latin1(f"{c_data.get('icon', '')} {c_key.upper()}"), 1, 1, "C", fill=True)

        for p in players:
            rank = p.get("rank", 1)
            p_name = p.get("name", "-")
            p_stat = p.get("stat", "-")
            p_pos = p.get("poste", "")

            bg_color = (255, 243, 199) if rank == 1 else ((241, 245, 249) if rank == 2 else (255, 237, 213))
            pdf.set_fill_color(*bg_color)
            pdf.set_draw_color(226, 232, 240)
            cur_y = pdf.get_y()
            pdf.set_x(px)
            pdf.rect(px, cur_y, p_col_w, 14.5, "FD")

            # Photo du joueur
            photo_path = photo_manager.get_player_photo_path(p_name)
            has_photo = bool(photo_path and os.path.isfile(photo_path) and os.path.basename(photo_path) != "avatar_par_defaut.png")
            if has_photo:
                try:
                    pdf.image(photo_path, x=px + 2, y=cur_y + 1.2, w=12, h=12)
                except Exception:
                    has_photo = False

            text_x = px + (16 if has_photo else 4)
            pdf.set_xy(text_x, cur_y + 1.5)
            pdf.set_font("Helvetica", "B", 8.5)
            pdf.set_text_color(15, 23, 42)
            pdf.cell(p_col_w - 18, 3.8, safe_latin1(f"#{rank} {p_name}"), align="L")

            pdf.set_xy(text_x, cur_y + 5.5)
            pdf.set_font("Helvetica", "I", 7.0)
            pdf.set_text_color(100, 116, 139)
            pdf.cell(p_col_w - 18, 3.2, safe_latin1(f"Poste : {p_pos}"), align="L")

            pdf.set_xy(text_x, cur_y + 9.2)
            pdf.set_font("Helvetica", "B", 9.0)
            pdf.set_text_color(163, 0, 0 if rank == 1 else 31)
            pdf.cell(p_col_w - 18, 3.8, safe_latin1(p_stat), align="L")

            pdf.set_y(cur_y + 14.5)

    # --------------------------------------------------------------------------
    # PAGES 2, 3 ET 4 : BILAN DÉTAILLÉ PAR LIGNE DE POSTE (AVEC GRAPHIQUES 2x2)
    # --------------------------------------------------------------------------
    groupes_rca = ["5 de devant", "3ème Ligne", "Trois-quarts"]

    for grp in groupes_rca:
        df_sub = df_hebdo[df_hebdo["Groupe"] == grp] if not df_hebdo.empty and "Groupe" in df_hebdo.columns else pd.DataFrame()
        if df_sub.empty:
            continue

        pdf.current_subtitle = f"BILAN HEBDOMADAIRE — {grp.upper()} ({len(df_sub)} JOUEURS)"
        pdf.add_page()
        pdf.set_y(22)

        # 1. Graphiques en barres Objectif Hebdomadaire vs Réalisé Cumulé
        pdf.set_xy(12, 22)
        pdf.set_font("Helvetica", "B", 10.0)
        pdf.set_text_color(11, 25, 44)
        pdf.cell(273, 5, safe_latin1(f"1. OBJECTIFS HEBDOMADAIRES VS RÉALISÉ DU GROUPE ({nb_seances} SÉANCES CUMULÉES) :"), align="L")

        try:
            buf_jauges = generer_graphique_hebdo_jauges_matplotlib(
                grp, df_sub, nb_seances=nb_seances, objectifs_config=objectifs_config
            )
            pdf.image(buf_jauges, x=12, y=28, w=273, h=56)
        except Exception as e:
            print(f"Erreur graphique jauges hebdo PDF pour {grp} : {e}")

        # 2. Tableau individuel des joueurs
        y_tbl_grp = 88.0
        pdf.set_xy(12, y_tbl_grp)
        pdf.set_font("Helvetica", "B", 10.0)
        pdf.set_text_color(11, 25, 44)
        pdf.cell(273, 5, safe_latin1(f"2. STATISTIQUES INDIVIDUELLES CUMULÉES SUR LA SEMAINE :"), align="L")

        # Table header
        pdf.set_xy(12, y_tbl_grp + 6.0)
        pdf.set_fill_color(31, 78, 121)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 8)

        cols_def = [
            ("Joueur", 46, "L"),
            ("Poste", 26, "C"),
            ("Présences", 30, "C"),
            ("Temps (min)", 22, "C"),
            ("Dist. (km)", 24, "C"),
            ("H.I. (m)", 26, "C"),
            ("Sprint (m)", 24, "C"),
            ("Vmax (km/h)", 25, "C"),
            ("m/min", 24, "C"),
            ("Acc/Dec", 26, "C"),
        ]

        for title, w, align in cols_def:
            pdf.cell(w, 6.5, safe_latin1(title), 1, 0, align, fill=True)
        pdf.ln(6.5)

        pdf.set_font("Helvetica", "", 8)
        pdf.set_text_color(15, 23, 42)
        fill_row = False

        for _, row in df_sub.iterrows():
            pdf.set_x(12)
            bg = (248, 250, 252) if fill_row else (255, 255, 255)
            pdf.set_fill_color(*bg)

            p_nom = str(row["Nom ; Prénom"])[:22]
            p_pos = str(row.get("Poste_Detaille", "RCA"))[:13]
            p_pres = f"{row.get('Nombre de séances', 1)} séance{'s' if row.get('Nombre de séances', 1) > 1 else ''}"
            p_tps = f"{float(row.get('Temps total (min)', 0.0)):.0f}"
            p_dist = f"{float(row.get('Distance Totale (km)', 0.0)):.2f}"
            p_hi = f"{int(round(float(row.get('Haute intensité (m)', 0.0)))):,d}".replace(",", " ")
            p_vhsr = f"{int(round(float(row.get('Distance Sprint VHSR (m)', 0.0)))):,d}".replace(",", " ")
            p_vmax = f"{float(row.get('Top Speed (km/h)', 0.0)):.1f}"
            p_dpm = f"{float(row.get('Distance/min (m/min)', 0.0)):.1f}"
            p_acc = f"{int(round(float(row.get('Accel/Dec total', 0.0))))}"

            pdf.cell(46, 5.8, safe_latin1(p_nom), 1, 0, "L", fill=True)
            pdf.cell(26, 5.8, safe_latin1(p_pos), 1, 0, "C", fill=True)
            pdf.cell(30, 5.8, safe_latin1(p_pres), 1, 0, "C", fill=True)
            pdf.cell(22, 5.8, safe_latin1(p_tps), 1, 0, "C", fill=True)
            pdf.cell(24, 5.8, safe_latin1(p_dist), 1, 0, "C", fill=True)
            pdf.cell(26, 5.8, safe_latin1(p_hi), 1, 0, "C", fill=True)
            pdf.cell(24, 5.8, safe_latin1(p_vhsr), 1, 0, "C", fill=True)
            pdf.cell(25, 5.8, safe_latin1(p_vmax), 1, 0, "C", fill=True)
            pdf.cell(24, 5.8, safe_latin1(p_dpm), 1, 0, "C", fill=True)
            pdf.cell(26, 5.8, safe_latin1(p_acc), 1, 1, "C", fill=True)
            fill_row = not fill_row

    return bytes(pdf.output())
