"""
Module de Connexion & d'Intégration RPE (Google Sheets)
Club : Rugby Club Amiénois (RCA) - Catapult GPS Processor
Description :
    Connecte en lecture seule l'application de traitement GPS au Google Sheet
    de l'application de suivi de charge RPE (suivi-charge-rca).
    Permet d'enrichir automatiquement :
    1. Le Rapport Journalier (séance du jour) avec le RPE moyen par groupe et équipe.
    2. Le Rapport Hebdomadaire (cumul de la semaine et historique multi-semaines).

Règles de jointure & robustesse :
    - Clé de jointure : Nom + Prénom + Date (nettoyage, minuscules, suppression des accents et espaces).
    - Attribution des postes : Exclusivement le tableau GPS Catapult (les postes du sheet RPE sont ignorés).
    - Gestion des valeurs manquantes : 'NR' (Non Renseigné) pour l'individuel, et ignorées (skipna=True) pour les moyennes.
    - Mise en cache Streamlit : @st.cache_data(ttl=300) pour préserver les quotas Google Sheets API (429 Rate Limit).
"""

from __future__ import annotations

import datetime
import os
import re
import unicodedata
from typing import Any, Dict, List, Optional, Tuple, Union

import numpy as np
import pandas as pd
import streamlit as st

# Scopes Google Sheets en lecture seule (sécurité maximale)
GSHEET_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Noms courants de la feuille de calcul et des onglets de réponses RPE
DEFAULT_SPREADSHEET_NAME = "suivi-charge-rca"
COMMON_WORKSHEET_NAMES = [
    "Saisies",
    "Form Responses 1",
    "Réponses au formulaire 1",
    "RPE",
    "Charge",
    "Données",
    "Donnees",
    "Sheet1",
]


# ==============================================================================
# 1. NORMALISATION DU TEXTE & DES DATES (CORRESPONDANCE ULTRA-ROBUSTE)
# ==============================================================================

def normalize_text_key(text: Any) -> str:
    """
    Normalise une chaîne de texte pour la comparaison :
    - Nettoyage des parenthèses et préfixes boîtiers (ex: 'Jean-Étienne ( TIJ)' -> 'Jean-Étienne')
    - Suppression totale des accents (NFKD)
    - Remplacement des tirets, apostrophes, ponctuations par des espaces
    - Passage en minuscules et suppression des espaces superflus début/milieu/fin
    """
    if text is None or pd.isna(text):
        return ""
    s = str(text).strip()
    # Supprime le contenu entre parenthèses (diminutifs, surnoms comme TIJ, annotations)
    s = re.sub(r"\(.*?\)", " ", s)
    s = s.strip().lower()
    # Décomposition Unicode pour éliminer les accents (é, è, ê, ï, à, etc.)
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    # Remplacement des caractères non-alphanumériques par des espaces
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return " ".join(s.split())


# Alias et variantes orthographiques spécifiques à l'effectif sénior RCA
RCA_NAME_ALIASES: Dict[str, str] = {
    "tij": "lesueur jean etienne",
    "jean etienne tij": "lesueur jean etienne",
    "jean etienne": "lesueur jean etienne",
    "gabin d": "denier gabin",
    "d anthymme": "demonchy anthymme",
    "anthymme d": "demonchy anthymme",
    "vdh simon": "van der heijden simon",
    "simon vdh": "van der heijden simon",
}


def are_player_names_matching(name_gps: str, name_rpe: str) -> bool:
    """
    Vérifie si le nom du joueur GPS (issu du mapping de séance) correspond au nom
    saisi dans le Google Sheets RPE :
    1. Égalité stricte après normalisation textuelle (sans accents, sans ponctuation, minuscules).
    2. Correspondance canonique d'alias RCA (ex: 'Jean-Étienne ( TIJ)' <-> 'LESUEUR Jean-Étienne').
    3. Permutation libre des tokens (ex: 'DUPONT Antoine' == 'Antoine DUPONT').
    4. Inclusion mutuelle de tokens (ex: 'PECQUERY Thomas' vs 'Pecquery').
    """
    norm_g = normalize_text_key(name_gps)
    norm_r = normalize_text_key(name_rpe)
    if not norm_g or not norm_r:
        return False

    # 1. Égalité directe
    if norm_g == norm_r:
        return True

    # 2. Alias canoniques RCA
    canon_g = RCA_NAME_ALIASES.get(norm_g, norm_g)
    canon_r = RCA_NAME_ALIASES.get(norm_r, norm_r)
    if canon_g == canon_r:
        return True

    # 3. Correspondance par ensemble de tokens (tolérance totale inversion Nom/Prénom)
    tokens_g = set(norm_g.split())
    tokens_r = set(norm_r.split())
    if tokens_g and tokens_r and tokens_g == tokens_r:
        return True

    # 4. Inclusion si un prénom ou un nom seul est renseigné
    if tokens_g and tokens_r and (tokens_g.issubset(tokens_r) or tokens_r.issubset(tokens_g)):
        return True

    # 5. Tokens sur noms canonisés
    tokens_cg = set(canon_g.split())
    tokens_cr = set(canon_r.split())
    if tokens_cg and tokens_cr and (tokens_cg == tokens_cr or tokens_cg.issubset(tokens_cr) or tokens_cr.issubset(tokens_cg)):
        return True

    return False


def parse_and_normalize_date(date_val: Any) -> Optional[datetime.date]:
    """
    Convertit n'importe quel format de date (ISO, français, datetime, timestamp, float Excel)
    en un objet datetime.date standardisé.
    """
    if date_val is None or pd.isna(date_val):
        return None

    if isinstance(date_val, datetime.datetime):
        return date_val.date()
    if isinstance(date_val, datetime.date):
        return date_val

    # Si c'est un entier ou float Excel (ex: 46332)
    if isinstance(date_val, (int, float)):
        try:
            return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(date_val)))
        except Exception:
            pass

    s = str(date_val).strip()
    if not s or s.lower() in ["none", "nan", "nat", "-"]:
        return None

    # Si c'est un numéro de série Excel sous forme de chaîne ('46332' ou '46332.0')
    try:
        f_val = float(s)
        if 40000 <= f_val <= 60000:
            return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(f_val)))
    except Exception:
        pass

    # Extraction de la partie date (supprime l'heure si '06/11/2026 19:30:00')
    s_part = s.split(" ")[0].strip()

    # 1. Format JJ/MM/AAAA ou JJ-MM-AAAA
    m_fr = re.match(r"^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$", s_part)
    if m_fr:
        d, m, y = int(m_fr.group(1)), int(m_fr.group(2)), int(m_fr.group(3))
        if y < 100:
            y += 2000
        try:
            return datetime.date(y, m, d)
        except ValueError:
            pass

    # 2. Format ISO AAAA-MM-JJ
    m_iso = re.match(r"^(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})$", s_part)
    if m_iso:
        y, m, d = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
        try:
            return datetime.date(y, m, d)
        except ValueError:
            pass

    # 3. Format textuel avec nom de jour en préfixe (ex: 'Vendredi 06/11/2026')
    m_day_pref = re.search(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})", s)
    if m_day_pref:
        d, m, y = int(m_day_pref.group(1)), int(m_day_pref.group(2)), int(m_day_pref.group(3))
        if y < 100:
            y += 2000
        try:
            return datetime.date(y, m, d)
        except ValueError:
            pass

    # Repli avec pd.to_datetime
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.date()
    except Exception:
        pass

    return None


def parse_rpe_value(val: Any) -> Optional[float]:
    """
    Extrait la valeur numérique du RPE (échelle de Foster 1 à 10).
    Supporte '7', '7.5', '7/10', '7 - Très difficile', etc.
    Retourne None si la valeur est absente ou invalide.
    """
    if val is None or pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s.upper() in ["NR", "NON RENSEIGNÉ", "NON RENSEIGNE", "-", "NAN", "NULL"]:
        return None

    # Recherche du premier nombre dans la chaîne
    m = re.match(r"^(\d+(?:[.,]\d+)?)", s)
    if m:
        try:
            v = float(m.group(1).replace(",", "."))
            if 0.0 <= v <= 10.0:
                return v
        except Exception:
            pass
    return None


def parse_duration_minutes(val: Any, default: float = 80.0) -> float:
    """Extrait la durée de séance en minutes (par défaut 80 min)."""
    if val is None or pd.isna(val):
        return default
    s = str(val).strip()
    m = re.search(r"(\d+(?:[.,]\d+)?)", s)
    if m:
        try:
            v = float(m.group(1).replace(",", "."))
            if 10.0 <= v <= 240.0:
                return v
        except Exception:
            pass
    return default


# ==============================================================================
# 2. CONNEXION GOOGLE SHEETS & CACHE STREAMLIT (@st.cache_data(ttl=300))
# ==============================================================================

def _get_secret_safely(key: str, default: Any = None) -> Any:
    """Accède aux secrets Streamlit sans lever d'exception en environnement de test."""
    try:
        if hasattr(st, "secrets") and key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return default


def get_gcp_service_account_credentials() -> Optional[Any]:
    """
    Résout les identifiants du compte de service Google Cloud :
    1. Depuis st.secrets["gcp_service_account"] (dict ou table TOML)
    2. Depuis st.secrets["connections"]["gsheets"]
    3. Depuis la variable d'environnement GCP_SERVICE_ACCOUNT (chemin fichier ou JSON)
    """
    try:
        from google.oauth2.service_account import Credentials
    except ImportError:
        return None

    # 1. Vérification dans st.secrets["gcp_service_account"]
    gcp_dict = _get_secret_safely("gcp_service_account")
    if gcp_dict:
        try:
            sec_dict = dict(gcp_dict)
            if "private_key" in sec_dict and "\\n" in sec_dict["private_key"]:
                sec_dict["private_key"] = sec_dict["private_key"].replace("\\n", "\n")
            return Credentials.from_service_account_info(sec_dict, scopes=GSHEET_SCOPES)
        except Exception as e:
            print(f"Erreur parsing st.secrets['gcp_service_account'] : {e}")

    # 2. Vérification dans st.secrets["connections"]["gsheets"]
    conn_sec = _get_secret_safely("connections")
    if isinstance(conn_sec, dict) and "gsheets" in conn_sec:
        gsh_sec = conn_sec["gsheets"]
        if isinstance(gsh_sec, dict):
            # Certains schémas incluent les clés de service account directement
            if "type" in gsh_sec and "private_key" in gsh_sec:
                try:
                    s_copy = dict(gsh_sec)
                    if "\\n" in s_copy.get("private_key", ""):
                        s_copy["private_key"] = s_copy["private_key"].replace("\\n", "\n")
                    return Credentials.from_service_account_info(s_copy, scopes=GSHEET_SCOPES)
                except Exception:
                    pass

    # 3. Variable d'environnement GCP_SERVICE_ACCOUNT
    env_sa = os.environ.get("GCP_SERVICE_ACCOUNT", "").strip()
    if env_sa:
        if os.path.exists(env_sa):
            try:
                return Credentials.from_service_account_file(env_sa, scopes=GSHEET_SCOPES)
            except Exception:
                pass
        elif env_sa.startswith("{"):
            try:
                import json
                sa_data = json.loads(env_sa)
                return Credentials.from_service_account_info(sa_data, scopes=GSHEET_SCOPES)
            except Exception:
                pass

    return None


def get_spreadsheet_target() -> str:
    """Retourne l'URL, l'ID ou le nom du Google Sheet RPE."""
    # Secrets explicites pour l'app RPE
    for k in ["RPE_GSHEET_URL", "RPE_SPREADSHEET_URL", "RPE_SPREADSHEET", "rpe_spreadsheet"]:
        v = _get_secret_safely(k)
        if v:
            return str(v).strip()

    # Secret Streamlit Connections
    conn = _get_secret_safely("connections")
    if isinstance(conn, dict) and "gsheets" in conn:
        gsh = conn["gsheets"]
        if isinstance(gsh, dict) and "spreadsheet" in gsh:
            return str(gsh["spreadsheet"]).strip()

    gsheet_direct = _get_secret_safely("gsheets")
    if isinstance(gsheet_direct, dict) and "spreadsheet" in gsheet_direct:
        return str(gsheet_direct["spreadsheet"]).strip()

    return os.environ.get("RPE_SPREADSHEET_URL", DEFAULT_SPREADSHEET_NAME)


@st.cache_data(ttl=300)
def fetch_raw_rpe_data() -> pd.DataFrame:
    """
    Récupère en lecture seule les saisies brutes de charge RPE depuis le Google Sheet.
    Mise en cache Streamlit de 300 secondes (5 minutes) pour respecter strictement
    les quotas de requêtes de l'API Google (prévention des erreurs HTTP 429).
    """
    creds = get_gcp_service_account_credentials()
    sheet_target = get_spreadsheet_target()

    if creds is not None:
        try:
            import gspread
            client = gspread.authorize(creds)

            # Ouverture par URL, par clé ou par nom
            if sheet_target.startswith("https://"):
                spreadsheet = client.open_by_url(sheet_target)
            elif len(sheet_target) > 30 and "/" not in sheet_target:
                spreadsheet = client.open_by_key(sheet_target)
            else:
                try:
                    spreadsheet = client.open(sheet_target)
                except Exception:
                    spreadsheet = client.open_by_key(sheet_target)

            # Recherche de l'onglet cible
            target_ws_name = _get_secret_safely("RPE_WORKSHEET", "")
            worksheet = None

            if target_ws_name:
                try:
                    worksheet = spreadsheet.worksheet(target_ws_name)
                except Exception:
                    worksheet = None

            if worksheet is None:
                all_ws_names = [ws.title for ws in spreadsheet.worksheets()]
                for candidate in COMMON_WORKSHEET_NAMES:
                    if candidate in all_ws_names:
                        worksheet = spreadsheet.worksheet(candidate)
                        break

            if worksheet is None:
                worksheet = spreadsheet.sheet1

            records = worksheet.get_all_records()
            if records:
                df = pd.DataFrame(records)
                return df
        except Exception as e:
            print(f"Avertissement : Connexion Google Sheet RPE échouée ({e}).")

    # Fallback local : vérification d'un fichier de cache local
    cache_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "rpe_cache.csv")
    if os.path.exists(cache_path):
        try:
            return pd.read_csv(cache_path, encoding="utf-8")
        except Exception:
            pass

    return pd.DataFrame()


def get_rpe_sync_status() -> Tuple[bool, str]:
    """Retourne l'état de connexion de l'intégration Google Sheets RPE."""
    creds = get_gcp_service_account_credentials()
    if creds is None:
        return False, "Identifiants Google Cloud non configurés (st.secrets['gcp_service_account']). Mode hors-ligne actif."
    target = get_spreadsheet_target()
    return True, f"Connecté au Google Sheet RPE : {target}"


# ==============================================================================
# 3. IDENTIFICATION DES COLONNES DU SHEET RPE
# ==============================================================================

def find_column_by_aliases(
    df: pd.DataFrame,
    aliases: List[str],
    exclude: Optional[List[str]] = None,
) -> Optional[str]:
    """
    Trouve la colonne correspondante dans un DataFrame selon une liste d'alias.
    Priorité absolue à la correspondance exacte normalisée pour éviter les faux positifs (ex: 'Nom' dans 'Nom et Prénom').
    """
    if df is None or df.empty:
        return None
    excl_set = {normalize_text_key(e) for e in exclude if e} if exclude else set()

    # Passe 1 : Correspondance EXACTE prioritaire
    for a in aliases:
        a_clean = normalize_text_key(a)
        for col in df.columns:
            col_clean = normalize_text_key(col)
            if col_clean in excl_set:
                continue
            if a_clean == col_clean:
                return col

    # Passe 2 : Correspondance par mot complet délimité
    for a in aliases:
        a_clean = normalize_text_key(a)
        pattern = r"\b" + re.escape(a_clean) + r"\b"
        for col in df.columns:
            col_clean = normalize_text_key(col)
            if col_clean in excl_set:
                continue
            if re.search(pattern, col_clean):
                return col

    return None


def identify_rpe_columns(df_rpe: pd.DataFrame) -> Dict[str, Optional[str]]:
    """Détecte les colonnes de date, nom, prénom, RPE et durée dans le DataFrame RPE."""
    # 1. Colonne Date : Priorité à la date de séance explicite avant l'horodateur de formulaire
    col_date = find_column_by_aliases(df_rpe, [
        "Date de la séance", "Date de seance", "Date séance", "Date seance",
        "Date de l'entraînement", "Date entraînement", "Date entrainement", "Date", "Horodateur", "Timestamp"
    ])

    # 2. Colonne RPE
    col_rpe = find_column_by_aliases(df_rpe, [
        "RPE", "Score RPE", "RPE (1-10)", "Note RPE", "Difficulté de la séance", "Difficulté", "Intensité", "Intensite"
    ], exclude=[col_date] if col_date else None)

    # 3. Colonne Durée
    col_duration = find_column_by_aliases(df_rpe, [
        "Durée (min)", "Durée de la séance", "Duree (min)", "Durée", "Duree", "Duration", "Temps"
    ], exclude=[col_date, col_rpe] if (col_date or col_rpe) else None)

    # 4. Colonne Joueur unique (Nom complet ou Nom Prénom)
    col_joueur_unique = find_column_by_aliases(df_rpe, [
        "Nom et Prénom", "Nom & Prénom", "Nom Prénom", "Nom ; Prénom", "Nom complet",
        "Joueur", "Athlète", "Athlete", "Player Name"
    ], exclude=[col_date, col_rpe, col_duration] if (col_date or col_rpe or col_duration) else None)

    # 5. Colonnes Nom et Prénom séparées
    excl_names = [c for c in [col_date, col_rpe, col_duration, col_joueur_unique] if c]
    col_nom = find_column_by_aliases(df_rpe, ["Nom de famille", "Nom"], exclude=excl_names)
    excl_prenom = excl_names + ([col_nom] if col_nom else [])
    col_prenom = find_column_by_aliases(df_rpe, ["Prénom", "Prenom"], exclude=excl_prenom)

    return {
        "col_date": col_date,
        "col_nom": col_nom,
        "col_prenom": col_prenom,
        "col_joueur_unique": col_joueur_unique,
        "col_rpe": col_rpe,
        "col_duration": col_duration,
    }


# ==============================================================================
# 4. FUSION GPS & RPE (JOINTURE SUR NOM + PRÉNOM + DATE)
# ==============================================================================

def enrich_gps_dataframe_with_rpe(
    df_gps: pd.DataFrame,
    session_date: Any,
    df_rpe: Optional[pd.DataFrame] = None,
    mapping_gps: Optional[Dict[Any, str]] = None,
) -> pd.DataFrame:
    """
    Enrichit le DataFrame GPS Catapult (df_indiv) avec les valeurs RPE individuelles :
    - Clé de jointure : Nom résolu + Date normalisés (sécurisée contre les inversions et surnoms).
    - Attribution des boîtiers GPS : Si un identifiant de capteur subsiste dans le tableau,
      la table de mapping de séance le résout impérativement vers le joueur officiel.
    - Attribution des postes : Strictement conservée depuis le tableau GPS.
    - Si un joueur n'a pas répondu : RPE = np.nan, RPE_Display = 'NR'.
    - Charge Foster calculée : Session_RPE = RPE × Durée (min).
    """
    if df_gps is None or not isinstance(df_gps, pd.DataFrame) or df_gps.empty:
        return df_gps

    df_out = df_gps.copy()
    target_dt = parse_and_normalize_date(session_date)

    c_gps_player = find_column_by_aliases(df_out, ["Nom ; Prénom", "Player Name", "Athlete Name", "Joueur", "Nom"])

    # 1. Résolution stricte des numéros de capteurs GPS vers les noms de joueurs de la séance
    if mapping_gps and c_gps_player and c_gps_player in df_out.columns:
        map_lookup = {}
        for k, v in mapping_gps.items():
            if v and str(v).strip() and str(v).strip() not in ["— Sélectionner un joueur —", "-"]:
                v_clean = str(v).strip()
                k_str = str(k).strip()
                map_lookup[k_str] = v_clean
                try:
                    f_k = float(k_str)
                    if f_k.is_integer():
                        map_lookup[str(int(f_k))] = v_clean
                        map_lookup[int(f_k)] = v_clean
                except Exception:
                    pass
        df_out[c_gps_player] = df_out[c_gps_player].apply(
            lambda val: map_lookup.get(val, map_lookup.get(str(val).strip(), val))
        )

    if df_rpe is None or df_rpe.empty:
        df_rpe = fetch_raw_rpe_data()

    # Si aucune donnée RPE n'est disponible
    if df_rpe is None or df_rpe.empty or target_dt is None:
        df_out["RPE"] = np.nan
        df_out["Session_RPE"] = np.nan
        df_out["RPE_Display"] = "NR"
        df_out["RPE_Duration"] = np.nan
        return df_out

    cols_map = identify_rpe_columns(df_rpe)
    col_date = cols_map["col_date"]
    col_nom = cols_map["col_nom"]
    col_prenom = cols_map["col_prenom"]
    col_player = cols_map["col_joueur_unique"]
    col_rpe = cols_map["col_rpe"]
    col_dur = cols_map["col_duration"]

    # Construction de la liste des réponses RPE pour la date cible
    # [(raw_player_name, (rpe_val, session_rpe, duration))]
    rpe_entries_for_date: List[Tuple[str, Tuple[float, float, float]]] = []

    for _, r in df_rpe.iterrows():
        # 1. Vérification de la date
        r_date = parse_and_normalize_date(r[col_date]) if col_date and col_date in r else None
        if r_date is None or r_date != target_dt:
            continue

        # 2. Construction du nom complet du joueur
        full_name_str = ""
        if col_player and col_player in r and pd.notna(r[col_player]):
            full_name_str = str(r[col_player]).strip()
        elif col_nom and col_prenom and col_nom in r and col_prenom in r and col_nom != col_prenom:
            val_n = str(r[col_nom]).strip() if pd.notna(r[col_nom]) else ""
            val_p = str(r[col_prenom]).strip() if pd.notna(r[col_prenom]) else ""
            full_name_str = f"{val_n} {val_p}".strip()
        elif col_nom and col_nom in r and pd.notna(r[col_nom]):
            full_name_str = str(r[col_nom]).strip()

        if not full_name_str:
            continue

        # 3. Extraction de la valeur RPE et de la durée
        rpe_val = parse_rpe_value(r[col_rpe]) if col_rpe and col_rpe in r else None
        if rpe_val is None:
            continue

        dur_val = parse_duration_minutes(r[col_dur]) if col_dur and col_dur in r else 80.0
        session_rpe = round(rpe_val * dur_val, 1)

        rpe_entries_for_date.append((full_name_str, (rpe_val, session_rpe, dur_val)))

    # 4. Jointure sécurisée avec le tableau GPS
    rpe_col_vals: List[Optional[float]] = []
    srpe_col_vals: List[Optional[float]] = []
    disp_col_vals: List[str] = []
    dur_col_vals: List[Optional[float]] = []

    for _, row in df_out.iterrows():
        p_name = str(row[c_gps_player]).strip() if c_gps_player and c_gps_player in row else ""
        matched_entry: Optional[Tuple[float, float, float]] = None

        # Recherche de correspondance par algorithme tolérant nom/prénom/alias
        for r_name_raw, entry in reversed(rpe_entries_for_date):
            if are_player_names_matching(p_name, r_name_raw):
                matched_entry = entry
                break

        if matched_entry is not None:
            r_val, s_rpe, dur = matched_entry
            rpe_col_vals.append(r_val)
            srpe_col_vals.append(s_rpe)
            disp_col_vals.append(f"{r_val:.1f}" if abs(r_val - round(r_val)) > 0.05 else f"{int(round(r_val))}")
            dur_col_vals.append(dur)
        else:
            rpe_col_vals.append(np.nan)
            srpe_col_vals.append(np.nan)
            disp_col_vals.append("NR")
            dur_col_vals.append(np.nan)

    df_out["RPE"] = rpe_col_vals
    df_out["Session_RPE"] = srpe_col_vals
    df_out["RPE_Display"] = disp_col_vals
    df_out["RPE_Duration"] = dur_col_vals

    return df_out


# ==============================================================================
# 5. CALCUL DES MOYENNES DE GROUPES & ÉQUIPE (AVEC EXCLUSION STRICTE DES 'NR')
# ==============================================================================

def calculate_rpe_group_metrics(df_enriched: pd.DataFrame) -> Dict[str, Any]:
    """
    Calcule les moyennes RPE et de charge Session-RPE (Foster) par ligne de poste et pour l'équipe :
    - 5 de devant (1ère & 2ème lignes)
    - 3ème Ligne
    - Trois-quarts
    - Équipe A (Effectif global)
    Règle cruciale : les valeurs 'NR' (NaN) sont strictement ignorées (skipna=True)
    afin de ne pas biaiser mathématiquement les moyennes.
    """
    default_stats: Dict[str, Any] = {
        "equipe": None,
        "5 de devant": None,
        "3ème Ligne": None,
        "Trois-quarts": None,
        "foster_equipe": None,
        "foster_5_devant": None,
        "foster_3eme_ligne": None,
        "foster_trois_quarts": None,
        "rpe_moyen_equipe": None,
        "rpe_moyen_5_devant": None,
        "rpe_moyen_3eme_ligne": None,
        "rpe_moyen_trois_quarts": None,
        "nb_repondants_tot": 0,
        "nb_joueurs_tot": len(df_enriched) if df_enriched is not None else 0,
        "taux_reponse_pct": 0.0,
    }

    if df_enriched is None or not isinstance(df_enriched, pd.DataFrame) or df_enriched.empty:
        return default_stats

    if "RPE" not in df_enriched.columns:
        return default_stats

    # Conversion propre en valeurs numériques avec exclusion des NaN
    rpe_num = pd.to_numeric(df_enriched["RPE"], errors="coerce")
    srpe_num = pd.to_numeric(df_enriched.get("Session_RPE", pd.Series(dtype=float)), errors="coerce")

    # Équipe globale
    valid_mask = rpe_num.notna()
    nb_rep = int(valid_mask.sum())
    nb_tot = len(df_enriched)

    default_stats["nb_repondants_tot"] = nb_rep
    default_stats["nb_joueurs_tot"] = nb_tot
    default_stats["taux_reponse_pct"] = round((nb_rep / nb_tot * 100.0), 1) if nb_tot > 0 else 0.0

    if nb_rep > 0:
        default_stats["equipe"] = round(float(rpe_num[valid_mask].mean()), 1)
        default_stats["rpe_moyen_equipe"] = default_stats["equipe"]
        if srpe_num.notna().any():
            default_stats["foster_equipe"] = round(float(srpe_num[valid_mask].mean()), 1)

    # Lignes de postes officielles RCA (dérivées exclusivement du tableau GPS)
    col_grp = "Groupe" if "Groupe" in df_enriched.columns else None

    group_mapping = {
        "5 de devant": ["5 de devant", "devant", "1ere ligne", "2eme ligne"],
        "3ème Ligne": ["3ème ligne", "3eme ligne", "troisieme ligne"],
        "Trois-quarts": ["trois-quarts", "trois quarts", "arrieres", "arrières", "3/4"],
    }

    for grp_std, aliases in group_mapping.items():
        if col_grp:
            mask_grp = df_enriched[col_grp].astype(str).str.strip().str.lower().apply(
                lambda s: any(a in s for a in aliases)
            )
            df_g = df_enriched[mask_grp]
        else:
            df_g = pd.DataFrame()

        if not df_g.empty:
            g_rpe = pd.to_numeric(df_g["RPE"], errors="coerce").dropna()
            g_srpe = pd.to_numeric(df_g.get("Session_RPE", pd.Series(dtype=float)), errors="coerce").dropna()

            if not g_rpe.empty:
                val_grp_rpe = round(float(g_rpe.mean()), 1)
                default_stats[grp_std] = val_grp_rpe
                if grp_std == "5 de devant":
                    default_stats["rpe_moyen_5_devant"] = val_grp_rpe
                elif grp_std == "3ème Ligne":
                    default_stats["rpe_moyen_3eme_ligne"] = val_grp_rpe
                elif grp_std == "Trois-quarts":
                    default_stats["rpe_moyen_trois_quarts"] = val_grp_rpe
            if not g_srpe.empty:
                foster_key = "foster_5_devant" if grp_std == "5 de devant" else ("foster_3eme_ligne" if grp_std == "3ème Ligne" else "foster_trois_quarts")
                default_stats[foster_key] = round(float(g_srpe.mean()), 1)

    return default_stats


# ==============================================================================
# 6. CALCUL RPE POUR LA SYNTHÈSE HEBDOMADAIRE (CUMUL & MOYENNES MULTI-SÉANCES)
# ==============================================================================

def calculate_weekly_rpe_from_sessions(synthese_dict: Dict[str, Any]) -> Dict[str, Any]:
    """
    Agrège les métriques RPE sur l'ensemble des séances enregistrées de la semaine :
    - Moyenne RPE hebdomadaire (Équipe, 5 de devant, 3ème Ligne, Trois-quarts)
    - Charge Foster cumulée (Session-RPE cumulé par joueur puis moyenne par groupe)
    """
    res: Dict[str, Any] = {
        "rpe_moyen_equipe": None,
        "foster_cumul_equipe": None,
        "rpe_5_devant": None,
        "rpe_3eme_ligne": None,
        "rpe_trois_quarts": None,
        "foster_5_devant": None,
        "foster_3eme_ligne": None,
        "foster_trois_quarts": None,
    }

    if not synthese_dict:
        return res

    all_players: List[Dict[str, Any]] = []
    for s_info in synthese_dict.values():
        if isinstance(s_info, dict):
            if "players_data" in s_info:
                p_data = s_info.get("players_data", [])
                if isinstance(p_data, list):
                    all_players.extend(p_data)
            elif "df_indiv" in s_info and isinstance(s_info["df_indiv"], pd.DataFrame):
                all_players.extend(s_info["df_indiv"].to_dict(orient="records"))

    if not all_players or not any("RPE" in p for p in all_players):
        # Fallback depuis les rpe_stats stockées au niveau séance
        team_rpes = []
        g_rpes = {"5 de devant": [], "3ème Ligne": [], "Trois-quarts": []}
        for s_info in synthese_dict.values():
            if isinstance(s_info, dict) and "rpe_stats" in s_info and isinstance(s_info["rpe_stats"], dict):
                r_st = s_info["rpe_stats"]
                if r_st.get("equipe") is not None:
                    team_rpes.append(r_st["equipe"])
                elif r_st.get("rpe_moyen_equipe") is not None:
                    team_rpes.append(r_st["rpe_moyen_equipe"])
                for g_key in g_rpes:
                    if r_st.get(g_key) is not None:
                        g_rpes[g_key].append(r_st[g_key])
        if team_rpes:
            res["rpe_moyen_equipe"] = round(float(np.mean(team_rpes)), 1)
            for g_key, vals in g_rpes.items():
                if vals:
                    k_out = "rpe_5_devant" if g_key == "5 de devant" else ("rpe_3eme_ligne" if g_key == "3ème Ligne" else "rpe_trois_quarts")
                    res[k_out] = round(float(np.mean(vals)), 1)
            return res
        return res

    df_all = pd.DataFrame(all_players)
    if df_all.empty or "RPE" not in df_all.columns:
        return res

    stats = calculate_rpe_group_metrics(df_all)

    res["rpe_moyen_equipe"] = stats.get("equipe")
    res["rpe_5_devant"] = stats.get("5 de devant")
    res["rpe_3eme_ligne"] = stats.get("3ème Ligne")
    res["rpe_trois_quarts"] = stats.get("Trois-quarts")

    # Calcul du cumul Foster hebdomadaire par joueur
    c_p = find_column_by_aliases(df_all, ["Nom ; Prénom", "Player Name", "Nom"])
    if c_p and "Session_RPE" in df_all.columns:
        df_all["Session_RPE_num"] = pd.to_numeric(df_all["Session_RPE"], errors="coerce")
        df_foster_p = df_all.groupby(c_p).agg({
            "Session_RPE_num": "sum",
            "Groupe": "first" if "Groupe" in df_all.columns else lambda x: "Tous",
        }).reset_index()

        valid_f = df_foster_p[df_foster_p["Session_RPE_num"] > 0]
        if not valid_f.empty:
            res["foster_cumul_equipe"] = round(float(valid_f["Session_RPE_num"].mean()), 1)

            for grp_name, foster_k in [
                ("5 de devant", "foster_5_devant"),
                ("3ème Ligne", "foster_3eme_ligne"),
                ("Trois-quarts", "foster_trois_quarts"),
            ]:
                if "Groupe" in valid_f.columns:
                    sub = valid_f[valid_f["Groupe"].str.lower().str.contains(grp_name[:4].lower(), na=False)]
                    if not sub.empty:
                        res[foster_k] = round(float(sub["Session_RPE_num"].mean()), 1)

    return res
