"""
Application Streamlit - Traitement des Données GPS Catapult One & Dashboard de Charge (STAPS)
Club : Rugby Club Amiénois (RCA)
Auteur : Expert Python / Streamlit Senior en Data Visualisation Sportive
Description :
    Application web complète combinant :
    1. Un Dashboard dynamique (style Canva) de comparaison Objectif vs Réalisé par groupe de joueurs
       (5 de devant, 3ème Ligne, Trois-quarts) réalisé avec Plotly Graph Objects.
    2. L'importation et le parsing robuste des exports bruts CSV Catapult One / PlayerTek / OpenField.
    3. La conversion précise des durées (Timedelta HH:MM:SS, secondes, millisecondes) en minutes réelles.
    4. L'injection directe dans le classeur modèle officiel (Modele_Vierge_STAPS.xlsx) pour préserver
       toutes les formules natives de la préparatrice physique, graphiques et onglets.
"""

from __future__ import annotations

import io
import os
import re
from typing import Any, Dict, List, Optional, Tuple, Union

import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ==============================================================================
# 0. CONSTANTES, MAQUETTES ET DONNÉES DE SIMULATION (MOCK DATA)
# ==============================================================================

TEMPLATE_FILENAME = "Modele_Vierge_STAPS.xlsx"
OUTPUT_DOWNLOAD_FILENAME = "Synthese_Complete_Seance.xlsx"

# Palette de couleurs officielles Rugby Club Amiénois / Style Canva
COLOR_OBJECTIF = "#2E5A88"  # Bleu roi / marine
COLOR_REALISE = "#A30000"   # Rouge / Bordeaux RCA
COLOR_TEXT_MAIN = "#1F4E79"
COLOR_MUTED = "#555555"

FINAL_COLUMNS: List[str] = [
    'Nom ; Prénom',
    'Temps de jeu ',
    'Distance (km)',
    ' Distance haute intensité (m)',
    'Top Speed (km/h)',
    'Distance Per Min (m/min)',
    'Distance in Speed Zone 1  (km)',
    'Distance in Speed Zone 2  (km)',
    'Distance in Speed Zone 3  (km)',
    'Distance in Speed Zone 4  (km)',
    'Distance in Speed Zone 5  (km)',
    'Max Deceleration (m/s/s)',
    'Max Acceleration (m/s/s)',
    'Accelerations Zone Count: 3 - 4 m/s/s',
    'Accelerations Zone Count: > 4 m/s/s',
    'Deceleration Zone Count: 3 - 4 m/s/s',
    'Deceleration Zone Count: > 4 m/s/s',
    'Accel/Dec',
    '%Marche ',
    '%course',
    '%HSR ',
    '%VHSR',
    'HSR m/min',
]

PCT_COLUMNS_CLEAN: set[str] = {"%marche", "%course", "%hsr", "%vhsr"}

# Données simulées (Mock Data) pour la visualisation immédiate de l'esthétique Canva
MOCK_OBJECTIFS: Dict[str, Dict[str, Tuple[float, float]]] = {
    "5 de devant": {
        "Objectif Distance (m)": (4500, 2475),  # 55% d'accomplissement (explicitement cité)
        "Obj. HSR": (250, 160),                  # 64%
        "Obj. VHSR": (50, 32),                   # 64%
        "Obj. ACC/DEC": (35, 28),                # 80%
    },
    "3ème Ligne": {
        "Objectif Distance (m)": (5500, 4125),  # 75%
        "Obj. HSR": (450, 380),                  # 84%
        "Obj. VHSR": (100, 82),                  # 82%
        "Obj. ACC/DEC": (45, 41),                # 91%
    },
    "Trois-quarts": {
        "Objectif Distance (m)": (6200, 5580),  # 90%
        "Obj. HSR": (700, 620),                  # 89%
        "Obj. VHSR": (250, 215),                 # 86%
        "Obj. ACC/DEC": (50, 47),                # 94%
    },
}

# Données simulées (Mock Data) pour le Podium des Tops de la Séance
MOCK_PODIUM: Dict[str, Dict[str, Any]] = {
    "Vitesse Max (km/h)": {
        "icon": "⚡",
        "players": [
            {"rank": 1, "name": "Antoine D.", "stat": "32.4 km/h", "seed": "Antoine+D"},
            {"rank": 2, "name": "Thomas L.", "stat": "31.8 km/h", "seed": "Thomas+L"},
            {"rank": 3, "name": "Lucas M.", "stat": "31.1 km/h", "seed": "Lucas+M"},
        ],
    },
    "% de marche": {
        "icon": "🚶",
        "players": [
            {"rank": 1, "name": "Paul B.", "stat": "58.4 %", "seed": "Paul+B"},
            {"rank": 2, "name": "Julien R.", "stat": "56.2 %", "seed": "Julien+R"},
            {"rank": 3, "name": "Maxime G.", "stat": "54.8 %", "seed": "Maxime+G"},
        ],
    },
    "Distance Haute Intensité (m)": {
        "icon": "🔥",
        "players": [
            {"rank": 1, "name": "Romain C.", "stat": "742 m", "seed": "Romain+C"},
            {"rank": 2, "name": "Clément V.", "stat": "685 m", "seed": "Clement+V"},
            {"rank": 3, "name": "Nicolas F.", "stat": "610 m", "seed": "Nicolas+F"},
        ],
    },
}


def get_template_path() -> Optional[str]:
    """
    Recherche le fichier modèle Modele_Vierge_STAPS.xlsx dans le répertoire du script
    ou dans le répertoire de travail courant.
    """
    possible_paths = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), TEMPLATE_FILENAME),
        os.path.join(os.getcwd(), TEMPLATE_FILENAME),
        TEMPLATE_FILENAME,
    ]
    for p in possible_paths:
        if os.path.exists(p):
            return p
    return None


# ==============================================================================
# 1. DATA VISUALISATION PLOTLY (STYLE CANVA ÉPURÉ)
# ==============================================================================

def creer_graphique_jauge(titre: str, val_objectif: float, val_realise: float) -> go.Figure:
    """
    Génère un graphique à barres (Bar Chart) stylisé inspiré du design Canva :
    - Deux barres côte à côte : 'Objectif' (Bleu roi/marine) et 'Réalisé' (Rouge/Bordeaux).
    - Valeur brute affichée au-dessus de chaque barre.
    - Pourcentage d'accomplissement (ex: '55%') affiché en grand texte blanc et gras à l'intérieur
      de la barre 'Réalisé'.
    - Design minimaliste : fond transparent, sans grille, sans axe Y, typographie moderne sans-serif.
    """
    val_obj = float(val_objectif)
    val_real = float(val_realise)
    pct = (val_real / val_obj * 100) if val_obj > 0 else 0.0

    # Formatage lisible des valeurs brutes au-dessus des barres (séparateur de milliers)
    fmt_obj = f"{int(round(val_obj)):,d}".replace(",", " ") if abs(val_obj - round(val_obj)) < 0.01 else f"{val_obj:.1f}"
    fmt_real = f"{int(round(val_real)):,d}".replace(",", " ") if abs(val_real - round(val_real)) < 0.01 else f"{val_real:.1f}"

    max_val = max(val_obj, val_real, 1.0)
    y_max = max_val * 1.30  # Marge supérieure pour éviter que le texte ne soit tronqué

    fig = go.Figure()

    # Barres Objectif et Réalisé
    fig.add_trace(go.Bar(
        x=["Objectif", "Réalisé"],
        y=[val_obj, val_real],
        marker=dict(
            color=[COLOR_OBJECTIF, COLOR_REALISE],
            line=dict(width=0),
            cornerradius=8,  # Bords arrondis pour le rendu moderne Canva
        ),
        text=[f"<b>{fmt_obj}</b>", f"<b>{fmt_real}</b>"],
        textposition="outside",
        textfont=dict(family="Segoe UI, Montserrat, Inter, sans-serif", size=13, color="#1A2B4C"),
        width=[0.42, 0.42],
        hovertemplate="<b>%{x}</b> : %{text}<extra></extra>",
        showlegend=False,
    ))

    # Pourcentage d'accomplissement centré dans la barre Réalisé
    if val_real > 0.14 * max_val:
        y_pct = val_real * 0.50
        font_color = "#FFFFFF"
        font_size = 19
    else:
        # Si la barre est trop basse, le pourcentage se place juste au-dessus
        y_pct = val_real + (max_val * 0.08)
        font_color = COLOR_REALISE
        font_size = 14

    fig.add_annotation(
        x="Réalisé",
        y=y_pct,
        text=f"<b>{pct:.0f}%</b>",
        showarrow=False,
        font=dict(family="Segoe UI, Montserrat, Inter, sans-serif", size=font_size, color=font_color),
        align="center",
    )

    # Configuration épurée du layout
    fig.update_layout(
        title=dict(
            text=f"<b>{titre}</b>",
            font=dict(family="Segoe UI, Montserrat, Inter, sans-serif", size=14, color=COLOR_TEXT_MAIN),
            x=0.5,
            xanchor="center",
            y=0.96,
        ),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=10, r=10, t=46, b=26),
        height=275,
        xaxis=dict(
            showgrid=False,
            showline=False,
            zeroline=False,
            tickfont=dict(family="Segoe UI, Montserrat, Inter, sans-serif", size=12, color="#4A5568", weight="bold"),
        ),
        yaxis=dict(
            showgrid=False,
            showline=False,
            showticklabels=False,
            zeroline=False,
            range=[0, y_max],
        ),
    )
    return fig


def render_dashboard_objectifs(group_name: str, metrics: Dict[str, Tuple[float, float]]) -> None:
    """
    Affiche 4 graphiques stylisés sur une seule ligne (st.columns(4)) pour le groupe sélectionné.
    """
    col1, col2, col3, col4 = st.columns(4)
    metric_keys = list(metrics.keys())

    cols = [col1, col2, col3, col4]
    for idx, key in enumerate(metric_keys[:4]):
        val_obj, val_real = metrics[key]
        fig = creer_graphique_jauge(key, val_obj, val_real)
        with cols[idx]:
            st.plotly_chart(
                fig,
                use_container_width=True,
                config={"displayModeBar": False},  # Masque la barre d'outils Plotly pour un effet Canva
            )


def render_podium_section(podium_data: Dict[str, Dict[str, Any]]) -> None:
    """
    Affiche la section Podium (Style Canva) avec les Tops de la séance :
    - Séparateur visuel st.divider() et sous-titre '🏆 PODIUM : Les Tops de la Séance'.
    - st.columns(3) pour diviser la largeur en 3 blocs verticaux égaux :
      * 'Vitesse Max (km/h)'
      * '% de marche'
      * 'Distance Haute Intensité (m)'
    - Pour chaque catégorie :
      * Le 1er (vainqueur) est mis en valeur au sommet avec un avatar plus grand (78px),
        un badge doré 🥇 et un fond dégradé doux.
      * Le 2ème (🥈) et le 3ème (🥉) sont disposés côte à côte juste en dessous avec
        un avatar compact (52px).
      * Avatars circulaires dynamiques générés via l'API Dicebear (fond bleu RCA #2E5A88).
      * Statistiques en gras et mise en valeur en bordeaux RCA (#A30000).
    """
    st.divider()
    st.markdown('<div class="section-header">🏆 PODIUM : Les Tops de la Séance</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-desc">'
        'Mise à l\'honneur des meilleures performances individuelles de la séance sur les métriques athlétiques clés.'
        '</div>',
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns(3)
    cols = [col1, col2, col3]
    categories = list(podium_data.keys())

    for idx, cat_name in enumerate(categories):
        cat_info = podium_data[cat_name]
        players = cat_info["players"]
        p1, p2, p3 = players[0], players[1], players[2]

        card_html = f"""
        <div class="podium-card-container">
            <div class="podium-category-header">
                <span class="podium-cat-icon">{cat_info['icon']}</span>
                <span class="podium-cat-title">{cat_name}</span>
            </div>

            <!-- 1er / Vainqueur de la catégorie (Mise en avant au sommet) -->
            <div class="podium-winner-box">
                <div class="podium-rank-badge rank-1">🥇 1er</div>
                <div class="avatar-wrapper">
                    <img src="https://api.dicebear.com/7.x/initials/svg?seed={p1['seed']}&backgroundColor=2E5A88" 
                         alt="{p1['name']}" class="podium-avatar img-winner" />
                </div>
                <div class="podium-player-name name-winner">{p1['name']}</div>
                <div class="podium-player-stat stat-winner">{p1['stat']}</div>
            </div>

            <!-- 2ème et 3ème côte à côte juste en dessous -->
            <div class="podium-runners-row">
                <!-- 2ème -->
                <div class="podium-runner-box">
                    <div class="podium-rank-badge rank-2">🥈 2ème</div>
                    <div class="avatar-wrapper">
                        <img src="https://api.dicebear.com/7.x/initials/svg?seed={p2['seed']}&backgroundColor=2E5A88" 
                             alt="{p2['name']}" class="podium-avatar img-runner" />
                    </div>
                    <div class="podium-player-name">{p2['name']}</div>
                    <div class="podium-player-stat">{p2['stat']}</div>
                </div>

                <!-- 3ème -->
                <div class="podium-runner-box">
                    <div class="podium-rank-badge rank-3">🥉 3ème</div>
                    <div class="avatar-wrapper">
                        <img src="https://api.dicebear.com/7.x/initials/svg?seed={p3['seed']}&backgroundColor=2E5A88" 
                             alt="{p3['name']}" class="podium-avatar img-runner" />
                    </div>
                    <div class="podium-player-name">{p3['name']}</div>
                    <div class="podium-player-stat">{p3['stat']}</div>
                </div>
            </div>
        </div>
        """
        with cols[idx]:
            st.markdown(card_html, unsafe_allow_html=True)


# ==============================================================================
# 2. PARSING ROBUSTE & CONVERSION NUMÉRIQUE / TEMPORELLE
# ==============================================================================

def parse_duration_to_minutes(val: Any) -> float:
    """
    Convertit une valeur de durée individuelle en minutes réelles (float).
    Prend en charge :
    - Format chaîne texte 'HH:MM:SS' ou 'H:MM:SS' via Timedelta
    - Format chaîne texte 'MM:SS'
    - Nombres en millisecondes (> 100 000)
    - Nombres en secondes (> 240, ex: 8508 s -> 141.8 min)
    - Nombres déjà en minutes (ex: 60.5, 80.0, 120.0)
    """
    if pd.isna(val) or val is None or val == "":
        return 0.0

    if isinstance(val, (int, float)):
        num = float(val)
        if num > 100000:
            return num / 60000.0
        elif num > 240:
            return num / 60.0
        return num

    val_str = str(val).strip().replace(",", ".")

    if ":" in val_str:
        try:
            td = pd.to_timedelta(val_str)
            return float(td.total_seconds() / 60.0)
        except Exception:
            parts = val_str.split(":")
            try:
                if len(parts) == 3:
                    return float(parts[0]) * 60.0 + float(parts[1]) + float(parts[2]) / 60.0
                elif len(parts) == 2:
                    return float(parts[0]) + float(parts[1]) / 60.0
            except (ValueError, TypeError):
                return 0.0

    try:
        num = float(val_str)
        if num > 100000:
            return num / 60000.0
        elif num > 240:
            return num / 60.0
        return num
    except (ValueError, TypeError):
        return 0.0


def parse_duration_series(series: pd.Series, col_name: str = "") -> pd.Series:
    """
    Convertit une série complète de durées Catapult en minutes réelles (float).
    """
    if series.empty:
        return pd.Series(dtype=float)

    col_lower = str(col_name).lower()
    raw_str = series.astype(str).str.strip().str.replace(",", ".")

    has_colon = raw_str.str.contains(":", na=False).any()
    if has_colon:
        return series.apply(parse_duration_to_minutes).astype(float)

    num_series = pd.to_numeric(raw_str, errors="coerce").fillna(0.0).astype(float)

    if any(u in col_lower for u in ["(ms)", " ms", "millisec"]):
        return (num_series / 60000.0).astype(float)
    if any(u in col_lower for u in ["(s)", "(sec)", "(secs)", " seconds", " sec"]):
        return (num_series / 60.0).astype(float)
    if any(u in col_lower for u in ["(min)", "(mins)", " minutes"]):
        return num_series.astype(float)

    max_val = num_series.max() if len(num_series) > 0 else 0.0
    positive_vals = num_series[num_series > 0]
    median_val = positive_vals.median() if len(positive_vals) > 0 else 0.0

    if max_val > 100000:
        return (num_series / 60000.0).astype(float)
    elif max_val > 240 or median_val > 180:
        return (num_series / 60.0).astype(float)
    else:
        return num_series.astype(float)


def clean_numeric_value(val: Any) -> float:
    """
    Convertit une valeur quelconque contenant éventuellement une virgule en float.
    """
    if pd.isna(val) or val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace(" ", "").replace("\xa0", "").replace(",", ".")
    try:
        return float(val_str)
    except (ValueError, TypeError):
        return 0.0


def find_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """
    Recherche une colonne correspondant à une liste de candidats.
    """
    col_map = {re.sub(r"\s+", " ", str(c).strip().lower()): c for c in df.columns}
    
    for candidate in candidates:
        norm_candidate = re.sub(r"\s+", " ", candidate.strip().lower())
        if norm_candidate in col_map:
            return col_map[norm_candidate]
            
    for candidate in candidates:
        norm_candidate = re.sub(r"\s+", " ", candidate.strip().lower())
        for norm_col, orig_col in col_map.items():
            if norm_candidate in norm_col:
                return orig_col
                
    return None


def find_catapult_header_row(file_bytes: bytes, encoding: str = "utf-8", max_scan: int = 10) -> int:
    """
    Scanne les 10 premières lignes du fichier brut pour repérer dynamiquement
    celle qui correspond au véritable en-tête du tableau (contenant 'Player Name' ou 'Date').
    """
    try:
        text_preview = file_bytes[:16384].decode(encoding, errors="replace")
        lines = text_preview.splitlines()[:max_scan]
    except Exception:
        return 0

    for idx, raw_line in enumerate(lines):
        line = raw_line.strip().lower()
        if not line:
            continue
        if any(term in line for term in ["player name", "athlete name", "nom ; prénom", "nom;prénom"]):
            return idx

    for idx, raw_line in enumerate(lines):
        line = raw_line.strip().lower()
        if not line:
            continue
        if "date" in line:
            if any(sep in line for sep in [";", ",", "\t"]) or re.search(r'\bdate\b', line):
                return idx

    for idx, raw_line in enumerate(lines):
        line = raw_line.strip().lower()
        if not line:
            continue
        if any(term in line for term in ["player", "athlete", "joueur", "split", "period", "distance"]):
            if any(sep in line for sep in [";", ",", "\t"]):
                return idx

    return 0


def load_catapult_csv(file_input: Union[bytes, io.BytesIO, Any]) -> pd.DataFrame:
    """
    Lit un fichier CSV Catapult One en détectant dynamiquement le vrai header
    parmi les 10 premières lignes.
    """
    if hasattr(file_input, "getvalue"):
        file_bytes = file_input.getvalue()
    elif hasattr(file_input, "read"):
        if hasattr(file_input, "seek"):
            file_input.seek(0)
        file_bytes = file_input.read()
    elif isinstance(file_input, bytes):
        file_bytes = file_input
    else:
        file_bytes = bytes(file_input)

    encodings_to_try: List[str] = ["utf-8-sig", "utf-8", "latin1", "cp1252", "iso-8859-1"]
    separators_to_try: List[str] = [";", ",", "\t"]

    last_error: Optional[Exception] = None

    for enc in encodings_to_try:
        header_idx = find_catapult_header_row(file_bytes, encoding=enc, max_scan=10)

        for sep in separators_to_try:
            try:
                df = pd.read_csv(
                    io.BytesIO(file_bytes),
                    sep=sep,
                    encoding=enc,
                    header=header_idx,
                    low_memory=False,
                )
                if len(df.columns) >= 3:
                    df.columns = [str(c).strip() for c in df.columns]
                    lower_cols = [c.lower() for c in df.columns]
                    if any("player" in c or "athlete" in c or "nom" in c or "date" in c for c in lower_cols):
                        return df
            except Exception as exc:
                last_error = exc
                continue

    for enc in encodings_to_try:
        header_idx = find_catapult_header_row(file_bytes, encoding=enc, max_scan=10)
        for sep in separators_to_try:
            try:
                df = pd.read_csv(
                    io.BytesIO(file_bytes),
                    sep=sep,
                    encoding=enc,
                    header=header_idx,
                    low_memory=False,
                )
                if len(df.columns) >= 3:
                    df.columns = [str(c).strip() for c in df.columns]
                    return df
            except Exception as exc:
                last_error = exc
                continue

    if last_error is not None:
        raise last_error
    raise ValueError("Impossible de décoder le fichier CSV avec les encodages et délimiteurs standards.")


# ==============================================================================
# 3. CALCULS STAPS & INGÉNIERIE DES CARACTÉRISTIQUES (GROUPBY JOUEUR)
# ==============================================================================

def process_catapult_data(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Traite le DataFrame brut Catapult One :
    - Regroupe les données par joueur (session globale).
    - Convertit rigoureusement le Temps de jeu en minutes réelles (float).
    - Calcule et agrège toutes les métriques requises.
    - Applique les formules de feature engineering STAPS (Distance/min, HSR/min).
    - Construit et ordonne rigoureusement df_indiv selon la maquette.
    """
    # 1. Identification de la colonne Joueur
    player_col = find_column(df_raw, [
        "Nom ; Prénom", "Player Name", "Athlete Name", "Player", "Athlete", "Joueur", "Nom", "Name"
    ])
    if not player_col:
        raise ValueError(
            "Colonne Joueur introuvable dans le CSV. Colonnes attendues : 'Player Name', 'Athlete Name', etc."
        )

    # 2. Détection des colonnes sources
    dur_col = find_column(df_raw, [
        "Duration", "Temps de jeu", "Duration (mins)", "Duration (s)", "Duration (sec)",
        "Durée", "Time", "Session Duration"
    ])
    dist_col = find_column(df_raw, ["Distance (km)", "Total Distance (km)", "Distance"])
    dist_m_col = find_column(df_raw, ["Distance (m)", "Total Distance (m)"])

    hsr_col = find_column(df_raw, [
        "Distance haute intensité (m)", "Distance haute intensite (m)", "High Speed Running (m)", 
        "HSR Distance (m)", "HSR (m)", "HSR Distance (>5m/s)"
    ])
    vhsr_col = find_column(df_raw, [
        "Sprint Distance (m)", "VHSR Distance (m)", "VHSR (m)", "Very High Speed Running (m)"
    ])

    top_speed_col = find_column(df_raw, [
        "Top Speed (km/h)", "Max Velocity (km/h)", "Top Speed", "Max Velocity", "Vmax"
    ])

    max_accel_col = find_column(df_raw, [
        "Max Acceleration (m/s/s)", "Max Acceleration (m/s2)", "Max Acceleration", "Max Accel"
    ])
    max_decel_col = find_column(df_raw, [
        "Max Deceleration (m/s/s)", "Max Deceleration (m/s2)", "Max Deceleration", "Max Decel"
    ])

    # Speed Zones
    z1_col = find_column(df_raw, ["Distance in Speed Zone 1", "Speed Zone 1", "Zone 1"])
    z2_col = find_column(df_raw, ["Distance in Speed Zone 2", "Speed Zone 2", "Zone 2"])
    z3_col = find_column(df_raw, ["Distance in Speed Zone 3", "Speed Zone 3", "Zone 3"])
    z4_col = find_column(df_raw, ["Distance in Speed Zone 4", "Speed Zone 4", "Zone 4"])
    z5_col = find_column(df_raw, ["Distance in Speed Zone 5", "Speed Zone 5", "Zone 5"])

    # Accel / Decel Zone Counts
    acc_34_col = find_column(df_raw, [
        "Accelerations Zone Count: 3 - 4 m/s/s", "Accelerations Zone Count: 3-4 m/s/s",
        "Accel 3-4", "Accelerations (3-4 m/s2)", "Accelerations 3 - 4", "Accélérations (>3m/s2)", "Accelerations (>3m/s2)"
    ])
    acc_gt4_col = find_column(df_raw, [
        "Accelerations Zone Count: > 4 m/s/s", "Accelerations Zone Count: >4 m/s/s",
        "Accel >4", "Accelerations (>4 m/s2)", "Accelerations > 4"
    ])
    dec_34_col = find_column(df_raw, [
        "Deceleration Zone Count: 3 - 4 m/s/s", "Deceleration Zone Count: 3-4 m/s/s",
        "Decel 3-4", "Decelerations (3-4 m/s2)", "Decelerations 3 - 4", "Décélérations (<-3m/s2)", "Decelerations (<-3m/s2)"
    ])
    dec_gt4_col = find_column(df_raw, [
        "Deceleration Zone Count: > 4 m/s/s", "Deceleration Zone Count: >4 m/s/s",
        "Decel >4", "Decelerations (>4 m/s2)", "Decelerations > 4"
    ])

    # 3. Création du DataFrame intermédiaire nettoyé
    df_clean = pd.DataFrame()
    df_clean["Nom ; Prénom"] = df_raw[player_col].astype(str).str.strip()

    # Temps de jeu (en minutes réelles float)
    if dur_col:
        df_clean["Duration_min"] = parse_duration_series(df_raw[dur_col], col_name=dur_col)
    else:
        df_clean["Duration_min"] = 0.0

    # Distance totale (en km)
    if dist_col:
        vals = df_raw[dist_col].apply(clean_numeric_value)
        if "(m)" in dist_col.lower() or vals.median() > 100.0:
            df_clean["Distance_km"] = vals / 1000.0
        else:
            df_clean["Distance_km"] = vals
    elif dist_m_col:
        df_clean["Distance_km"] = df_raw[dist_m_col].apply(clean_numeric_value) / 1000.0
    else:
        df_clean["Distance_km"] = 0.0

    # Top Speed (km/h)
    if top_speed_col:
        df_clean["Top_Speed"] = df_raw[top_speed_col].apply(clean_numeric_value)
    else:
        df_clean["Top_Speed"] = 0.0

    # Max Acceleration & Max Deceleration
    if max_accel_col:
        df_clean["Max_Accel"] = df_raw[max_accel_col].apply(clean_numeric_value)
    else:
        df_clean["Max_Accel"] = 0.0

    if max_decel_col:
        df_clean["Max_Decel"] = df_raw[max_decel_col].apply(lambda x: abs(clean_numeric_value(x)))
    else:
        df_clean["Max_Decel"] = 0.0

    # Zones de vitesse 1 à 5 (forcées en kilomètres)
    for z_idx, z_name, c_source in [
        (1, "Z1_km", z1_col),
        (2, "Z2_km", z2_col),
        (3, "Z3_km", z3_col),
        (4, "Z4_km", z4_col),
        (5, "Z5_km", z5_col),
    ]:
        if c_source:
            vals = df_raw[c_source].apply(clean_numeric_value)
            if "(m)" in c_source.lower() or vals.median() > 20.0:
                vals = vals / 1000.0
            df_clean[z_name] = vals
        else:
            df_clean[z_name] = 0.0

    # Distance haute intensité (m) = Somme HSR + VHSR (ou Sprint)
    if hsr_col and vhsr_col:
        df_clean["HSR_m"] = df_raw[hsr_col].apply(clean_numeric_value) + df_raw[vhsr_col].apply(clean_numeric_value)
    elif hsr_col:
        df_clean["HSR_m"] = df_raw[hsr_col].apply(clean_numeric_value)
    else:
        df_clean["HSR_m"] = (df_clean["Z3_km"] + df_clean["Z4_km"] + df_clean["Z5_km"]) * 1000.0

    # Comptages Accélérations & Décélérations
    df_clean["Acc_34"] = df_raw[acc_34_col].apply(clean_numeric_value) if acc_34_col else 0.0
    df_clean["Acc_gt4"] = df_raw[acc_gt4_col].apply(clean_numeric_value) if acc_gt4_col else 0.0
    df_clean["Dec_34"] = df_raw[dec_34_col].apply(clean_numeric_value) if dec_34_col else 0.0
    df_clean["Dec_gt4"] = df_raw[dec_gt4_col].apply(clean_numeric_value) if dec_gt4_col else 0.0

    # 4. Groupby Joueur sur la session globale
    grouped = df_clean.groupby("Nom ; Prénom", as_index=False).agg({
        "Duration_min": "sum",
        "Distance_km": "sum",
        "HSR_m": "sum",
        "Top_Speed": "max",
        "Z1_km": "sum",
        "Z2_km": "sum",
        "Z3_km": "sum",
        "Z4_km": "sum",
        "Z5_km": "sum",
        "Max_Decel": "max",
        "Max_Accel": "max",
        "Acc_34": "sum",
        "Acc_gt4": "sum",
        "Dec_34": "sum",
        "Dec_gt4": "sum",
    })

    # Double sécurité post-groupby sur Duration_min
    grouped["Duration_min"] = grouped["Duration_min"].astype(float)
    positive_durations = grouped["Duration_min"][grouped["Duration_min"] > 0]
    if len(positive_durations) > 0:
        med_dur = positive_durations.median()
        if med_dur > 100000:
            grouped["Duration_min"] = grouped["Duration_min"] / 60000.0
        elif med_dur > 240:
            grouped["Duration_min"] = grouped["Duration_min"] / 60.0

    # 5. Formules spécifiques & Feature Engineering STAPS
    temps_jeu = grouped["Duration_min"].astype(float)
    d_total_m = grouped["Distance_km"] * 1000.0
    dist_tot_km = grouped["Distance_km"]

    # Distance Per Min (m/min) : Distance totale (m) / Temps de jeu réel (min)
    dist_per_min = np.where(temps_jeu > 0, d_total_m / temps_jeu, 0.0)

    # HSR m/min = Distance haute intensité (m) / Temps de jeu réel (min)
    hsr_per_min = np.where(temps_jeu > 0, grouped["HSR_m"] / temps_jeu, 0.0)

    # Accel/Dec = (Accel 3-4) + (Accel >4) + (Decel 3-4) + (Decel >4)
    accel_decel_total = (
        grouped["Acc_34"] + grouped["Acc_gt4"] + grouped["Dec_34"] + grouped["Dec_gt4"]
    )

    # Pourcentages par zone de vitesse relatifs à la distance totale
    pct_marche = np.where(dist_tot_km > 0, grouped["Z1_km"] / dist_tot_km, 0.0)
    pct_course = np.where(dist_tot_km > 0, grouped["Z2_km"] / dist_tot_km, 0.0)
    pct_hsr = np.where(dist_tot_km > 0, grouped["Z3_km"] / dist_tot_km, 0.0)
    pct_vhsr = np.where(dist_tot_km > 0, (grouped["Z4_km"] + grouped["Z5_km"]) / dist_tot_km, 0.0)

    # 6. Assemblage du DataFrame final df_indiv
    df_indiv = pd.DataFrame()
    df_indiv['Nom ; Prénom'] = grouped['Nom ; Prénom']
    df_indiv['Temps de jeu '] = temps_jeu.astype(float)
    df_indiv['Distance (km)'] = grouped['Distance_km']
    df_indiv[' Distance haute intensité (m)'] = grouped['HSR_m']
    df_indiv['Top Speed (km/h)'] = grouped['Top_Speed']
    df_indiv['Distance Per Min (m/min)'] = dist_per_min
    df_indiv['Distance in Speed Zone 1  (km)'] = grouped['Z1_km']
    df_indiv['Distance in Speed Zone 2  (km)'] = grouped['Z2_km']
    df_indiv['Distance in Speed Zone 3  (km)'] = grouped['Z3_km']
    df_indiv['Distance in Speed Zone 4  (km)'] = grouped['Z4_km']
    df_indiv['Distance in Speed Zone 5  (km)'] = grouped['Z5_km']
    df_indiv['Max Deceleration (m/s/s)'] = grouped['Max_Decel']
    df_indiv['Max Acceleration (m/s/s)'] = grouped['Max_Accel']
    df_indiv['Accelerations Zone Count: 3 - 4 m/s/s'] = grouped['Acc_34']
    df_indiv['Accelerations Zone Count: > 4 m/s/s'] = grouped['Acc_gt4']
    df_indiv['Deceleration Zone Count: 3 - 4 m/s/s'] = grouped['Dec_34']
    df_indiv['Deceleration Zone Count: > 4 m/s/s'] = grouped['Dec_gt4']
    df_indiv['Accel/Dec'] = accel_decel_total
    df_indiv['%Marche '] = pct_marche
    df_indiv['%course'] = pct_course
    df_indiv['%HSR '] = pct_hsr
    df_indiv['%VHSR'] = pct_vhsr
    df_indiv['HSR m/min'] = hsr_per_min

    return df_indiv[FINAL_COLUMNS]


# ==============================================================================
# 4. INJECTION DANS LE TEMPLATE EXCEL VIA OPENPYXL (data_only=False)
# ==============================================================================

def populate_template_excel(df_indiv: pd.DataFrame, df_raw: Optional[pd.DataFrame] = None) -> io.BytesIO:
    """
    Charge le classeur modèle officiel (Modele_Vierge_STAPS.xlsx) avec openpyxl (data_only=False) :
    - Préserve l'intégralité des onglets, formules natives et graphiques de la préparatrice physique.
    - Injecte df_indiv dans 'données indiv. joueurs' à partir de la ligne 3.
    - Met à jour les formules de synthèse de ligne si présentes.
    - Injecte df_raw dans 'données brutes' (en-têtes en ligne 3, données en ligne 4+).
    - Sauvegarde et renvoie le classeur complet sous forme de flux io.BytesIO.
    """
    template_path = get_template_path()
    if not template_path:
        raise FileNotFoundError(
            f"Le fichier modèle '{TEMPLATE_FILENAME}' est introuvable. "
            "Veuillez vous assurer qu'il se trouve dans le même dossier que l'application."
        )

    # Chargement avec préservation absolue de toutes les formules natives
    wb = openpyxl.load_workbook(template_path, data_only=False)

    # --------------------------------------------------------------------------
    # A. Injection dans 'données indiv. joueurs'
    # --------------------------------------------------------------------------
    if 'données indiv. joueurs' in wb.sheetnames:
        ws_indiv = wb['données indiv. joueurs']
        num_players = len(df_indiv)

        # Écriture des joueurs à partir de la ligne 3
        for i, row in df_indiv.iterrows():
            r = 3 + i
            # Colonnes 1 à 22
            for col_idx, col_name in enumerate(FINAL_COLUMNS[:22], start=1):
                val = row[col_name]
                ws_indiv.cell(row=r, column=col_idx, value=val)

            # Colonne 23 (somme des 4 zones de % = 100%)
            ws_indiv.cell(row=r, column=23, value=f"=SUM(S{r}:V{r})")

            # Colonne 24 ('m/min' -> HSR m/min)
            ws_indiv.cell(row=r, column=24, value=row['HSR m/min'])

            # Colonnes d'extension du modèle de la préparatrice
            ws_indiv.cell(row=r, column=26, value=f"=I{r}*1000")
            ws_indiv.cell(row=r, column=27, value=f"=(SUM(J{r}:K{r}))*1000")

        # Nettoyage des lignes excédentaires si le modèle contenait plus de lignes
        last_injected_row = 3 + num_players - 1
        if ws_indiv.max_row > last_injected_row:
            for extra_r in range(last_injected_row + 1, ws_indiv.max_row + 1):
                for c in range(1, 35):
                    ws_indiv.cell(row=extra_r, column=c).value = None

    # --------------------------------------------------------------------------
    # B. Injection dans 'données brutes' (archivage brut de la session)
    # --------------------------------------------------------------------------
    if df_raw is not None and 'données brutes' in wb.sheetnames:
        ws_brutes = wb['données brutes']

        # En-têtes bruts en ligne 3
        for col_idx, col_name in enumerate(df_raw.columns, start=1):
            ws_brutes.cell(row=3, column=col_idx, value=col_name)

        # Lignes brutes à partir de la ligne 4
        for row_idx, row_vals in enumerate(df_raw.itertuples(index=False), start=4):
            for col_idx, cell_val in enumerate(row_vals, start=1):
                val_to_write = cell_val if not pd.isna(cell_val) else None
                ws_brutes.cell(row=row_idx, column=col_idx, value=val_to_write)

        # Nettoyage des éventuelles lignes antérieures dépassées
        last_brute_row = 4 + len(df_raw) - 1
        if ws_brutes.max_row > last_brute_row:
            for extra_r in range(last_brute_row + 1, ws_brutes.max_row + 1):
                for c in range(1, len(df_raw.columns) + 1):
                    ws_brutes.cell(row=extra_r, column=c).value = None

    # --------------------------------------------------------------------------
    # C. Sauvegarde en mémoire virtuelle
    # --------------------------------------------------------------------------
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def create_fallback_excel(df_indiv: pd.DataFrame) -> io.BytesIO:
    """
    Génération Excel de secours autonome si le fichier modèle Modele_Vierge_STAPS.xlsx est absent.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_indiv.to_excel(writer, sheet_name="données indiv. joueurs", index=False)
        ws = writer.sheets["données indiv. joueurs"]
        ws.row_dimensions[1].height = 30
        ws.freeze_panes = "B2"
        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        for col_idx in range(1, len(FINAL_COLUMNS) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            col_letter = get_column_letter(col_idx)
            if col_idx == 1:
                ws.column_dimensions[col_letter].width = 15
            else:
                ws.column_dimensions[col_letter].width = max(len(str(FINAL_COLUMNS[col_idx - 1]).strip()) + 3, 14)
    output.seek(0)
    return output


# ==============================================================================
# 5. APPLICATION STREAMLIT (INTERFACE UTILISATEUR COMPLÈTE)
# ==============================================================================

def main() -> None:
    st.set_page_config(
        page_title="Dashboard & Suivi STAPS | Rugby Club Amiénois",
        page_icon="🏉",
        layout="wide",
        initial_sidebar_state="collapsed",
    )

    # Styles CSS modernes et soignés (Esthétique Canva / Design épuré RCA)
    st.markdown(
        """
        <style>
        .stApp {
            font-family: 'Segoe UI', Montserrat, Tahoma, sans-serif;
            background-color: #F8FAFC;
        }
        .main-title {
            color: #1F4E79;
            font-size: 2.2rem;
            font-weight: 800;
            margin-bottom: 0.1rem;
            letter-spacing: -0.5px;
        }
        .sub-title {
            color: #4A5568;
            font-size: 1.05rem;
            margin-bottom: 1.6rem;
            font-weight: 400;
        }
        .section-header {
            color: #1F4E79;
            font-size: 1.45rem;
            font-weight: 700;
            margin-top: 1.2rem;
            margin-bottom: 0.3rem;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .section-desc {
            color: #64748B;
            font-size: 0.95rem;
            margin-bottom: 1.2rem;
        }
        .legend-badge-obj {
            background-color: #2E5A88;
            color: white;
            padding: 3px 8px;
            border-radius: 4px;
            font-weight: 600;
            font-size: 0.82rem;
        }
        .legend-badge-real {
            background-color: #A30000;
            color: white;
            padding: 3px 8px;
            border-radius: 4px;
            font-weight: 600;
            font-size: 0.82rem;
        }
        div[data-testid="stMetricValue"] {
            font-size: 1.6rem !important;
            color: #1F4E79 !important;
            font-weight: 700 !important;
        }
        .stTabs [data-baseweb="tab-list"] {
            gap: 10px;
            margin-bottom: 1rem;
        }
        .stTabs [data-baseweb="tab"] {
            height: 44px;
            white-space: pre-wrap;
            background-color: #FFFFFF;
            border-radius: 8px 8px 0px 0px;
            color: #4A5568;
            font-weight: 600;
            font-size: 1.02rem;
            padding: 8px 18px;
            border: 1px solid #E2E8F0;
            border-bottom: none;
            transition: all 0.2s ease;
        }
        .stTabs [aria-selected="true"] {
            background-color: #1F4E79 !important;
            color: #FFFFFF !important;
            border-color: #1F4E79 !important;
        }
        .stDownloadButton button {
            background-color: #1F4E79 !important;
            color: white !important;
            font-weight: 600 !important;
            font-size: 1.05rem !important;
            padding: 0.65rem 2rem !important;
            border-radius: 8px !important;
            border: none !important;
            width: 100% !important;
            transition: all 0.3s ease;
            box-shadow: 0 4px 10px rgba(31, 78, 121, 0.25) !important;
        }
        .stDownloadButton button:hover {
            background-color: #163654 !important;
            box-shadow: 0 6px 15px rgba(31, 78, 121, 0.35) !important;
        }

        /* Cartes Podium (Esthétique Canva / RCA) */
        .podium-card-container {
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 14px;
            padding: 18px 14px;
            box-shadow: 0 4px 14px rgba(0, 0, 0, 0.04);
            margin-bottom: 0.5rem;
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        .podium-card-container:hover {
            box-shadow: 0 6px 20px rgba(46, 90, 136, 0.12);
        }
        .podium-category-header {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            padding-bottom: 12px;
            border-bottom: 1px solid #F1F5F9;
            margin-bottom: 14px;
        }
        .podium-cat-icon {
            font-size: 1.3rem;
        }
        .podium-cat-title {
            font-size: 1.08rem;
            font-weight: 700;
            color: #1F4E79;
            letter-spacing: -0.2px;
        }
        .podium-winner-box {
            background: linear-gradient(180deg, #F0F7FF 0%, #FFFFFF 100%);
            border: 1.5px solid #BFDBFE;
            border-radius: 12px;
            padding: 14px 10px;
            text-align: center;
            margin-bottom: 12px;
            box-shadow: 0 3px 10px rgba(46, 90, 136, 0.08);
            position: relative;
        }
        .podium-rank-badge {
            display: inline-block;
            font-size: 0.8rem;
            font-weight: 700;
            padding: 2px 10px;
            border-radius: 12px;
            margin-bottom: 8px;
        }
        .rank-1 {
            background-color: #FEF3C7;
            color: #92400E;
            border: 1px solid #FDE68A;
        }
        .rank-2 {
            background-color: #F1F5F9;
            color: #475569;
            border: 1px solid #E2E8F0;
        }
        .rank-3 {
            background-color: #FFEDD5;
            color: #9A3412;
            border: 1px solid #FED7AA;
        }
        .avatar-wrapper {
            display: flex;
            justify-content: center;
            align-items: center;
            margin-bottom: 6px;
        }
        .podium-avatar {
            border-radius: 50%;
            object-fit: cover;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.1);
        }
        .img-winner {
            width: 78px;
            height: 78px;
            border: 3px solid #2E5A88;
        }
        .img-runner {
            width: 52px;
            height: 52px;
            border: 2px solid #CBD5E1;
        }
        .podium-player-name {
            font-weight: 700;
            color: #1E293B;
            font-size: 0.95rem;
            margin-top: 4px;
            line-height: 1.2;
        }
        .name-winner {
            font-size: 1.05rem;
            font-weight: 800;
            color: #1F4E79;
        }
        .podium-player-stat {
            color: #A30000;
            font-weight: 800;
            font-size: 1.15rem;
            margin-top: 3px;
            letter-spacing: -0.3px;
        }
        .stat-winner {
            font-size: 1.35rem;
            margin-top: 4px;
        }
        .podium-runners-row {
            display: flex;
            gap: 10px;
            justify-content: space-between;
        }
        .podium-runner-box {
            flex: 1;
            background-color: #F8FAFC;
            border: 1px solid #E2E8F0;
            border-radius: 10px;
            padding: 10px 6px;
            text-align: center;
            transition: background-color 0.2s ease;
        }
        .podium-runner-box:hover {
            background-color: #F1F5F9;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # Entête principal de l'application
    st.markdown('<div class="main-title">🏉 Rugby Club Amiénois | Suivi des Charges STAPS</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-title">Dashboard dynamique d\'analyse athlétique GPS & Injection automatisée dans le classeur de la préparatrice physique</div>',
        unsafe_allow_html=True,
    )

    # ==========================================================================
    # SECTION 1 : RETOUR JOURNÉE : COMPARAISON OBJECTIF VS RÉALISÉ (DASHBOARD)
    # ==========================================================================
    st.markdown('<div class="section-header">📊 RETOUR JOURNÉE : Comparaison Objectif vs Réalisé</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-desc">'
        'Rapport visuel par groupe positionnel : '
        '<span class="legend-badge-obj">■ Objectif Séance</span> &nbsp;'
        '<span class="legend-badge-real">■ Réalisé GPS</span> '
        '&nbsp;(Les pourcentages d\'accomplissement sont intégrés directement dans les barres)'
        '</div>',
        unsafe_allow_html=True,
    )

    # Filtrage par groupe de joueurs via st.tabs
    tab_avants, tab_3eme, tab_arrieres = st.tabs([
        "🏉 5 de devant",
        "🏉 3ème Ligne",
        "🏉 Trois-quarts"
    ])

    with tab_avants:
        render_dashboard_objectifs("5 de devant", MOCK_OBJECTIFS["5 de devant"])

    with tab_3eme:
        render_dashboard_objectifs("3ème Ligne", MOCK_OBJECTIFS["3ème Ligne"])

    with tab_arrieres:
        render_dashboard_objectifs("Trois-quarts", MOCK_OBJECTIFS["Trois-quarts"])

    # ==========================================================================
    # SECTION 2 : PODIUM : LES TOPS DE LA SÉANCE
    # ==========================================================================
    render_podium_section(MOCK_PODIUM)

    st.markdown("---")

    # ==========================================================================
    # SECTION 2 : IMPORTATION & TRAITEMENT DES EXPORTS GPS CATAPULT ONE
    # ==========================================================================
    st.markdown('<div class="section-header">📁 Traitement des Données GPS Catapult One</div>', unsafe_allow_html=True)

    # Vérification de la présence du Template officiel
    template_path = get_template_path()
    if not template_path:
        st.warning(
            f"⚠️ **Fichier modèle manquant :** Le fichier `{TEMPLATE_FILENAME}` est introuvable dans le même dossier que `app.py`. "
            "Veuillez déposer ce fichier modèle dans le répertoire pour permettre l'export avec tous les onglets et formules de la préparatrice physique."
        )

    uploaded_file = st.file_uploader(
        "Importer l'export brut CSV Catapult One / PlayerTek / OpenField :",
        type=["csv", "txt"],
        help="Détection dynamique de l'en-tête réel ('Player Name' ou 'Date') et conversion automatique du temps de jeu en minutes.",
    )

    if uploaded_file is not None:
        try:
            file_bytes = uploaded_file.getvalue()
            df_raw = load_catapult_csv(file_bytes)
            
            with st.spinner("Traitement des métriques GPS et calcul des indicateurs STAPS..."):
                df_indiv = process_catapult_data(df_raw)

            st.success(f"Traitement réussi ! {len(df_indiv)} joueurs identifiés et agrégés sur la session.")

            # Métriques clés globales
            kpi_cols = st.columns(4)
            with kpi_cols[0]:
                st.metric("Joueurs Analysés", len(df_indiv))
            with kpi_cols[1]:
                st.metric("Distance Cumulée", f"{df_indiv['Distance (km)'].sum():.2f} km")
            with kpi_cols[2]:
                st.metric("Vitesse Max Séance", f"{df_indiv['Top Speed (km/h)'].max():.1f} km/h")
            with kpi_cols[3]:
                st.metric("Total Haute Intensité", f"{df_indiv[' Distance haute intensité (m)'].sum():.0f} m")

            st.markdown("### 📋 Synthèse Individuelle STAPS (`df_indiv`)")
            
            # Formatage du DataFrame pour l'affichage interactif : arrondi à 2 décimales
            df_indiv['Temps de jeu '] = df_indiv['Temps de jeu '].astype(float)
            numeric_cols = df_indiv.select_dtypes(include=[np.number]).columns
            df_indiv[numeric_cols] = df_indiv[numeric_cols].round(2)
            
            st.dataframe(df_indiv, use_container_width=True, hide_index=True)

            # Export Excel : Injection dans le modèle officiel de la préparatrice
            st.markdown("### 📥 Télécharger le Fichier Excel de Séance")
            if template_path:
                excel_buffer = populate_template_excel(df_indiv=df_indiv, df_raw=df_raw)
                st.download_button(
                    label=f"Télécharger {OUTPUT_DOWNLOAD_FILENAME}",
                    data=excel_buffer.getvalue(),
                    file_name=OUTPUT_DOWNLOAD_FILENAME,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Télécharge le classeur officiel complet avec tous les onglets d'origine et les formules actualisées.",
                )
            else:
                fallback_buffer = create_fallback_excel(df_indiv=df_indiv)
                st.download_button(
                    label=f"Télécharger {OUTPUT_DOWNLOAD_FILENAME} (Format de secours)",
                    data=fallback_buffer.getvalue(),
                    file_name=OUTPUT_DOWNLOAD_FILENAME,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Fichier de secours généré sans les onglets modèles annexes.",
                )

        except Exception as e:
            st.error(f"Erreur lors du traitement du fichier Catapult One : {str(e)}")
            with st.expander("Détails techniques de l'erreur"):
                st.exception(e)
    else:
        st.info("👋 Veuillez téléverser un export brut CSV Catapult One pour générer la synthèse STAPS et mettre à jour le modèle.")


if __name__ == "__main__":
    main()
