"""
Application Streamlit - Traitement des Données GPS Catapult One & Suivi des Charges GPS
Club : Rugby Club Amiénois (RCA)
Auteur : Expert Python / Streamlit Senior en Data Visualisation Sportive
Description :
    Application web complète combinant :
    1. Un Uploader universel (.csv et .xlsx) en haut de page avec détection dynamique du vrai header.
    2. Un module de Mapping GPS dynamique et intelligent :
       - Scanne le DataFrame brut pour identifier les boîtiers réellement utilisés (ex: "1", "2", "Bam's", "1Er Mi Temps").
       - Affiche autant de sélecteurs qu'il y a d'identifiants uniques dans l'export.
       - Base de données officielle de l'effectif sénior RCA extraite du PDF (avec postes associés).
       - Barre de recherche intégrée avec format_func affichant "Nom Prénom - Poste" tout en retournant le Nom.
       - Prévisualisation instantanée de la photo du joueur.
    3. Un Dashboard dynamique (style Canva) de comparaison Objectif vs Réalisé connecté aux vraies données GPS
       avec filtrage par groupe de joueurs ("5 de devant", "3ème Ligne", "Trois-quarts").
    4. Un Podium dynamique des Tops de la séance (Top speed, % marche, Haute intensité) calculé en temps réel,
       avec photos réelles des joueurs encodées en Base64.
    5. La conversion précise des durées en minutes réelles et l'agrégation par joueur (groupby).
    6. L'injection directe dans le classeur modèle officiel (Modele_Vierge_STAPS.xlsx) pour préserver
       toutes les formules natives de la préparatrice physique, graphiques et onglets.
"""

from __future__ import annotations

import base64
import datetime
import io
import json
import os
import re
import textwrap
import urllib.parse
from typing import Any, Dict, List, Optional, Tuple, Union

try:
    from streamlit_mic_recorder import speech_to_text
    HAS_MIC_RECORDER = True
except Exception:
    HAS_MIC_RECORDER = False

from fpdf import FPDF
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ==============================================================================
# 0. EFFECTIF SÉNIORS RCA & CONFIGURATION ATHLÉTIQUE
# ==============================================================================

TEMPLATE_FILENAME = "Modele_Vierge_STAPS.xlsx"
OUTPUT_DOWNLOAD_FILENAME = "Synthese_Complete_Seance.xlsx"

# Dossier contenant les photos des joueurs (ex: DUPONT_Antoine.jpg)
PHOTOS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "photos_joueurs")
DEFAULT_AVATAR = os.path.join(PHOTOS_DIR, "avatar_par_defaut.jpg")

# Palette de couleurs officielles Rugby Club Amiénois / Style Canva
COLOR_OBJECTIF = "#2E5A88"  # Bleu roi / marine RCA
COLOR_REALISE = "#A30000"   # Rouge / Bordeaux RCA
COLOR_TEXT_MAIN = "#1F4E79"
COLOR_MUTED = "#555555"

# ------------------------------------------------------------------------------
# 1. BASE DE DONNÉES EFFECTIF SÉNIORS DU RUGBY CLUB AMIÉNOIS (EXTRACTION PDF)
# ------------------------------------------------------------------------------
joueurs_data: Dict[str, str] = {
    # Piliers (5 de devant)
    "SALUA Landry": "Pilier",
    "SALUA Noam": "Pilier",
    "FRION Quentin": "Pilier",
    "COZIC Adrien": "Pilier",
    "PECQUERY Thomas": "Pilier",
    "TRAMON Noam": "Pilier",
    "FOURNET Giovanni": "Pilier",
    "MAZE Nolan": "Pilier",
    "BICHARI Jérémy": "Pilier",
    "LEFEVER Thibault": "Pilier",
    "ROGIER Arnaud": "Pilier",
    "VASELLI Sébastien": "Pilier",
    "TANTER Edouard": "Pilier",
    "DAULT Bastien": "Pilier",
    "SCHNEERBERGER Julien": "Pilier",
    "DUGROPREZ Julien": "Pilier",
    # Talonneurs (5 de devant)
    "LESUEUR Jean-Étienne": "Talonneur",
    "Jean-Étienne ( TIJ)": "Talonneur",
    "GENERMONT Roméo": "Talonneur",
    "GUIBERT Samuel": "Talonneur",
    "GIET Axel": "Talonneur",
    "D. Anthymme": "Talonneur",
    # 2èmes lignes (5 de devant)
    "WAEYAERT Thomas": "2ème ligne",
    "GRENET Paul": "2ème ligne",
    "BOIDIN Donovan": "2ème ligne",
    "VDH Simon": "2ème ligne",
    "DENIER Gabin": "2ème ligne",
    "Gabin D.": "2ème ligne",
    "ROY Etienne": "2ème ligne",
    "FOUQUET Victorien": "2ème ligne",
    # 3èmes lignes (Alignement Préparatrice Physique RCA : SOTILLO Alejo et SAINTOMER Lucas inclus)
    "SOTILLO Alejo": "3ème ligne",
    "SAINTOMER Lucas": "3ème ligne",
    "SAITOMER Lucas": "3ème ligne",
    "SOTILLO Lucas": "3ème ligne",
    "VASELLI Enzo": "3ème ligne",
    "ORION Lucas": "3ème ligne",
    "OLIER Lucas": "3ème ligne",
    "CARPENTIER Simon": "3ème ligne",
    "LAFEUILLE Emilien": "3ème ligne",
    "MARTIN Benjamin": "3ème ligne",
    "MONTI Ciro": "3ème ligne",
    "WIOTTE Louis": "3ème ligne",
    "GENTIEN Tim": "3ème ligne",
    "CARLIER Louis": "3ème ligne",
    "COSSARD Léandre": "3ème ligne",
    "SAINTOMER Louis": "3ème ligne",
    "AUDIERE Adrien": "3ème ligne",
    "BINNINGER Alban": "3ème ligne",
    "BLANC Jules": "3ème ligne",
    "KAZUMBA Camille": "3ème ligne",
    # Demis de mêlée (Trois-quarts)
    "TESSON Gabriel": "Demi de mêlée",
    "BAMIERE Clément": "Demi de mêlée",
    "BAM'S": "Demi de mêlée",
    "SANTONI Stephan": "Demi de mêlée",
    "BOEUF Elliot": "Demi de mêlée",
    "BŒUF Eliot": "Demi de mêlée",
    "MAUGARD Nicolas": "Demi de mêlée",
    "BOUMENDIL Antoine": "Demi de mêlée",
    # Demis d'ouverture (Trois-quarts)
    "HARRY Max": "Demi d'ouverture",
    "HARRY Maxence": "Demi d'ouverture",
    "CARPENTIER Maël": "Demi d'ouverture",
    "ULMUCK Victor": "Demi d'ouverture",
    "USUBILLAGA Victor": "Demi d'ouverture",
    "GOMES Louis": "Demi d'ouverture",
    "SUEUR Xavier": "Demi d'ouverture",
    # Centres (Trois-quarts)
    "PRÉVOTÉ Oliver": "Centre",
    "PREVOTE Olivier": "Centre",
    "PELESEUMA Ope": "Centre",
    "OPE": "Centre",
    "TREMBLAY Brieuc": "Centre",
    "DE BRYUN Octave": "Centre",
    "DE BRUYN Octave": "Centre",
    "VENIN Benoît": "Centre",
    "HURTEL Oscar": "Centre",
    "Oscar": "Centre",
    "DELAVALLÉE Quentin": "Centre",
    "MANDIVENGA Brendon": "Centre",
    # Ailiers (Trois-quarts - dont Jules D. / DELACROIX Jules)
    "DELACROIX Jules": "Ailier",
    "CARPENTIER Joseph": "Ailier",
    "SALUA Maka": "Ailier",
    "SALUA Makatea": "Ailier",
    "SELTZER Antoine": "Ailier",
    "BATON Léonard": "Ailier",
    "BATTON Léonard": "Ailier",
    "MARZIANO Nathan": "Ailier",
    "SRIDI Youness": "Ailier",
    "LEROUX Mathis": "Ailier",
    "GANE Thomas": "Ailier",
    "DELATTRE Jacques": "Ailier",
    # Arrières (Trois-quarts)
    "LESPINASSE Lucien": "Arrière",
    "PRÉVOST Thomas": "Arrière",
    "PREVOST Thomas": "Arrière",
    "MOINE Arthur": "Arrière",
    "VILAIN Beryl": "Arrière",
}


def mapper_poste_vers_groupe_rca(poste_detaille: str, nom_joueur: str = "") -> str:
    """
    Convertit automatiquement et de façon robuste un poste détaillé du rugby en l'un
    des 3 grands groupes d'analyse officiels de la préparatrice physique du RCA :
    1. '5 de devant' : Pilier, Pilier droit, Pilier gauche, Talonneur, 2ème ligne, Deuxième ligne, Première ligne, 5 de devant.
    2. '3ème Ligne' : 3ème ligne, Troisième ligne, 3ème ligne aile, 3ème ligne centre, Flanker, Numéro 8.
    3. 'Trois-quarts' : Demis (mêlée/ouverture/9/10), Ouvrière, Demis, Centre, Ailier, Arrière, Trois-quarts, 3/4.

    Prend également en compte les alignements prioritaires demandés par la préparatrice physique :
    - Lucas S. et Alejo S. rattachés en '3ème Ligne'.
    - Bastien D., Thomas P., Edouard T., Jean-Étienne L. en '5 de devant'.
    - Enzo V., Lucas O., Simon C. en '3ème Ligne'.
    - Jules D. (DELACROIX Jules), Elliot B., Maël C., Victor U., Olivier P., Ope P., Maka S., Arthur M., Octave D., Lucien L. en 'Trois-quarts'.
    """
    nom_clean = str(nom_joueur).strip().lower() if nom_joueur else ""

    # 1. Règle d'alignement prioritaire par joueur spécifique (Staff RCA)
    if nom_clean:
        # Lucas S. & Alejo S. -> 3ème Ligne (selon méthodologie préparatrice physique RCA)
        if "sotillo" in nom_clean:
            return "3ème Ligne"
        if ("saintomer" in nom_clean or "saitomer" in nom_clean) and "lucas" in nom_clean:
            return "3ème Ligne"

        # 5 de devant spécifiques
        if any(k in nom_clean for k in [
            "dault bastien", "bastien d", "dault",
            "pecquery thomas", "thomas p", "pecquery",
            "tanter edouard", "edouard t", "tanter",
            "lesueur jean-étienne", "lesueur jean-etienne", "jean-étienne l", "jean-etienne l", "lesueur", "tij",
            "cozic adrien", "cozic"
        ]):
            return "5 de devant"

        # 3ème Ligne spécifiques
        if any(k in nom_clean for k in [
            "vaselli enzo", "enzo v",
            "orion lucas", "olier lucas", "lucas o",
            "carpentier simon", "simon c",
            "blanc jules", "lafeuille emilien", "monti ciro"
        ]):
            return "3ème Ligne"

        # Trois-quarts spécifiques
        if any(k in nom_clean for k in [
            "delacroix jules", "jules d", "delacroix",
            "boeuf elliot", "bœuf eliot", "boeuf eliot", "bœuf elliot", "elliot b", "eliot b",
            "carpentier maël", "carpentier mael", "maël c", "mael c",
            "ulmuck victor", "usubillaga victor", "victor u", "ulmuck", "usubillaga",
            "prévoté oliver", "prevote olivier", "olivier p", "oliver p", "prevote", "prévoté",
            "peleseuma ope", "ope p", "ope",
            "salua maka", "salua makatea", "maka s",
            "moine arthur", "arthur m",
            "de bruyn octave", "de bryun octave", "octave d",
            "lespinasse lucien", "lucien l", "lespinasse",
            "bam's", "bamiere", "harry max", "harry maxence"
        ]):
            return "Trois-quarts"

    # 2. Conversion automatique par poste détaillé
    p_clean = str(poste_detaille).strip().lower()
    if not p_clean:
        return "Trois-quarts"

    # 5 de devant
    # Pilier, Pilier droit, Pilier gauche, Talonneur, 2ème ligne, Deuxième ligne, Première ligne, 5 de devant
    if any(k in p_clean for k in [
        "pilier", "prop", "talon", "hooker",
        "2ème ligne", "2eme ligne", "deuxième ligne", "deuxieme ligne", "2nd ligne", "2e ligne", "lock",
        "première ligne", "premiere ligne", "1ère ligne", "1ere ligne",
        "5 de devant", "cinq de devant", "front row", "tight five"
    ]):
        return "5 de devant"

    # 3ème Ligne
    # 3ème ligne, Troisième ligne, 3ème ligne aile, 3ème ligne centre, Flanker, Numéro 8
    if any(k in p_clean for k in [
        "3ème ligne", "3eme ligne", "troisième ligne", "troisieme ligne", "3e ligne",
        "3ème ligne aile", "3eme ligne aile", "3ème ligne centre", "3eme ligne centre",
        "flanker", "flank", "troisième ligne aile", "troisième ligne centre",
        "numéro 8", "numero 8", "num 8", "n°8", "no8", "n8",
        "third row", "back row", "loose forward"
    ]) or p_clean == "8":
        return "3ème Ligne"

    # Trois-quarts
    # Demi de mêlée (9), Demi d'ouverture (10), Ouvrière, Demis, Centre, Ailier, Arrière, Trois-quarts, 3/4
    if any(k in p_clean for k in [
        "demi de mêlée", "demi de melee", "demi melee", "scrum half", "scrum-half",
        "demi d'ouverture", "demi d ouverture", "demi ouverture", "fly half", "fly-half", "ouvreur", "ouvrière", "ouvriere",
        "demi", "demis",
        "centre", "center", "premier centre", "second centre",
        "ailier", "wing", "winger",
        "arrière", "arriere", "fullback", "full back",
        "trois-quarts", "trois quarts", "trois-quart", "trois quart", "3/4", "backs"
    ]) or p_clean in ["9", "10", "11", "12", "13", "14", "15"]:
        return "Trois-quarts"

    # 3. Règles heuristiques de repli
    if "5" in p_clean or "devant" in p_clean:
        return "5 de devant"
    if "3" in p_clean:
        return "3ème Ligne"

    return "Trois-quarts"


# Cartographie Poste -> Groupe athlétique officiel RCA (3 catégories)
POSTE_VERS_GROUPE: Dict[str, str] = {
    # 5 de devant
    "Pilier": "5 de devant",
    "Pilier droit": "5 de devant",
    "Pilier gauche": "5 de devant",
    "Talonneur": "5 de devant",
    "2ème ligne": "5 de devant",
    "Deuxième ligne": "5 de devant",
    "Première ligne": "5 de devant",
    "5 de devant": "5 de devant",
    # 3ème Ligne
    "3ème ligne": "3ème Ligne",
    "Troisième ligne": "3ème Ligne",
    "3ème ligne aile": "3ème Ligne",
    "3ème ligne centre": "3ème Ligne",
    "Flanker": "3ème Ligne",
    "Numéro 8": "3ème Ligne",
    # Trois-quarts
    "Demi de mêlée": "Trois-quarts",
    "Demi d'ouverture": "Trois-quarts",
    "Ouvrière": "Trois-quarts",
    "Demis": "Trois-quarts",
    "Centre": "Trois-quarts",
    "Centres": "Trois-quarts",
    "Ailier": "Trois-quarts",
    "Ailiers": "Trois-quarts",
    "Arrière": "Trois-quarts",
    "Arrières": "Trois-quarts",
    "Trois-quarts": "Trois-quarts",
}

# Objectifs athlétiques de référence par groupe (Moyenne individuelle cible par séance)
OBJECTIFS_CONFIG: Dict[str, Dict[str, float]] = {
    "5 de devant": {
        "Objectif Distance (m)": 4500.0,
        "Obj. HSR": 250.0,
        "Obj. VHSR": 50.0,
        "Obj. ACC/DEC": 35.0,
    },
    "3ème Ligne": {
        "Objectif Distance (m)": 5500.0,
        "Obj. HSR": 450.0,
        "Obj. VHSR": 100.0,
        "Obj. ACC/DEC": 45.0,
    },
    "Trois-quarts": {
        "Objectif Distance (m)": 6200.0,
        "Obj. HSR": 700.0,
        "Obj. VHSR": 250.0,
        "Obj. ACC/DEC": 50.0,
    },
    "Tous les groupes": {
        "Objectif Distance (m)": 5400.0,
        "Obj. HSR": 460.0,
        "Obj. VHSR": 130.0,
        "Obj. ACC/DEC": 43.0,
    },
}

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

# Données de simulation avec joueurs réels du RCA
MOCK_DEMO_DATA: List[Dict[str, Any]] = [
    {"Player Name": "BAMIERE Clément", "Position": "Demi de mêlée", "Distance (km)": 6.3, "Zone 4 (km)": 0.48, "Zone 5 (km)": 0.24, "Top Speed": 32.4, "Duration": "01:22:00", "Z1 (km)": 2.5, "Z2 (km)": 2.2, "Z3 (km)": 0.88, "Acc34": 18, "Acc4": 8, "Dec34": 14, "Dec4": 6},
    {"Player Name": "PRÉVOST Thomas", "Position": "Arrière", "Distance (km)": 6.6, "Zone 4 (km)": 0.52, "Zone 5 (km)": 0.28, "Top Speed": 31.8, "Duration": "01:22:00", "Z1 (km)": 2.6, "Z2 (km)": 2.3, "Z3 (km)": 0.90, "Acc34": 20, "Acc4": 7, "Dec34": 15, "Dec4": 5},
    {"Player Name": "HARRY Max", "Position": "Demi d'ouverture", "Distance (km)": 6.0, "Zone 4 (km)": 0.44, "Zone 5 (km)": 0.20, "Top Speed": 31.1, "Duration": "01:22:00", "Z1 (km)": 2.4, "Z2 (km)": 2.1, "Z3 (km)": 0.86, "Acc34": 16, "Acc4": 6, "Dec34": 13, "Dec4": 4},
    {"Player Name": "VASELLI Enzo", "Position": "3ème ligne", "Distance (km)": 5.6, "Zone 4 (km)": 0.41, "Zone 5 (km)": 0.11, "Top Speed": 29.5, "Duration": "01:22:00", "Z1 (km)": 2.6, "Z2 (km)": 1.8, "Z3 (km)": 0.68, "Acc34": 16, "Acc4": 6, "Dec34": 13, "Dec4": 5},
    {"Player Name": "MARTIN Benjamin", "Position": "3ème ligne", "Distance (km)": 5.4, "Zone 4 (km)": 0.38, "Zone 5 (km)": 0.09, "Top Speed": 28.8, "Duration": "01:22:00", "Z1 (km)": 2.7, "Z2 (km)": 1.7, "Z3 (km)": 0.63, "Acc34": 15, "Acc4": 5, "Dec34": 12, "Dec4": 4},
    {"Player Name": "MONTI Ciro", "Position": "3ème ligne", "Distance (km)": 5.5, "Zone 4 (km)": 0.39, "Zone 5 (km)": 0.10, "Top Speed": 29.0, "Duration": "01:22:00", "Z1 (km)": 2.7, "Z2 (km)": 1.7, "Z3 (km)": 0.61, "Acc34": 15, "Acc4": 5, "Dec34": 12, "Dec4": 4},
    {"Player Name": "LESUEUR Jean-Étienne", "Position": "Talonneur", "Distance (km)": 4.6, "Zone 4 (km)": 0.17, "Zone 5 (km)": 0.04, "Top Speed": 26.1, "Duration": "01:18:00", "Z1 (km)": 2.9, "Z2 (km)": 1.3, "Z3 (km)": 0.23, "Acc34": 14, "Acc4": 4, "Dec34": 11, "Dec4": 3},
    {"Player Name": "COZIC Adrien", "Position": "Pilier", "Distance (km)": 4.4, "Zone 4 (km)": 0.15, "Zone 5 (km)": 0.03, "Top Speed": 25.2, "Duration": "01:15:00", "Z1 (km)": 2.8, "Z2 (km)": 1.2, "Z3 (km)": 0.22, "Acc34": 12, "Acc4": 3, "Dec34": 10, "Dec4": 2},
    {"Player Name": "SALUA Landry", "Position": "Pilier", "Distance (km)": 4.2, "Zone 4 (km)": 0.12, "Zone 5 (km)": 0.02, "Top Speed": 24.5, "Duration": "01:15:00", "Z1 (km)": 2.8, "Z2 (km)": 1.1, "Z3 (km)": 0.18, "Acc34": 11, "Acc4": 2, "Dec34": 9, "Dec4": 2},
]


def get_template_path() -> Optional[str]:
    """Recherche le fichier modèle Modele_Vierge_STAPS.xlsx."""
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
# 1. GESTION DES PHOTOS JOUEURS & ENCODAGE BASE64
# ==============================================================================

def obtenir_chemin_photo(nom_joueur: str) -> Optional[str]:
    """
    Retourne le chemin de la photo du joueur dans photos_joueurs/.
    Gère les variantes de nommage (underscores, casse, tirets).
    """
    if not nom_joueur or nom_joueur.strip() == "":
        return DEFAULT_AVATAR if os.path.isfile(DEFAULT_AVATAR) else None

    nom = nom_joueur.strip()
    variantes = [
        nom.replace(" ", "_"),            # DUPONT_Antoine
        nom.replace(" ", "_").lower(),    # dupont_antoine
        nom.replace(" ", "_").title(),    # Dupont_Antoine
        nom.replace(" ", "-"),            # DUPONT-Antoine
        nom,                              # DUPONT Antoine
    ]

    extensions = [".jpg", ".jpeg", ".png", ".webp"]

    for v in variantes:
        for ext in extensions:
            chemin = os.path.join(PHOTOS_DIR, f"{v}{ext}")
            if os.path.isfile(chemin):
                return chemin

    # Recherche par nom de famille si le fichier contient le premier mot
    nom_famille = nom.split()[0].lower() if " " in nom else nom.lower()
    try:
        for f in os.listdir(PHOTOS_DIR):
            if nom_famille in f.lower() and os.path.splitext(f)[1].lower() in extensions:
                return os.path.join(PHOTOS_DIR, f)
    except FileNotFoundError:
        pass

    return DEFAULT_AVATAR if os.path.isfile(DEFAULT_AVATAR) else None


def photo_to_base64_data_uri(nom_joueur: str) -> str:
    """
    Convertit la photo du joueur en Data URI Base64 pour le HTML.
    Si absente, retourne un avatar SVG Dicebear dynamique.
    """
    chemin = obtenir_chemin_photo(nom_joueur)
    if chemin and os.path.isfile(chemin):
        ext = os.path.splitext(chemin)[1].lower()
        mime_map = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".webp": "image/webp"}
        mime_type = mime_map.get(ext, "image/jpeg")
        try:
            with open(chemin, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("utf-8")
            return f"data:{mime_type};base64,{b64}"
        except Exception:
            pass

    seed = urllib.parse.quote(nom_joueur.strip()) if nom_joueur else "RCA"
    return f"https://api.dicebear.com/7.x/initials/svg?seed={seed}&backgroundColor=2E5A88"


# ==============================================================================
# 2. DATA VISUALISATION PLOTLY & PODIUM (STYLE CANVA ÉPURÉ)
# ==============================================================================

def creer_graphique_jauge(titre: str, val_objectif: float, val_realise: float) -> go.Figure:
    """
    Génère un graphique à barres stylisé inspiré du design Canva RCA :
    - Deux barres côte à côte : 'Objectif' (Bleu roi) et 'Réalisé' (Rouge RCA).
    - Valeurs et pourcentages d'accomplissement intégrés.
    """
    val_obj = float(val_objectif)
    val_real = float(val_realise)
    pct = (val_real / val_obj * 100) if val_obj > 0 else 0.0

    fmt_obj = f"{int(round(val_obj)):,d}".replace(",", " ") if abs(val_obj - round(val_obj)) < 0.01 else f"{val_obj:.1f}"
    fmt_real = f"{int(round(val_real)):,d}".replace(",", " ") if abs(val_real - round(val_real)) < 0.01 else f"{val_real:.1f}"

    max_val = max(val_obj, val_real, 1.0)
    y_max = max_val * 1.30

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=["Objectif", "Réalisé"],
        y=[val_obj, val_real],
        marker=dict(
            color=[COLOR_OBJECTIF, COLOR_REALISE],
            line=dict(width=0),
            cornerradius=8,
        ),
        text=[f"<b>{fmt_obj}</b>", f"<b>{fmt_real}</b>"],
        textposition="outside",
        textfont=dict(family="Segoe UI, Montserrat, Inter, sans-serif", size=13, color="#1A2B4C"),
        width=[0.42, 0.42],
        hovertemplate="<b>%{x}</b> : %{text}<extra></extra>",
        showlegend=False,
    ))

    if val_real > 0.14 * max_val:
        y_pct = val_real * 0.50
        font_color = "#FFFFFF"
        font_size = 19
    else:
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
    """Affiche 4 graphiques stylisés sur une seule ligne (st.columns(4))."""
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
                config={"displayModeBar": False},
            )


def render_podium_section(podium_data: Dict[str, Dict[str, Any]]) -> None:
    """
    Affiche la section Podium avec les Tops de la séance :
    - 3 catégories : Top speed, % marche, Haute intensité.
    - Avatars réels des joueurs encodés en Base64.
    """
    st.divider()
    st.markdown('<div class="section-header">🏆 PODIUM</div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    cols = [col1, col2, col3]
    categories = list(podium_data.keys())

    for idx, cat_name in enumerate(categories):
        cat_info = podium_data[cat_name]
        players = cat_info["players"]
        p1, p2, p3 = players[0], players[1], players[2]

        raw_html = textwrap.dedent(f"""
            <div class="podium-card-container">
                <div class="podium-category-header">
                    <span class="podium-cat-icon">{cat_info.get('icon', '🏉')}</span>
                    <span class="podium-cat-title">{cat_name}</span>
                </div>
                <div class="podium-winner-box">
                    <div class="podium-rank-badge rank-1">🥇 1er</div>
                    <div class="avatar-wrapper">
                        <img src="{p1['avatar_src']}"
                             alt="{p1['name']}" class="podium-avatar img-winner" />
                    </div>
                    <div class="podium-player-name name-winner">{p1['name']}</div>
                    <div class="podium-player-stat stat-winner">{p1['stat']}</div>
                </div>
                <div class="podium-runners-row">
                    <div class="podium-runner-box">
                        <div class="podium-rank-badge rank-2">🥈 2ème</div>
                        <div class="avatar-wrapper">
                            <img src="{p2['avatar_src']}"
                                 alt="{p2['name']}" class="podium-avatar img-runner" />
                        </div>
                        <div class="podium-player-name">{p2['name']}</div>
                        <div class="podium-player-stat">{p2['stat']}</div>
                    </div>
                    <div class="podium-runner-box">
                        <div class="podium-rank-badge rank-3">🥉 3ème</div>
                        <div class="avatar-wrapper">
                            <img src="{p3['avatar_src']}"
                                 alt="{p3['name']}" class="podium-avatar img-runner" />
                        </div>
                        <div class="podium-player-name">{p3['name']}</div>
                        <div class="podium-player-stat">{p3['stat']}</div>
                    </div>
                </div>
            </div>
        """)

        card_html = "\n".join(line.strip() for line in raw_html.splitlines() if line.strip())
        with cols[idx]:
            st.markdown(card_html, unsafe_allow_html=True)


# ==============================================================================
# 3. PARSING ROBUSTE & CONVERSION NUMÉRIQUE / TEMPORELLE
# ==============================================================================

def parse_duration_to_minutes(val: Any) -> float:
    """Convertit une valeur de durée individuelle en minutes réelles (float)."""
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
    """Convertit une série complète de durées Catapult en minutes réelles (float)."""
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
    """Convertit une valeur quelconque avec virgule ou espace en float."""
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
    """Recherche flexible d'une colonne insensible aux espaces et à la casse."""
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


def find_catapult_header_row(file_bytes: bytes, encoding: str = "utf-8", max_scan: int = 15) -> int:
    """Scanne les 15 premières lignes pour repérer dynamiquement l'en-tête réel."""
    try:
        text_preview = file_bytes[:16384].decode(encoding, errors="replace")
        lines = text_preview.splitlines()[:max_scan]
    except Exception:
        return 0

    for idx, raw_line in enumerate(lines):
        line = raw_line.strip().lower()
        if not line:
            continue
        if any(term in line for term in ["player name", "athlete name", "nom ; prénom", "nom;prénom", "player", "athlete"]):
            return idx

    for idx, raw_line in enumerate(lines):
        line = raw_line.strip().lower()
        if not line:
            continue
        if "date" in line and any(term in line for term in ["session", "split", "distance", "duration", "speed"]):
            return idx

    return 0


def load_catapult_csv(file_bytes: bytes) -> pd.DataFrame:
    """Lit un fichier CSV Catapult One avec détection automatique d'encodage et délimiteur."""
    encodings_to_try: List[str] = ["utf-8-sig", "utf-8", "latin1", "cp1252", "iso-8859-1"]
    separators_to_try: List[str] = [";", ",", "\t"]
    last_error: Optional[Exception] = None

    for enc in encodings_to_try:
        header_idx = find_catapult_header_row(file_bytes, encoding=enc, max_scan=15)
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
                        return df.dropna(how="all")
            except Exception as exc:
                last_error = exc
                continue

    if last_error is not None:
        raise last_error
    raise ValueError("Impossible de décoder le fichier CSV avec les encodages et délimiteurs standards.")


def load_catapult_file(uploaded_file: Any) -> pd.DataFrame:
    """Lit un export Catapult One (.csv, .xlsx, .xls) en sautant les métadonnées."""
    file_name = uploaded_file.name.lower()
    file_bytes = uploaded_file.getvalue()

    if file_name.endswith((".xlsx", ".xls")):
        preview_df = pd.read_excel(io.BytesIO(file_bytes), header=None, nrows=15)
        header_idx = 0
        for idx, row in preview_df.iterrows():
            row_str = " ".join([str(v).lower() for v in row if pd.notna(v)])
            if any(term in row_str for term in ["player name", "athlete name", "nom ; prénom", "nom;prénom", "athlete", "player"]):
                header_idx = idx
                break
            elif "date" in row_str and any(term in row_str for term in ["session", "split", "distance", "duration"]):
                header_idx = idx
                break

        df = pd.read_excel(io.BytesIO(file_bytes), header=header_idx)
        df.columns = [str(c).strip() for c in df.columns]
        return df.dropna(how="all")
    else:
        return load_catapult_csv(file_bytes)


# ==============================================================================
# 4. MAPPING GPS DYNAMIQUE & INTELLIGENT (EFFECTIF DU CLUB)
# ==============================================================================

def ui_attribution_gps(df_raw: Optional[pd.DataFrame] = None) -> Dict[Union[int, str], str]:
    """
    Interface dynamique d'attribution des boîtiers GPS :
    - Filtre et extrait uniquement les numéros uniques de boîtiers GPS présents dans df_raw.
    - Convertit les identifiants en entiers (int) et les trie par ordre croissant (1, 2, 3...).
    - Itère DIRECTEMENT sur la liste des identifiants détectés (pas de range()).
    - S'il manque un boîtier (ex: le 18), passe directement au suivant (17 puis 19).
    - Affiche EXACTEMENT le même nombre de sélecteurs que de boîtiers détectés (ex: 21 sélecteurs pour 21 boîtiers).
    - Utilise format_func pour afficher 'Nom Prénom - Poste' tout en retournant le nom du joueur.
    - Affiche la photo officielle du joueur sélectionné ou l'avatar RCA.
    """
    st.markdown('<div class="section-header">📡 Attribution des Boîtiers GPS aux Joueurs</div>', unsafe_allow_html=True)

    # 1. Extraction et tri numérique strict des identifiants GPS détectés
    liste_gps_detectes: List[int] = []
    if df_raw is not None and not df_raw.empty:
        col_id = find_column(df_raw, [
            "Player Name", "Athlete Name", "Nom ; Prénom", "Player", "Athlete", "Joueur", "Nom", "Name"
        ])
        if col_id and col_id in df_raw.columns:
            nums_trouves: set[int] = set()
            for val in df_raw[col_id].dropna():
                val_str = str(val).strip()
                # Conversion directe si valeur numérique (int, float, ou str numérique)
                try:
                    num_float = float(val_str)
                    if num_float.is_integer() and num_float > 0:
                        nums_trouves.add(int(num_float))
                        continue
                except (ValueError, TypeError):
                    pass

                # Extraction si chaîne contenant un identifiant numérique (ex: "GPS 17")
                match = re.search(r"\b(\d+)\b", val_str)
                if match:
                    try:
                        n = int(match.group(1))
                        if n > 0:
                            nums_trouves.add(n)
                    except (ValueError, TypeError):
                        pass

            liste_gps_detectes = sorted(list(nums_trouves))

    # Si aucun boîtier numérique n'a été détecté dans le fichier
    if not liste_gps_detectes:
        st.info("ℹ️ Aucun numéro de boîtier GPS à attribuer (les joueurs sont déjà identifiés ou aucun numéro n'est présent).")
        return {}

    st.caption(
        f"**{len(liste_gps_detectes)} boîtier(s) GPS actif(s) détecté(s)** sur cette séance. "
        "Attribuez chaque boîtier à son joueur :"
    )

    # 2. Liste des joueurs RCA disponibles (tri alphabétique)
    placeholder = "— Sélectionner un joueur —"
    noms_joueurs_tries = sorted(list(joueurs_data.keys()))
    options = [placeholder] + noms_joueurs_tries

    def format_affichage_joueur(nom: str) -> str:
        """Affiche 'Nom Prénom - Poste' dans la liste déroulante et la barre de recherche."""
        if nom == placeholder:
            return placeholder
        poste = joueurs_data.get(nom, "")
        if poste:
            return f"{nom} - {poste}"
        return nom

    if "gps_mapping" not in st.session_state:
        st.session_state["gps_mapping"] = {}

    mapping: Dict[Union[int, str], str] = {}
    col_gauche, col_droite = st.columns(2)
    split_index = (len(liste_gps_detectes) + 1) // 2

    # 3. Boucle STRICTE sur la liste des identifiants détectés (aucun range() utilisé)
    for idx, gps_id in enumerate(liste_gps_detectes):
        conteneur = col_gauche if idx < split_index else col_droite

        with conteneur:
            c_photo, c_select = st.columns([1, 3])

            prev = (
                st.session_state["gps_mapping"].get(gps_id)
                or st.session_state["gps_mapping"].get(str(gps_id))
            )
            default_idx = options.index(prev) if prev in options else 0

            with c_select:
                choix = st.selectbox(
                    f"Boîtier GPS n°{gps_id}",
                    options=options,
                    index=default_idx,
                    format_func=format_affichage_joueur,
                    key=f"gps_sel_boitier_{gps_id}",
                )

            with c_photo:
                if choix != placeholder:
                    photo = obtenir_chemin_photo(choix)
                    if photo and os.path.isfile(photo):
                        st.image(photo, width=48)
                    else:
                        st.markdown("👤", help=choix)
                else:
                    st.markdown(
                        "<div style='width:48px;height:48px;border-radius:50%;"
                        "background:#E2E8F0;display:flex;align-items:center;"
                        "justify-content:center;color:#94A3B8;font-size:1.4rem;'>?</div>",
                        unsafe_allow_html=True,
                    )

            if choix != placeholder:
                mapping[gps_id] = choix
                mapping[str(gps_id)] = choix
                st.session_state["gps_mapping"][gps_id] = choix
                st.session_state["gps_mapping"][str(gps_id)] = choix

    # 4. Sécurité anti-doublon lors de l'attribution manuelle des GPS
    joueurs_assignes: Dict[str, List[int]] = {}
    for g_id in liste_gps_detectes:
        j_choisi = mapping.get(g_id)
        if j_choisi and j_choisi != placeholder:
            joueurs_assignes.setdefault(j_choisi, []).append(g_id)

    doublons = {j: ids for j, ids in joueurs_assignes.items() if len(ids) > 1}
    if doublons:
        details_doublons = ", ".join([
            f"**{j}** (boîtiers {', '.join(f'n°{i}' for i in ids)})"
            for j, ids in doublons.items()
        ])
        st.warning(
            f"⚠️ **Attention - Doublon détecté dans l'attribution GPS :** {details_doublons}. "
            "Un même joueur est assigné à plusieurs boîtiers différents sur la même séance. "
            "Veuillez corriger la sélection afin d'éviter la fusion erronée des distances et la réduction artificielle de l'effectif."
        )

    return mapping


def appliquer_mapping(df_catapult: pd.DataFrame, dictionnaire_gps: Dict[Union[int, str], str]) -> pd.DataFrame:
    """
    Remplace les numéros de boîtiers GPS dans la colonne joueur
    par les noms officiels sélectionnés de l'effectif.
    Gère de façon robuste les types int, float (1.0), et str ('1', '1.0').
    """
    if not dictionnaire_gps:
        return df_catapult

    df = df_catapult.copy()
    player_col = find_column(df, [
        "Player Name", "Athlete Name", "Nom ; Prénom", "Player", "Athlete", "Joueur", "Nom", "Name"
    ])
    if not player_col:
        return df

    # Normalisation du dictionnaire avec toutes les variantes de clés
    dict_map: Dict[Any, str] = {}
    for k, v in dictionnaire_gps.items():
        if not v:
            continue
        dict_map[k] = v
        dict_map[str(k).strip()] = v
        try:
            k_num = float(str(k).strip())
            if k_num.is_integer():
                k_int = int(k_num)
                dict_map[k_int] = v
                dict_map[float(k_int)] = v
                dict_map[str(k_int)] = v
                dict_map[f"{k_int}.0"] = v
        except (ValueError, TypeError):
            pass

    def _remplacer_id(val: Any) -> Any:
        if pd.isna(val) or val is None:
            return val
        if val in dict_map:
            return dict_map[val]
        val_str = str(val).strip()
        if val_str in dict_map:
            return dict_map[val_str]
        val_clean = re.sub(r"\.0$", "", val_str)
        if val_clean in dict_map:
            return dict_map[val_clean]
        try:
            n = float(val_str)
            if n.is_integer() and int(n) in dict_map:
                return dict_map[int(n)]
        except (ValueError, TypeError):
            pass
        return val

    df[player_col] = df[player_col].apply(_remplacer_id)
    return df


# ==============================================================================
# 5. CALCULS STAPS & INGÉNIERIE DES CARACTÉRISTIQUES (GROUPBY JOUEUR)
# ==============================================================================

def load_template_positions() -> Dict[str, str]:
    """Extrait les postes du classeur modèle officiel si présent."""
    pos_map: Dict[str, str] = {}
    template_path = get_template_path()
    if not template_path:
        return pos_map

    try:
        wb = openpyxl.load_workbook(template_path, data_only=True)
        if 'Liste joueurs' in wb.sheetnames:
            ws = wb['Liste joueurs']
            cur_grp: Optional[str] = None
            for row in ws.iter_rows(min_row=2, max_row=120):
                c0 = str(row[0].value or '').strip().lower()
                c1 = str(row[1].value or '').strip()
                if c0 and c0 != 'moyenne':
                    cur_grp = mapper_poste_vers_groupe_rca(c0)
                if c1 and c1.lower() != '1er mi temps ' and cur_grp:
                    pos_map[c1.lower()] = cur_grp
    except Exception:
        pass
    return pos_map


def assign_player_groups(df: pd.DataFrame, df_raw: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    """
    Assigne le groupe athlétique officiel ('5 de devant', '3ème Ligne', 'Trois-quarts') :
    1. En priorité absolue via le dictionnaire officiel 'joueurs_data' (effectif sénior RCA).
    2. Puis via les métadonnées brutes Catapult.
    3. Puis via le modèle Excel.
    4. Répartition de secours équilibrée.
    Conserve également le poste rugby précis dans la colonne 'Poste_Detaille'.
    """
    df = df.copy()
    player_col = "Nom ; Prénom"
    GROUPS = ["5 de devant", "3ème Ligne", "Trois-quarts"]
    template_positions = load_template_positions()
    unique_players = sorted(df[player_col].dropna().unique().tolist())

    group_list = []
    poste_detaille_list = []
    for idx, p in enumerate(df[player_col]):
        p_clean = str(p).strip()
        p_lower = p_clean.lower()

        # 1. Correspondance avec la base de données effectif RCA
        matched_poste = None
        for j_nom, j_poste in joueurs_data.items():
            if j_nom.lower() == p_lower or j_nom.lower() in p_lower or p_lower in j_nom.lower():
                matched_poste = j_poste
                break

        if matched_poste:
            grp = mapper_poste_vers_groupe_rca(matched_poste, nom_joueur=p_clean)
            poste_detaille_list.append(matched_poste)
            group_list.append(grp)
            continue

        # 2. Métadonnées brutes Catapult (si présentes dans df_raw)
        raw_poste = None
        if df_raw is not None and not df_raw.empty:
            p_raw_col = find_column(df_raw, ["Position Name", "Position", "Poste", "Pos"])
            name_raw_col = find_column(df_raw, ["Nom ; Prénom", "Player Name", "Athlete Name", "Player", "Athlete", "Joueur", "Nom", "Name"])
            if p_raw_col and name_raw_col and p_raw_col in df_raw.columns and name_raw_col in df_raw.columns:
                sub_r = df_raw[df_raw[name_raw_col].astype(str).str.strip().str.lower() == p_lower]
                if not sub_r.empty and pd.notna(sub_r[p_raw_col].iloc[0]):
                    raw_poste = str(sub_r[p_raw_col].iloc[0]).strip()

        if raw_poste:
            grp = mapper_poste_vers_groupe_rca(raw_poste, nom_joueur=p_clean)
            poste_detaille_list.append(raw_poste)
            group_list.append(grp)
            continue

        # 3. Modèle Excel
        if p_lower in template_positions:
            t_pos = template_positions[p_lower]
            grp = mapper_poste_vers_groupe_rca(t_pos, nom_joueur=p_clean)
            poste_detaille_list.append(t_pos)
            group_list.append(grp)
            continue

        matched_grp = None
        for t_name, t_grp in template_positions.items():
            if t_name in p_lower or p_lower in t_name:
                matched_grp = t_grp
                break
        if matched_grp:
            grp = mapper_poste_vers_groupe_rca(matched_grp, nom_joueur=p_clean)
            poste_detaille_list.append(matched_grp)
            group_list.append(grp)
            continue

        # 4. Repli déterministe
        p_idx = unique_players.index(p) if p in unique_players else idx
        fallback_grp = GROUPS[p_idx % 3]
        poste_detaille_list.append(fallback_grp)
        group_list.append(fallback_grp)

    df["Poste_Detaille"] = poste_detaille_list
    df["Groupe"] = group_list
    return df


def process_catapult_data(
    df_raw: pd.DataFrame,
    split_mode: str = "splits",
    split_filter: Optional[Union[str, List[str]]] = None,
    exclude_warmup: bool = False,
) -> pd.DataFrame:
    """
    Traite le DataFrame brut Catapult One selon les méthodes STAPS de la préparatrice physique RCA :
    - Par défaut ('splits') : Exclut la ligne globale 'all' et somme uniquement les ateliers réels (Split 1..N).
    - Mode 'all' : Retient la ligne globale 'all' (séance brute complète avec temps morts).
    - HSR (m) : Correspond strictement à la Zone 3 (Z3_km * 1000.0) pour jauges et graphiques de groupe.
    - VHSR (m) : Correspond strictement au cumul Zone 4 + Zone 5 ((Z4_km + Z5_km) * 1000.0).
    - Haute intensité Podium (m) : Cumul total Zone 3 + Zone 4 + Zone 5 ((Z3_km + Z4_km + Z5_km) * 1000.0).
    - ACC / DEC n° (> 3 m/s²) : Somme des 4 zones d'accélérations et décélérations intenses (> 3 m/s²).
    - Distance Per Min (m/min) : Moyenne des splits d'ateliers.
    - Convertit le Temps de jeu en minutes réelles.
    """
    player_col = find_column(df_raw, [
        "Nom ; Prénom", "Player Name", "Athlete Name", "Player", "Athlete", "Joueur", "Nom", "Name"
    ])
    if not player_col:
        raise ValueError("Colonne Joueur introuvable dans le fichier. Colonnes attendues : 'Player Name', 'Athlete Name', etc.")

    # 1. Gestion des lignes de session et des Splits (Méthode Staff RCA vs Ligne all)
    split_col = find_column(df_raw, [
        "Split Name", "Split", "Period Name", "Period", "Tags", "Tag", "Activity Name", "Session Name", "Session Title", "Segment"
    ])

    df_working = df_raw.copy()

    if split_col and split_col in df_working.columns:
        s_series = df_working[split_col].astype(str).str.strip()
        s_lower = s_series.str.lower()

        # A. Filtrage explicite si demandé (ex: multiselect d'ateliers spécifiques)
        if split_filter is not None:
            if isinstance(split_filter, str):
                mask = s_lower.str.contains(split_filter.strip().lower(), na=False)
                if mask.any():
                    df_working = df_working[mask].copy()
            elif isinstance(split_filter, (list, tuple, set)):
                targets = [str(x).strip().lower() for x in split_filter]
                mask = s_lower.isin(targets)
                if mask.any():
                    df_working = df_working[mask].copy()

        # B. Filtrage selon le mode choisi (Splits ateliers par défaut vs Ligne all)
        else:
            GLOBAL_TAGS = {
                "session", "whole session", "full session", "all", "total",
                "seance", "séance", "globale", "global", "complete", "complète",
                "session summary", "all session", "overall"
            }
            is_global = s_lower.isin(GLOBAL_TAGS) | s_lower.str.startswith("session ") | (s_lower == "session")
            has_global_rows = is_global.any()
            has_split_rows = (~is_global & (s_series != "") & (s_series.str.lower() != "nan")).any()

            norm_mode = str(split_mode).lower().strip()
            use_splits_aggregation = ("split" in norm_mode or "atelier" in norm_mode or norm_mode == "default") and has_split_rows

            if use_splits_aggregation and has_global_rows:
                # Mode Préparatrice Physique : Exclure la ligne globale "all" pour agréger uniquement les ateliers réels
                players_with_splits = set(df_working.loc[~is_global, player_col].astype(str).str.strip().unique())
                mask_safe = (~is_global) | (~df_working[player_col].astype(str).str.strip().isin(players_with_splits))
                df_working = df_working[mask_safe].copy()
            elif "all" in norm_mode and has_global_rows:
                # Mode Séance Brute Complète : Retenir la ligne globale "all"
                players_with_global = set(df_working.loc[is_global, player_col].astype(str).str.strip().unique())
                mask_safe = is_global | (~df_working[player_col].astype(str).str.strip().isin(players_with_global))
                df_working = df_working[mask_safe].copy()

        # C. Exclusion optionnelle de l'échauffement
        if exclude_warmup:
            warmup_mask = df_working[split_col].astype(str).str.lower().str.contains("echauff|échauff|warmup|warm-up|warm_up", na=False)
            if (~warmup_mask).any():
                df_working = df_working[~warmup_mask].copy()

    dur_col = find_column(df_working, [
        "Duration", "Temps de jeu", "Duration (mins)", "Duration (s)", "Duration (sec)", "Durée", "Time", "Session Duration"
    ])
    dist_col = find_column(df_working, ["Distance (km)", "Total Distance (km)", "Distance"])
    dist_m_col = find_column(df_working, ["Distance (m)", "Total Distance (m)"])

    hsr_col = find_column(df_working, [
        "Distance in Speed Zone 3 (km)", "Distance in Speed Zone 3  (km)",
        "Distance haute intensité (m)", "High Speed Running (m)", "HSR Distance (m)", "HSR Distance (>5m/s)"
    ])
    vhsr_col = find_column(df_working, [
        "Distance in Speed Zone 4 (km)", "Distance in Speed Zone 4  (km)",
        "Distance in Speed Zone 5 (km)", "Distance in Speed Zone 5  (km)",
        "Sprint Distance (m)", "VHSR Distance (m)", "Very High Speed Running (m)"
    ])

    top_speed_col = find_column(df_working, ["Top Speed (km/h)", "Max Velocity (km/h)", "Top Speed", "Max Velocity", "Vmax"])
    hr_max_col = find_column(df_working, ["Hr Max (bpm)", "Max Heart Rate (bpm)", "HR Max", "Max HR", "Heart Rate Max", "Hr Max"])
    dpm_col = find_column(df_working, ["Distance Per Min (m/min)", "Meterage Per Minute", "m/min", "Distance per min", "Dist/min"])

    max_accel_col = find_column(df_working, ["Max Acceleration (m/s/s)", "Max Acceleration (m/s2)", "Max Acceleration", "Max Accel"])
    max_decel_col = find_column(df_working, ["Max Deceleration (m/s/s)", "Max Deceleration (m/s2)", "Max Deceleration", "Max Decel"])

    z1_col = find_column(df_working, ["Distance in Speed Zone 1 (km)", "Distance in Speed Zone 1  (km)", "Speed Zone 1", "Zone 1"])
    z2_col = find_column(df_working, ["Distance in Speed Zone 2 (km)", "Distance in Speed Zone 2  (km)", "Speed Zone 2", "Zone 2"])
    z3_col = find_column(df_working, ["Distance in Speed Zone 3 (km)", "Distance in Speed Zone 3  (km)", "Speed Zone 3", "Zone 3"])
    z4_col = find_column(df_working, ["Distance in Speed Zone 4 (km)", "Distance in Speed Zone 4  (km)", "Speed Zone 4", "Zone 4"])
    z5_col = find_column(df_working, ["Distance in Speed Zone 5 (km)", "Distance in Speed Zone 5  (km)", "Speed Zone 5", "Zone 5"])

    acc_34_col = find_column(df_working, ["Accelerations Zone Count: 3 - 4 m/s/s", "Accelerations Zone Count: 3-4 m/s/s", "Accel 3-4", "Accelerations (>3m/s2)"])
    acc_gt4_col = find_column(df_working, ["Accelerations Zone Count: > 4 m/s/s", "Accelerations Zone Count: >4 m/s/s", "Accel >4"])
    dec_34_col = find_column(df_working, ["Deceleration Zone Count: 3 - 4 m/s/s", "Deceleration Zone Count: 3-4 m/s/s", "Decel 3-4", "Décélérations (<-3m/s2)"])
    dec_gt4_col = find_column(df_working, ["Deceleration Zone Count: > 4 m/s/s", "Deceleration Zone Count: >4 m/s/s", "Decel >4"])
    impacts_col = find_column(df_working, ["Impacts", "Total Impacts", "Player Load"])

    df_clean = pd.DataFrame()
    df_clean["Nom ; Prénom"] = df_working[player_col].astype(str).str.strip()

    if dur_col:
        df_clean["Duration_min"] = parse_duration_series(df_working[dur_col], col_name=dur_col)
    else:
        df_clean["Duration_min"] = 0.0

    if dist_col:
        vals = df_working[dist_col].apply(clean_numeric_value)
        if "(m)" in dist_col.lower() or vals.median() > 100.0:
            df_clean["Distance_km"] = vals / 1000.0
        else:
            df_clean["Distance_km"] = vals
    elif dist_m_col:
        df_clean["Distance_km"] = df_working[dist_m_col].apply(clean_numeric_value) / 1000.0
    else:
        df_clean["Distance_km"] = 0.0

    df_clean["Top_Speed"] = df_working[top_speed_col].apply(clean_numeric_value) if top_speed_col else 0.0
    df_clean["Hr_Max"] = df_working[hr_max_col].apply(clean_numeric_value) if hr_max_col else 0.0
    df_clean["Distance_Per_Min"] = df_working[dpm_col].apply(clean_numeric_value) if dpm_col else 0.0
    df_clean["Max_Accel"] = df_working[max_accel_col].apply(clean_numeric_value) if max_accel_col else 0.0
    df_clean["Max_Decel"] = df_working[max_decel_col].apply(lambda x: abs(clean_numeric_value(x))) if max_decel_col else 0.0
    df_clean["Impacts"] = df_working[impacts_col].apply(clean_numeric_value) if impacts_col else 0.0

    for z_name, c_source in [("Z1_km", z1_col), ("Z2_km", z2_col), ("Z3_km", z3_col), ("Z4_km", z4_col), ("Z5_km", z5_col)]:
        if c_source:
            vals = df_working[c_source].apply(clean_numeric_value)
            if "(m)" in c_source.lower() or vals.median() > 20.0:
                vals = vals / 1000.0
            df_clean[z_name] = vals
        else:
            df_clean[z_name] = 0.0

    # Si z3_col était absent mais hsr_col présent (ex: "HSR Distance (>5m/s)")
    if not z3_col and hsr_col:
        v_hsr = df_working[hsr_col].apply(clean_numeric_value)
        val_hsr_km = (v_hsr / 1000.0) if "(m)" in hsr_col.lower() or v_hsr.median() > 10.0 else v_hsr
        df_clean["Z3_km"] = val_hsr_km

    # Nomenclature officielle préparatrice physique RCA :
    # 1. HSR (m) correspond strictement à la Zone 3
    df_clean["HSR_m"] = df_clean["Z3_km"] * 1000.0

    # 2. VHSR (m) correspond strictement au cumul Zone 4 + Zone 5
    df_clean["Z4_m"] = df_clean["Z4_km"] * 1000.0
    df_clean["Z5_m"] = df_clean["Z5_km"] * 1000.0
    df_clean["VHSR_m"] = (df_clean["Z4_km"] + df_clean["Z5_km"]) * 1000.0

    # Si z4_col et z5_col étaient absents mais vhsr_col présent (ex: "Sprint Distance (m)")
    if not z4_col and not z5_col and vhsr_col:
        v_vhsr = df_working[vhsr_col].apply(clean_numeric_value)
        val_vhsr_m = v_vhsr if "(m)" in vhsr_col.lower() or v_vhsr.median() > 5.0 else v_vhsr * 1000.0
        df_clean["VHSR_m"] = val_vhsr_m
        df_clean["Z4_m"] = val_vhsr_m * 0.7
        df_clean["Z5_m"] = val_vhsr_m * 0.3
        df_clean["Z4_km"] = df_clean["Z4_m"] / 1000.0
        df_clean["Z5_km"] = df_clean["Z5_m"] / 1000.0

    df_clean["Acc_34"] = df_working[acc_34_col].apply(clean_numeric_value) if acc_34_col else 0.0
    df_clean["Acc_gt4"] = df_working[acc_gt4_col].apply(clean_numeric_value) if acc_gt4_col else 0.0
    df_clean["Dec_34"] = df_working[dec_34_col].apply(clean_numeric_value) if dec_34_col else 0.0
    df_clean["Dec_gt4"] = df_working[dec_gt4_col].apply(clean_numeric_value) if dec_gt4_col else 0.0

    # Groupby Joueur
    agg_dict = {
        "Duration_min": "sum",
        "Distance_km": "sum",
        "HSR_m": "sum",
        "VHSR_m": "sum",
        "Z4_m": "sum",
        "Z5_m": "sum",
        "Top_Speed": "max",
        "Hr_Max": "max",
        "Distance_Per_Min": "mean",
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
        "Impacts": "sum",
    }
    grouped = df_clean.groupby("Nom ; Prénom", as_index=False).agg(agg_dict)

    grouped["Duration_min"] = grouped["Duration_min"].astype(float)
    pos_dur = grouped["Duration_min"][grouped["Duration_min"] > 0]
    if len(pos_dur) > 0:
        med_dur = pos_dur.median()
        if med_dur > 100000:
            grouped["Duration_min"] = grouped["Duration_min"] / 60000.0
        elif med_dur > 240:
            grouped["Duration_min"] = grouped["Duration_min"] / 60.0

    temps_jeu = grouped["Duration_min"].astype(float)
    d_total_m = grouped["Distance_km"] * 1000.0
    dist_tot_km = grouped["Distance_km"]

    # Distance Per Min : moyenne des splits si disponible, sinon d_total_m / temps_jeu
    if grouped["Distance_Per_Min"].max() > 0:
        dist_per_min = grouped["Distance_Per_Min"].values
    else:
        dist_per_min = np.where(temps_jeu > 0, d_total_m / temps_jeu, 0.0)

    hsr_per_min = np.where(temps_jeu > 0, grouped["HSR_m"] / temps_jeu, 0.0)

    # ACC / DEC n° (> 3 m/s²) : somme exacte des 4 zones
    accel_decel_total = grouped["Acc_34"] + grouped["Acc_gt4"] + grouped["Dec_34"] + grouped["Dec_gt4"]

    # % de marche directement calculé sur les données agrégées selon le mode actif
    pct_marche = np.where(dist_tot_km > 0, (grouped["Z1_km"] / dist_tot_km) * 100.0, 0.0)
    pct_course = np.where(dist_tot_km > 0, (grouped["Z2_km"] / dist_tot_km) * 100.0, 0.0)
    pct_hsr = np.where(dist_tot_km > 0, (grouped["Z3_km"] / dist_tot_km) * 100.0, 0.0)
    pct_vhsr = np.where(dist_tot_km > 0, ((grouped["Z4_km"] + grouped["Z5_km"]) / dist_tot_km) * 100.0, 0.0)

    # Haute intensité Podium (m) : cumul total Zone 3 + Zone 4 + Zone 5 (> 14.4 km/h)
    haute_intensite_podium = (grouped["Z3_km"] + grouped["Z4_km"] + grouped["Z5_km"]) * 1000.0

    df_indiv = pd.DataFrame()
    df_indiv['Nom ; Prénom'] = grouped['Nom ; Prénom']
    df_indiv['Temps de jeu '] = temps_jeu.astype(float)
    df_indiv['Distance (km)'] = grouped['Distance_km']
    df_indiv[' Distance haute intensité (m)'] = grouped['HSR_m']
    df_indiv['Haute intensité Podium (m)'] = haute_intensite_podium
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
    df_indiv['VHSR (m)'] = grouped['VHSR_m']
    if hr_max_col:
        df_indiv['Hr Max (bpm)'] = grouped['Hr_Max']

    df_indiv = assign_player_groups(df_indiv, df_raw=df_working)
    return df_indiv


# ==============================================================================
# 6. MOTEUR DE CALCULS DYNAMIQUES (JAUGES ET PODIUM)
# ==============================================================================

def calculer_metriques_jauges(df_subset: pd.DataFrame, group_name: str, mode_calc: str = "Moyenne / joueur") -> Dict[str, Tuple[float, float]]:
    """Calcule dynamiquement les 4 métriques des jauges pour le groupe sélectionné."""
    targets = OBJECTIFS_CONFIG.get(group_name, OBJECTIFS_CONFIG["Tous les groupes"])

    if len(df_subset) == 0:
        return {k: (v, 0.0) for k, v in targets.items()}

    use_mean = (mode_calc == "Moyenne / joueur")

    d_m = df_subset["Distance (km)"] * 1000.0
    val_dist = d_m.mean() if use_mean else d_m.sum()

    # HSR = Zone 3 (selon méthodologie STAPS RCA)
    if "Distance in Speed Zone 3  (km)" in df_subset.columns:
        hsr_m = df_subset["Distance in Speed Zone 3  (km)"] * 1000.0
    elif " Distance haute intensité (m)" in df_subset.columns:
        hsr_m = df_subset[" Distance haute intensité (m)"]
    else:
        hsr_m = pd.Series([0.0] * len(df_subset))
    val_hsr = hsr_m.mean() if use_mean else hsr_m.sum()

    # VHSR = Zone 4 + Zone 5
    if "VHSR (m)" in df_subset.columns:
        vhsr_m = df_subset["VHSR (m)"]
    elif "Distance in Speed Zone 4  (km)" in df_subset.columns and "Distance in Speed Zone 5  (km)" in df_subset.columns:
        vhsr_m = (df_subset["Distance in Speed Zone 4  (km)"] + df_subset["Distance in Speed Zone 5  (km)"]) * 1000.0
    else:
        vhsr_m = pd.Series([0.0] * len(df_subset))
    val_vhsr = vhsr_m.mean() if use_mean else vhsr_m.sum()

    acc_dec = df_subset["Accel/Dec"] if "Accel/Dec" in df_subset.columns else pd.Series([0.0] * len(df_subset))
    val_acc_dec = acc_dec.mean() if use_mean else acc_dec.sum()

    nb_joueurs = len(df_subset)
    multiplier = 1.0 if use_mean else max(float(nb_joueurs), 1.0)

    return {
        "Objectif Distance (m)": (targets["Objectif Distance (m)"] * multiplier, val_dist),
        "Obj. HSR": (targets["Obj. HSR"] * multiplier, val_hsr),
        "Obj. VHSR": (targets["Obj. VHSR"] * multiplier, val_vhsr),
        "Obj. ACC/DEC": (targets["Obj. ACC/DEC"] * multiplier, val_acc_dec),
    }


def generer_podium_dynamique(df_subset: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
    """Génère le Podium des 3 tops à partir des données réelles avec photo Base64."""
    if len(df_subset) == 0:
        return {}

    podium_result: Dict[str, Dict[str, Any]] = {}

    def _make_player_entry(rank: int, name: str, stat_str: str) -> Dict[str, Any]:
        return {
            "rank": rank,
            "name": name,
            "stat": stat_str,
            "avatar_src": photo_to_base64_data_uri(name),
        }

    def _placeholder(rank: int) -> Dict[str, Any]:
        return {"rank": rank, "name": "-", "stat": "-", "avatar_src": photo_to_base64_data_uri("")}

    # 1. TOP SPEED
    ts_col = "Top Speed (km/h)"
    df_sorted_ts = df_subset.sort_values(by=ts_col, ascending=False).dropna(subset=[ts_col])
    players_ts = []
    for rank, (_, row) in enumerate(df_sorted_ts.head(3).iterrows(), start=1):
        name = str(row["Nom ; Prénom"]).strip()
        val = float(row[ts_col])
        players_ts.append(_make_player_entry(rank, name, f"{val:.1f} km/h"))
    while len(players_ts) < 3:
        players_ts.append(_placeholder(len(players_ts) + 1))

    podium_result["Top speed (km/h)"] = {
        "icon": "⚡",
        "players": players_ts,
    }

    # 2. % DE MARCHE
    walk_col = "%Marche "
    if walk_col in df_subset.columns:
        df_sorted_walk = df_subset.sort_values(by=walk_col, ascending=False)
    else:
        df_sorted_walk = df_subset
    players_walk = []
    for rank, (_, row) in enumerate(df_sorted_walk.head(3).iterrows(), start=1):
        name = str(row["Nom ; Prénom"]).strip()
        val_raw = float(row.get(walk_col, 0.0))
        val_pct = val_raw * 100.0 if val_raw <= 1.0 else val_raw
        players_walk.append(_make_player_entry(rank, name, f"{val_pct:.1f} %"))
    while len(players_walk) < 3:
        players_walk.append(_placeholder(len(players_walk) + 1))

    podium_result["% de marche"] = {
        "icon": "🚶",
        "players": players_walk,
    }

    # 3. HAUTE INTENSITÉ (m) - ALIGNEMENT SUR LA MÉTHODE STAFF RCA (Cumul Zones 3, 4 et 5 > 14.4 km/h)
    hi_col = "Haute intensité Podium (m)" if "Haute intensité Podium (m)" in df_subset.columns else " Distance haute intensité (m)"
    if hi_col in df_subset.columns:
        df_sorted_hi = df_subset.sort_values(by=hi_col, ascending=False)
    else:
        df_sorted_hi = df_subset
    players_hi = []
    for rank, (_, row) in enumerate(df_sorted_hi.head(3).iterrows(), start=1):
        name = str(row["Nom ; Prénom"]).strip()
        val_m = float(row.get(hi_col, 0.0))
        players_hi.append(_make_player_entry(rank, name, f"{int(round(val_m)):,d} m".replace(",", " ")))
    while len(players_hi) < 3:
        players_hi.append(_placeholder(len(players_hi) + 1))

    podium_result["Haute intensité (m)"] = {
        "icon": "🔥",
        "players": players_hi,
    }

    return podium_result


# ==============================================================================
# 7. INJECTION DANS LE TEMPLATE EXCEL VIA OPENPYXL (data_only=False)
# ==============================================================================

def populate_template_excel(df_indiv: pd.DataFrame, df_raw: Optional[pd.DataFrame] = None) -> io.BytesIO:
    """Charge le classeur officiel (Modele_Vierge_STAPS.xlsx) et injecte les données."""
    template_path = get_template_path()
    if not template_path:
        raise FileNotFoundError(f"Le fichier modèle '{TEMPLATE_FILENAME}' est introuvable.")

    wb = openpyxl.load_workbook(template_path, data_only=False)

    if 'données indiv. joueurs' in wb.sheetnames:
        ws_indiv = wb['données indiv. joueurs']
        num_players = len(df_indiv)

        for i, row in df_indiv.iterrows():
            r = 3 + i
            for col_idx, col_name in enumerate(FINAL_COLUMNS[:22], start=1):
                val = row[col_name]
                # Formatage des pourcentages pour le format natif Excel "0%"
                if col_name in ['%Marche ', '%course', '%HSR ', '%VHSR'] and isinstance(val, (int, float)) and val > 1.0:
                    val = val / 100.0
                ws_indiv.cell(row=r, column=col_idx, value=val)

            # Formules et métriques de la préparatrice physique RCA
            ws_indiv.cell(row=r, column=4, value=f"=SUM(I{r}:K{r})*1000")
            ws_indiv.cell(row=r, column=18, value=f"=SUM(N{r}:Q{r})")
            ws_indiv.cell(row=r, column=23, value=f"=SUM(S{r}:V{r})")
            ws_indiv.cell(row=r, column=24, value=row['HSR m/min'])
            ws_indiv.cell(row=r, column=26, value=f"=I{r}*1000")
            ws_indiv.cell(row=r, column=27, value=f"=(SUM(J{r}:K{r}))*1000")

        last_injected_row = 3 + num_players - 1
        if ws_indiv.max_row > last_injected_row:
            for extra_r in range(last_injected_row + 1, ws_indiv.max_row + 1):
                for c in range(1, 35):
                    ws_indiv.cell(row=extra_r, column=c).value = None

    if df_raw is not None and 'données brutes' in wb.sheetnames:
        ws_brutes = wb['données brutes']
        for col_idx, col_name in enumerate(df_raw.columns, start=1):
            ws_brutes.cell(row=3, column=col_idx, value=col_name)

        for row_idx, row_vals in enumerate(df_raw.itertuples(index=False), start=4):
            for col_idx, cell_val in enumerate(row_vals, start=1):
                val_to_write = cell_val if not pd.isna(cell_val) else None
                ws_brutes.cell(row=row_idx, column=col_idx, value=val_to_write)

        last_brute_row = 4 + len(df_raw) - 1
        if ws_brutes.max_row > last_brute_row:
            for extra_r in range(last_brute_row + 1, ws_brutes.max_row + 1):
                for c in range(1, len(df_raw.columns) + 1):
                    ws_brutes.cell(row=extra_r, column=c).value = None

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def create_fallback_excel(df_indiv: pd.DataFrame) -> io.BytesIO:
    """Génération Excel de secours autonome si le modèle est absent."""
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
# 7b. GÉNÉRATION DU RAPPORT PDF POUR LES COACHS (FPDF & MATPLOTLIB)
# ==============================================================================

JOURS_FR = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
MOIS_FR = [
    "", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
]


def formater_date_seance_fr(val_date: Any) -> str:
    """
    Convertit n'importe quel format de date (numéro de série Excel ex: 46273, string, timestamp)
    en date lisible en français avec numéro de semaine :
    ex: 'Mardi 8 Septembre 2026 - Semaine 37'
    """
    dt: Optional[datetime.datetime] = None
    if val_date is not None and not pd.isna(val_date) and str(val_date).strip() != "":
        # 1. Numéro de série Excel (ex: 46273 ou "46273")
        try:
            num_val = float(str(val_date).strip().replace(",", "."))
            if 30000 < num_val < 60000:
                dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=num_val)
        except (ValueError, TypeError):
            pass

        # 2. Parsing standard via pandas
        if dt is None:
            try:
                ts = pd.to_datetime(val_date, errors="coerce", dayfirst=True)
                if pd.notna(ts):
                    dt = ts.to_pydatetime()
            except Exception:
                pass

    if dt is None:
        dt = datetime.datetime.now()

    nom_jour = JOURS_FR[dt.weekday()]
    jour = dt.day
    nom_mois = MOIS_FR[dt.month]
    annee = dt.year
    semaine = dt.isocalendar()[1]
    return f"{nom_jour} {jour} {nom_mois} {annee} - Semaine {semaine}"


def generer_graphique_groupes_matplotlib(df_indiv: pd.DataFrame) -> io.BytesIO:
    """
    Génère un diagramme en barres comparant la Distance Moyenne (m)
    et le HSR Moyen (m) par Groupe de joueurs (5 de devant, 3ème Ligne, Trois-quarts).
    """
    fig, ax1 = plt.subplots(figsize=(6.0, 4.5), dpi=160)
    groupes = ["5 de devant", "3ème Ligne", "Trois-quarts"]

    dist_vals: List[float] = []
    hsr_vals: List[float] = []
    for g in groupes:
        sub = df_indiv[df_indiv["Groupe"] == g]
        if not sub.empty:
            dist_vals.append(float(sub["Distance (km)"].mean() * 1000.0))
            hsr_vals.append(float(sub[" Distance haute intensité (m)"].mean()))
        else:
            dist_vals.append(0.0)
            hsr_vals.append(0.0)

    x = np.arange(len(groupes))
    width = 0.35
    ax2 = ax1.twinx()

    rects1 = ax1.bar(x - width/2, dist_vals, width, label="Distance (m)", color="#1F4E79", edgecolor="none")
    rects2 = ax2.bar(x + width/2, hsr_vals, width, label="HSR (m)", color="#A30000", edgecolor="none")

    ax1.set_ylabel("Distance Moyenne (m)", color="#1F4E79", fontweight="bold", fontsize=8.5)
    ax2.set_ylabel("HSR Moyen (m)", color="#A30000", fontweight="bold", fontsize=8.5)
    ax1.set_xticks(x)
    ax1.set_xticklabels(groupes, fontweight="bold", fontsize=8.5)
    ax1.set_title("Distance Moyenne & HSR Moyen par Groupe", fontweight="bold", color="#1F4E79", fontsize=10.5, pad=12)

    # Valeurs écrites directement sur chaque barre
    for rect in rects1:
        h = rect.get_height()
        if h > 0:
            ax1.annotate(f"{int(round(h))} m",
                         xy=(rect.get_x() + rect.get_width() / 2, h),
                         xytext=(0, 3), textcoords="offset points",
                         ha="center", va="bottom", fontsize=8, fontweight="bold", color="#1F4E79")

    for rect in rects2:
        h = rect.get_height()
        if h > 0:
            ax2.annotate(f"{int(round(h))} m",
                         xy=(rect.get_x() + rect.get_width() / 2, h),
                         xytext=(0, 3), textcoords="offset points",
                         ha="center", va="bottom", fontsize=8, fontweight="bold", color="#A30000")

    max_dist = max(dist_vals) if dist_vals and max(dist_vals) > 0 else 1000.0
    max_hsr = max(hsr_vals) if hsr_vals and max(hsr_vals) > 0 else 100.0
    ax1.set_ylim(0, max_dist * 1.25)
    ax2.set_ylim(0, max_hsr * 1.25)
    ax1.grid(axis="y", linestyle="--", alpha=0.3)
    ax1.tick_params(axis="y", labelcolor="#1F4E79")
    ax2.tick_params(axis="y", labelcolor="#A30000")

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def generer_graphique_postes_matplotlib(df_indiv: pd.DataFrame) -> io.BytesIO:
    """
    Génère un diagramme en barres affichant la Vitesse Max (Vmax) Moyenne par Poste.
    """
    df_temp = df_indiv.copy()
    postes_list = []
    for _, row in df_temp.iterrows():
        nom = str(row.get("Nom ; Prénom", "")).strip()
        poste = str(row.get("Poste_Detaille", "")).strip()
        if not poste:
            poste = joueurs_data.get(nom, "")
        if not poste:
            for j_nom, j_pos in joueurs_data.items():
                if j_nom.lower() in nom.lower() or nom.lower() in j_nom.lower():
                    poste = j_pos
                    break
        if not poste:
            poste = str(row.get("Groupe", "Non défini"))
        postes_list.append(poste)
    df_temp["_Poste"] = postes_list

    grp_vmax = df_temp.groupby("_Poste")["Top Speed (km/h)"].mean().sort_values(ascending=True)

    fig, ax = plt.subplots(figsize=(6.0, 4.5), dpi=160)
    y_pos = np.arange(len(grp_vmax))
    rects = ax.barh(y_pos, grp_vmax.values, color="#2E5A88", height=0.55, edgecolor="none")

    ax.set_yticks(y_pos)
    ax.set_yticklabels(grp_vmax.index, fontweight="bold", fontsize=8, color="#334155")
    ax.set_xlabel("Vitesse Max (km/h)", fontweight="bold", fontsize=8.5, color="#1F4E79")
    ax.set_title("Vitesse Max Moyenne par Poste (km/h)", fontweight="bold", color="#1F4E79", fontsize=10.5, pad=12)

    # Valeurs écrites directement sur les barres
    for rect in rects:
        w = rect.get_width()
        if w > 0:
            ax.annotate(f"{w:.1f} km/h",
                        xy=(w, rect.get_y() + rect.get_height() / 2),
                        xytext=(4, 0), textcoords="offset points",
                        ha="left", va="center", fontsize=8, fontweight="bold", color="#1F4E79")

    max_v = grp_vmax.max() if not grp_vmax.empty and grp_vmax.max() > 0 else 30.0
    ax.set_xlim(0, max_v * 1.25)
    ax.grid(axis="x", linestyle="--", alpha=0.3)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


# ==============================================================================
# 7c. SYNTHÈSE NARRATIVE PAR LLM (NOTE DU PRÉPARATEUR PHYSIQUE)
# ==============================================================================

def generer_synthese_llm(df_indiv: pd.DataFrame, api_key: Optional[str] = None) -> str:
    """
    Rédige un court paragraphe de synthèse professionnelle (2 à 3 phrases maximum)
    destiné au staff technique du RCA en utilisant le client Gemini.
    Sécurité anti-panne intégrée : retourne un texte de repli propre et factuel
    si l'API key n'est pas configurée ou si l'appel échoue.
    """
    if df_indiv.empty:
        return "Séance enregistrée - Aucune donnée individuelle disponible pour l'analyse des charges."

    nb_joueurs = len(df_indiv)
    dist_tot = float(df_indiv["Distance (km)"].sum()) if "Distance (km)" in df_indiv.columns else 0.0
    dist_moy = float(df_indiv["Distance (km)"].mean()) if "Distance (km)" in df_indiv.columns else 0.0
    vmax = float(df_indiv["Top Speed (km/h)"].max()) if "Top Speed (km/h)" in df_indiv.columns else 0.0

    hsr_col = " Distance haute intensité (m)" if " Distance haute intensité (m)" in df_indiv.columns else ("HSR (m)" if "HSR (m)" in df_indiv.columns else None)
    hsr_tot = float(df_indiv[hsr_col].sum()) if hsr_col and not df_indiv[hsr_col].empty else 0.0
    hsr_moy = float(df_indiv[hsr_col].mean()) if hsr_col and not df_indiv[hsr_col].empty else 0.0
    if pd.isna(hsr_moy):
        hsr_moy = 0.0
    if pd.isna(hsr_tot):
        hsr_tot = 0.0

    # Texte de repli propre et factuel garanti
    fallback_texte = (
        f"Séance enregistrée pour {nb_joueurs} joueurs avec une charge globale de {dist_tot:.1f} km "
        f"({dist_moy:.2f} km/joueur) et un pic de vitesse mesuré à {vmax:.1f} km/h. "
        f"Le volume moyen à haute intensité s'établit à {int(round(hsr_moy))} m par joueur. "
        f"Analyse des charges validée par le pôle performance RCA."
    )

    try:
        resolved_key = (
            api_key
            or os.environ.get("GEMINI_API_KEY")
            or os.environ.get("GOOGLE_API_KEY")
            or (hasattr(st, "secrets") and ("GEMINI_API_KEY" in st.secrets) and st.secrets["GEMINI_API_KEY"])
            or (hasattr(st, "secrets") and ("GOOGLE_API_KEY" in st.secrets) and st.secrets["GOOGLE_API_KEY"])
            or (hasattr(st, "session_state") and st.session_state.get("gemini_api_key"))
        )

        if not resolved_key or not str(resolved_key).strip():
            return fallback_texte

        from google import genai
        client = genai.Client(api_key=str(resolved_key).strip())

        prompt = (
            "Agis en tant que préparateur physique de rugby senior du Rugby Club Amiénois (RCA). "
            "Rédige une synthèse athlétique factuelle, professionnelle et percutante (strictement 2 à 3 phrases maximum, "
            "sans puces, sans tirets, sans titre et sans aucun emoji) pour l'entraîneur principal en te basant sur ces métriques GPS :\n"
            f"- Effectif évalué : {nb_joueurs} joueurs\n"
            f"- Distance totale cumulée : {dist_tot:.1f} km (moyenne de {dist_moy:.2f} km/joueur)\n"
            f"- Vitesse maximale de la séance : {vmax:.1f} km/h\n"
            f"- Volume de haute intensité (HSR > 19.8 km/h) : moyenne de {int(round(hsr_moy))} m/joueur (total cumulé : {int(round(hsr_tot))} m)\n"
            "Style attendu : expert en sciences du sport, direct, axé sur la charge globale, la vitesse maximale et la sollicitation de haute intensité."
        )

        texte = None
        # 1. Tentative avec client.interactions (gemini-3.8-flash)
        try:
            interaction = client.interactions.create(
                model="gemini-3.8-flash",
                input=prompt,
            )
            texte = getattr(interaction, "output_text", None)
        except Exception:
            pass

        # 2. Repli vers client.models.generate_content (gemini-3.8-flash)
        if not texte:
            try:
                resp = client.models.generate_content(
                    model="gemini-3.8-flash",
                    contents=prompt,
                )
                texte = getattr(resp, "text", None)
            except Exception:
                pass

        # 3. Repli vers gemini-2.5-flash
        if not texte:
            try:
                resp = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=prompt,
                )
                texte = getattr(resp, "text", None)
            except Exception:
                pass

        if texte and str(texte).strip():
            clean_str = " ".join(str(texte).strip().split())
            return clean_str

        return fallback_texte

    except Exception:
        return fallback_texte


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
    for k, v in replacements.items():
        txt = txt.replace(k, v)
    clean_chars = []
    for ch in txt:
        try:
            ch.encode("latin-1")
            clean_chars.append(ch)
        except UnicodeEncodeError:
            clean_chars.append(" ")
    return "".join(clean_chars)


def generer_rapport_pdf_coachs(
    df_indiv: pd.DataFrame,
    date_seance: Optional[Any] = None,
    synthese_note: Optional[str] = None,
) -> bytes:
    """
    Génère un rapport PDF optimisé pour les coachs du RCA (format 1 page A4 Paysage) :
    - En-tête officiel RCA, date lisible & semaine
    - 5 Encarts KPIs globaux de la séance (dont Moyenne Haute Intensité / Joueur)
    - Encart narratif : Note du Préparateur Physique (Générée par LLM ou repli sécurisé)
    - Moitié gauche : Graphique Matplotlib de Distance Moyenne & HSR Moyen par Groupe
    - Moitié droite : Podium dynamique de la séance (Top Speed, % Marche, Haute Intensité)
    """
    date_texte = safe_latin1(formater_date_seance_fr(date_seance))

    class RCACoachPDF(FPDF):
        def header(self):
            self.set_fill_color(31, 78, 121)  # Bleu RCA (#1F4E79)
            self.rect(0, 0, 297, 16, style="F")

            self.set_xy(12, 3)
            self.set_font("Helvetica", "B", 13)
            self.set_text_color(255, 255, 255)
            self.cell(140, 10, "RUGBY CLUB AMIENOIS  |  Rapport de Seance GPS - RCA", align="L")

            self.set_xy(155, 3)
            self.set_font("Helvetica", "", 9.5)
            self.set_text_color(226, 232, 240)
            self.cell(130, 10, date_texte, align="R")

            self.set_text_color(30, 41, 59)
            self.set_y(20)

        def footer(self):
            self.set_y(-12)
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(148, 163, 184)
            self.cell(0, 8, "Document Staff Technique RCA - Page 1/1", align="C")

    pdf = RCACoachPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False)
    pdf.set_margins(left=12, top=20, right=12)

    # =========================================================================
    # PAGE UNIQUE A4 PAYSAGE : EN-TÊTE, KPIS, NOTE NARRATIVE, GRAPHIQUE & PODIUM
    # =========================================================================
    pdf.add_page()

    if df_indiv.empty:
        pdf.set_font("Helvetica", "", 12)
        pdf.cell(0, 10, "Aucune donnee de joueur disponible pour ce rapport.", align="L")
        return bytes(pdf.output())

    # 1. Encarts KPIs globaux de Séance
    nb_joueurs = len(df_indiv)
    dist_tot = float(df_indiv["Distance (km)"].sum()) if "Distance (km)" in df_indiv.columns else 0.0
    dist_moy = float(df_indiv["Distance (km)"].mean()) if "Distance (km)" in df_indiv.columns else 0.0
    vmax = float(df_indiv["Top Speed (km/h)"].max()) if "Top Speed (km/h)" in df_indiv.columns else 0.0
    hsr_col = "Haute intensité Podium (m)" if "Haute intensité Podium (m)" in df_indiv.columns else " Distance haute intensité (m)"
    hsr_moy = float(df_indiv[hsr_col].mean()) if hsr_col in df_indiv.columns and not df_indiv[hsr_col].empty else 0.0
    if pd.isna(hsr_moy):
        hsr_moy = 0.0

    kpis = [
        ("Joueurs Analyses", f"{nb_joueurs}"),
        ("Distance Cumulee", f"{dist_tot:.1f} km"),
        ("Moyenne / Joueur", f"{dist_moy:.2f} km"),
        ("Vitesse Max Seance", f"{vmax:.1f} km/h"),
        ("Haute Intensite / Joueur", f"{int(round(hsr_moy)):,d} m".replace(",", " ")),
    ]

    card_w = 273 / len(kpis)
    pdf.set_y(19.5)
    for title, val in kpis:
        x_curr = pdf.get_x()
        y_curr = pdf.get_y()
        pdf.set_fill_color(241, 245, 249)
        pdf.set_draw_color(203, 213, 225)
        pdf.rect(x_curr, y_curr, card_w - 2, 12, style="FD")

        pdf.set_xy(x_curr, y_curr + 1)
        pdf.set_font("Helvetica", "", 7.5)
        pdf.set_text_color(100, 116, 139)
        pdf.cell(card_w - 2, 3.5, safe_latin1(title), align="C")

        pdf.set_xy(x_curr, y_curr + 5.0)
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_text_color(31, 78, 121)
        pdf.cell(card_w - 2, 5.5, safe_latin1(val), align="C")

        pdf.set_xy(x_curr + card_w, y_curr)

    # 2. Encart narratif : Note du Préparateur Physique (LLM ou repli sécurisé)
    if synthese_note is None or not str(synthese_note).strip():
        synthese_note = generer_synthese_llm(df_indiv)

    y_note = 33.5
    h_note = 14.5
    w_note = 273.0

    # Fond de l'encadré
    pdf.set_fill_color(248, 250, 252)  # #F8FAFC
    pdf.set_draw_color(203, 213, 225)  # #CBD5E1
    pdf.rect(12, y_note, w_note, h_note, style="FD")

    # Bordure gauche accentuée Bleu RCA (#1F4E79)
    pdf.set_fill_color(31, 78, 121)
    pdf.rect(12, y_note, 3.5, h_note, style="F")

    # Titre du bloc narratif
    pdf.set_xy(18, y_note + 1.2)
    pdf.set_font("Helvetica", "B", 7.5)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(w_note - 8, 3.5, "NOTE DU PREPARATEUR PHYSIQUE  |  SYNTHESE ATHLETIQUE", align="L")

    # Corps du texte nettoyé pour compatibilité Latin-1
    pdf.set_xy(18, y_note + 4.8)
    pdf.set_font("Helvetica", "", 7.5)
    pdf.set_text_color(51, 65, 85)
    pdf.multi_cell(w=w_note - 9, h=3.6, txt=safe_latin1(synthese_note), align="L")

    # 3. Graphique Matplotlib : Distance & HSR par Groupe (Moitié gauche)
    buf1 = generer_graphique_groupes_matplotlib(df_indiv)
    chart_y = 50.5
    chart_w = 133
    chart_h = 139
    pdf.image(buf1, x=12, y=chart_y, w=chart_w, h=chart_h)

    # 4. Podium des Tops de la Séance (Moitié droite)
    podium_data = generer_podium_dynamique(df_indiv)

    x_pod = 151
    w_pod = 134

    # Bandeau titre de la section Podium
    pdf.set_fill_color(31, 78, 121)  # #1F4E79
    pdf.rect(x_pod, chart_y, w_pod, 6.5, style="F")
    pdf.set_xy(x_pod, chart_y)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(w_pod, 6.5, "PODIUM DE LA SEANCE", align="C")

    categories_config = [
        {
            "key": "Top speed (km/h)",
            "title": "TOP SPEED (km/h)",
            "subtitle": "Vitesse Maximale",
            "color": (31, 78, 121),  # #1F4E79
        },
        {
            "key": "% de marche",
            "title": "% DE MARCHE",
            "subtitle": "Basse Intensite",
            "color": (46, 90, 136),  # #2E5A88
        },
        {
            "key": "Haute intensité (m)",
            "title": "HAUTE INTENSITE (m)",
            "subtitle": "Cumul > 14.4 km/h",
            "color": (163, 0, 0),    # #A30000
        },
    ]

    rank_badges = {
        1: {"label": "1er", "bg": (217, 119, 6), "fg": (255, 255, 255)},    # Or (#D97706)
        2: {"label": "2eme", "bg": (100, 116, 139), "fg": (255, 255, 255)}, # Argent (#64748B)
        3: {"label": "3eme", "bg": (180, 83, 9), "fg": (255, 255, 255)},    # Bronze (#B45309)
    }

    y_card_start = chart_y + 6.5 + 2.0  # 59.0 mm
    card_h = 41.5
    card_gap = 2.5

    for i, cat_cfg in enumerate(categories_config):
        y_card = y_card_start + i * (card_h + card_gap)

        # Fond de la carte catégorie
        pdf.set_fill_color(248, 250, 252)  # #F8FAFC
        pdf.set_draw_color(203, 213, 225)  # #CBD5E1
        pdf.rect(x_pod, y_card, w_pod, card_h, style="FD")

        # En-tête de la carte catégorie
        accent_color = cat_cfg["color"]
        pdf.set_fill_color(*accent_color)
        pdf.rect(x_pod, y_card, w_pod, 6.0, style="F")

        # Titre catégorie (gauche)
        pdf.set_xy(x_pod + 4, y_card)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(70, 6.0, cat_cfg["title"], align="L")

        # Sous-titre catégorie (droite)
        pdf.set_xy(x_pod + 74, y_card)
        pdf.set_font("Helvetica", "I", 7.5)
        pdf.set_text_color(226, 232, 240)
        pdf.cell(56, 6.0, cat_cfg["subtitle"], align="R")

        # Récupération des joueurs de la catégorie
        cat_data = podium_data.get(cat_cfg["key"], {})
        players = cat_data.get("players", [])

        # Rendu des 3 joueurs (Top 3)
        row_w = w_pod - 4  # 130 mm
        row_h = 9.8
        row_gap = 1.2
        y_rows_start = y_card + 6.0 + 1.2

        for rank_idx in range(1, 4):
            y_row = y_rows_start + (rank_idx - 1) * (row_h + row_gap)
            x_row = x_pod + 2

            p_entry = players[rank_idx - 1] if len(players) >= rank_idx else {}
            p_name = safe_latin1(p_entry.get("name", "-"))
            p_stat = safe_latin1(p_entry.get("stat", "-"))

            # Conteneur de ligne
            pdf.set_fill_color(255, 255, 255)
            pdf.set_draw_color(226, 232, 240)
            pdf.rect(x_row, y_row, row_w, row_h, style="FD")

            # Style du badge rang
            badge_info = rank_badges.get(rank_idx, {"label": f"{rank_idx}e", "bg": (100, 116, 139), "fg": (255, 255, 255)})

            # Vérification de l'existence de la photo locale
            photo_path = obtenir_chemin_photo(p_name) if p_name and p_name != "-" else None
            has_photo = bool(photo_path and os.path.isfile(photo_path) and os.path.basename(photo_path) != "avatar_par_defaut.jpg")

            if has_photo:
                # 1. Badge rang
                pdf.set_fill_color(*badge_info["bg"])
                pdf.rect(x_row + 2, y_row + 1.2, 9.5, 7.4, style="F")
                pdf.set_xy(x_row + 2, y_row + 1.2)
                pdf.set_font("Helvetica", "B", 7)
                pdf.set_text_color(*badge_info["fg"])
                pdf.cell(9.5, 7.4, badge_info["label"], align="C")

                # 2. Photo joueur
                try:
                    pdf.image(photo_path, x=x_row + 13, y=y_row + 0.9, w=8.0, h=8.0)
                    pdf.set_draw_color(203, 213, 225)
                    pdf.rect(x_row + 13, y_row + 0.9, 8.0, 8.0, style="D")
                except Exception:
                    has_photo = False

                # 3. Nom joueur
                name_disp = p_name if len(p_name) <= 24 else p_name[:22] + ".."
                pdf.set_xy(x_row + 23.5, y_row)
                pdf.set_font("Helvetica", "B", 8)
                pdf.set_text_color(30, 41, 59)
                pdf.cell(66.5, row_h, name_disp, align="L")
            else:
                # 1. Badge rang propre
                pdf.set_fill_color(*badge_info["bg"])
                pdf.rect(x_row + 2, y_row + 1.2, 11.5, 7.4, style="F")
                pdf.set_xy(x_row + 2, y_row + 1.2)
                pdf.set_font("Helvetica", "B", 7.5)
                pdf.set_text_color(*badge_info["fg"])
                pdf.cell(11.5, 7.4, badge_info["label"], align="C")

                # 2. Nom joueur
                name_disp = p_name if len(p_name) <= 28 else p_name[:26] + ".."
                pdf.set_xy(x_row + 16, y_row)
                pdf.set_font("Helvetica", "B", 8.5)
                pdf.set_text_color(30, 41, 59)
                pdf.cell(74, row_h, name_disp, align="L")

            # Statistique en gras à droite
            pdf.set_xy(x_row + 90, y_row)
            pdf.set_font("Helvetica", "B", 9)
            pdf.set_text_color(*accent_color)
            pdf.cell(38, row_h, p_stat, align="R")

    return bytes(pdf.output())


# ==============================================================================
# 7d. COPILOTE INTELLIGENT RCA (INTERPRÉTATION EN LANGAGE NATUREL & FILTRAGE DYNAMIQUE)
# ==============================================================================

ALIAS_COLONNES: Dict[str, str] = {
    "distance": "Distance (km)",
    "distance (km)": "Distance (km)",
    "distance_km": "Distance (km)",
    "km": "Distance (km)",
    "vmax": "Top Speed (km/h)",
    "top speed": "Top Speed (km/h)",
    "top speed (km/h)": "Top Speed (km/h)",
    "top_speed": "Top Speed (km/h)",
    "vitesse": "Top Speed (km/h)",
    "vitesse max": "Top Speed (km/h)",
    "vitesse maximale": "Top Speed (km/h)",
    "hsr": " Distance haute intensité (m)",
    "hsr (m)": " Distance haute intensité (m)",
    "haute intensite": " Distance haute intensité (m)",
    "haute intensité": " Distance haute intensité (m)",
    "distance haute intensite": " Distance haute intensité (m)",
    "distance haute intensité": " Distance haute intensité (m)",
    "vhsr": "VHSR (m)",
    "vhsr (m)": "VHSR (m)",
    "tres haute intensite": "VHSR (m)",
    "très haute intensité": "VHSR (m)",
    "groupe": "Groupe",
    "poste": "Groupe",
    "nom": "Nom ; Prénom",
    "joueur": "Nom ; Prénom",
    "marche": "%Marche ",
    "%marche": "%Marche ",
    "course": "%course",
    "%course": "%course",
    "temps": "Temps de jeu ",
    "temps de jeu": "Temps de jeu ",
    "distance/min": "Distance Per Min (m/min)",
    "m/min": "Distance Per Min (m/min)",
    "haute intensite podium": "Haute intensité Podium (m)",
    "haute intensité podium": "Haute intensité Podium (m)",
    "podium haute intensite": "Haute intensité Podium (m)",
    "acc/dec": "Accel/Dec",
    "accel/dec": "Accel/Dec",
    "accel": "Accel/Dec",
    "accelerations": "Accel/Dec",
    "poste detaille": "Poste_Detaille",
    "poste détaillé": "Poste_Detaille",
    "poste_detaille": "Poste_Detaille",
    "position": "Poste_Detaille",
}


def resoudre_nom_colonne(col_name: str, available_cols: List[str]) -> Optional[str]:
    """Résout le nom de colonne exact en fonction des alias et des variations de casse/espaces."""
    if not col_name:
        return None
    c_clean = col_name.strip()
    if c_clean in available_cols:
        return c_clean
    c_lower = c_clean.lower()
    for col in available_cols:
        if col.strip().lower() == c_lower:
            return col
    if c_lower in ALIAS_COLONNES and ALIAS_COLONNES[c_lower] in available_cols:
        return ALIAS_COLONNES[c_lower]
    for k, target in ALIAS_COLONNES.items():
        if k in c_lower and target in available_cols:
            return target
    for col in available_cols:
        if c_lower in col.lower() or col.lower() in c_lower:
            return col
    return None


def parse_requete_locale(query: str, colonnes: List[str]) -> Dict[str, Any]:
    """
    Analyseur local déterministe (mots-clés et expressions régulières)
    activé si le LLM n'est pas configuré ou en cas de repli anti-panne.
    """
    q = query.lower().strip()
    filtres: List[Dict[str, Any]] = []
    desc_items: List[str] = []

    # 1. Réinitialisation
    if any(k in q for k in ["reset", "reinit", "tout", "tous", "annul", "effac", "aucun filtre", "complet"]):
        return {
            "filtres": [],
            "tri": None,
            "limite": None,
            "description_filtre": "Effectif complet (réinitialisation)"
        }

    # 2. Filtrage par groupe de joueurs
    if "5 de devant" in q or "cinq de devant" in q:
        filtres.append({"colonne": "Groupe", "operateur": "==", "valeur": "5 de devant"})
        desc_items.append("Groupe: 5 de devant")
    elif any(k in q for k in ["3eme", "3ème", "troisieme", "troisième"]):
        filtres.append({"colonne": "Groupe", "operateur": "==", "valeur": "3ème Ligne"})
        desc_items.append("Groupe: 3ème Ligne")
    elif any(k in q for k in ["avant", "pack", "gros"]):
        filtres.append({"colonne": "Groupe", "operateur": "in", "valeur": ["5 de devant", "3ème Ligne"]})
        desc_items.append("Avants (5 de devant + 3ème Ligne)")
    elif any(k in q for k in ["arriere", "arrière", "trois-quart", "trois quart", "3/4", "trois-quarts", "lignes arrieres"]):
        filtres.append({"colonne": "Groupe", "operateur": "==", "valeur": "Trois-quarts"})
        desc_items.append("Trois-quarts")

    # 3. Vitesse Max / Top Speed
    is_speed_query = any(w in q for w in ["speed", "vitesse", "vmax", "pointe"]) or ("km/h" in q and "distance" not in q)
    if is_speed_query:
        m_v_less = (
            re.search(r"(?:sous|inferieur\s+a|inferieure\s+a|<|en-dessous\s+de|moins\s+de)\s*(\d+(?:[.,]\d+)?)\s*(?:km/?h)?", q)
            or re.search(r"(?:vitesse|speed|vmax)\s*<\s*(\d+(?:[.,]\d+)?)", q)
        )
        if m_v_less:
            val = float(m_v_less.group(1).replace(",", "."))
            filtres.append({"colonne": "Top Speed (km/h)", "operateur": "<", "valeur": val})
            desc_items.append(f"Vitesse Max < {val:.1f} km/h")
        else:
            m_v_more = (
                re.search(r"(?:superieur\s+a|superieure\s+a|>|au-dessus\s+de|plus\s+de)\s*(\d+(?:[.,]\d+)?)\s*(?:km/?h)?", q)
                or re.search(r"(?:vitesse|speed|vmax)\s*>\s*(\d+(?:[.,]\d+)?)", q)
            )
            if m_v_more:
                val = float(m_v_more.group(1).replace(",", "."))
                filtres.append({"colonne": "Top Speed (km/h)", "operateur": ">", "valeur": val})
                desc_items.append(f"Vitesse Max > {val:.1f} km/h")

    # 4. Haute Intensité (HSR)
    is_hsr_query = any(w in q for w in ["hsr", "haute intensite", "haute intensité", "haute-intensite"])
    if is_hsr_query:
        m_hsr_val = (
            re.search(r"(?:plus\s+de|>|superieur\s+a|superieure\s+a|au-dessus\s+de)\s*(\d+(?:[.,]\d+)?)\s*(?:m|metres?)?\s*(?:en|de)?\s*(?:hsr|haute\s+intensit[ée])", q)
            or re.search(r"(?:hsr|haute\s+intensit[ée])\s*(?:>|plus\s+de|superieur\s+a|superieure\s+a)?\s*(\d+(?:[.,]\d+)?)\s*(?:m|metres?)?", q)
            or re.search(r"(\d+(?:[.,]\d+)?)\s*(?:m|metres?)\s*(?:en|de)?\s*(?:hsr|haute\s+intensit[ée])", q)
        )
        if m_hsr_val:
            val = float(m_hsr_val.group(1).replace(",", "."))
            if any(w in q for w in ["moins", "sous", "inferieur", "<"]):
                filtres.append({"colonne": " Distance haute intensité (m)", "operateur": "<", "valeur": val})
                desc_items.append(f"HSR < {int(val)} m")
            else:
                filtres.append({"colonne": " Distance haute intensité (m)", "operateur": ">", "valeur": val})
                desc_items.append(f"HSR > {int(val)} m")

    # 5. Très Haute Intensité (VHSR)
    is_vhsr_query = any(w in q for w in ["vhsr", "tres haute", "très haute"])
    if is_vhsr_query:
        m_vhsr = re.search(r"(\d+(?:[.,]\d+)?)\s*(?:m|metres?)?", q)
        if m_vhsr:
            val = float(m_vhsr.group(1).replace(",", "."))
            filtres.append({"colonne": "VHSR (m)", "operateur": ">", "valeur": val})
            desc_items.append(f"VHSR > {int(val)} m")

    # 6. Distance totale (km)
    is_dist_query = any(w in q for w in ["distance", "parcouru", "couru", "volume"]) and not is_hsr_query
    if is_dist_query:
        m_dist = re.search(r"(\d+(?:[.,]\d+)?)\s*(?:km)?", q)
        if m_dist:
            val = float(m_dist.group(1).replace(",", "."))
            if val < 20.0:
                if any(w in q for w in ["moins", "sous", "inferieur", "<"]):
                    filtres.append({"colonne": "Distance (km)", "operateur": "<", "valeur": val})
                    desc_items.append(f"Distance < {val:.1f} km")
                else:
                    filtres.append({"colonne": "Distance (km)", "operateur": ">", "valeur": val})
                    desc_items.append(f"Distance > {val:.1f} km")

    # 7. Tri / Limite
    limite = None
    m_top_n = re.search(r"top\s*(\d+)", q)
    if m_top_n:
        limite = int(m_top_n.group(1))

    tri = None
    if is_speed_query:
        tri = {"colonne": "Top Speed (km/h)", "ordre": "desc"}
    elif is_dist_query:
        tri = {"colonne": "Distance (km)", "ordre": "desc"}
    elif is_hsr_query:
        tri = {"colonne": " Distance haute intensité (m)", "ordre": "desc"}

    desc = " + ".join(desc_items) if desc_items else f"Filtre : '{query}'"
    return {"filtres": filtres, "tri": tri, "limite": limite, "description_filtre": desc}


SYSTEM_PROMPT_COPILOTE = (
    "Tu es le Copilote RCA, l'assistant data science officiel du Rugby Club Amiénois (RCA). "
    "Tu es expert en physiologie du sport, en préparation physique du rugby et en analyse des données GPS Catapult. "
    "Tu réponds exclusivement aux questions concernant l'effectif sénior du RCA, les charges d'entraînement et la performance athlétique. "
    "Utilise un ton technique, rigoureux et professionnel."
)


def interpreter_requete_copilote(
    query: str,
    df_indiv: Optional[pd.DataFrame] = None,
    colonnes_disponibles: Optional[List[str]] = None,
    api_key: Optional[str] = None
) -> Dict[str, Any]:
    """
    Traduit une requête en langage naturel (français) en structure de filtrage JSON
    pour une recherche ciblée indépendante en utilisant le LLM Gemini (ou analyseur local de repli).
    Injecte le System Prompt officiel du Copilote RCA et le contexte chiffré réel de la séance.
    """
    q_clean = query.strip()
    if not q_clean:
        return {"filtres": [], "tri": None, "limite": None, "description_filtre": "Requête vide.", "commentaire_expert": ""}

    q_low = q_clean.lower()
    if any(w in q_low for w in ["reset", "reinit", "tout afficher", "tous les joueurs", "aucun filtre", "fermer", "annuler"]):
        return {"filtres": [], "tri": None, "limite": None, "description_filtre": "Réinitialisation : ensemble de l'effectif.", "commentaire_expert": ""}

    cols = colonnes_disponibles or (list(df_indiv.columns) if df_indiv is not None else list(ALIAS_COLONNES.values()))

    # Injection du contexte réel chiffré de la séance
    contexte_chiffre = ""
    if df_indiv is not None and not df_indiv.empty:
        nb_joueurs = len(df_indiv)
        dist_tot = float(df_indiv["Distance (km)"].sum()) if "Distance (km)" in df_indiv.columns else 0.0
        dist_moy = float(df_indiv["Distance (km)"].mean()) if "Distance (km)" in df_indiv.columns else 0.0
        vmax = float(df_indiv["Top Speed (km/h)"].max()) if "Top Speed (km/h)" in df_indiv.columns else 0.0
        hsr_col = " Distance haute intensité (m)" if " Distance haute intensité (m)" in df_indiv.columns else ("HSR (m)" if "HSR (m)" in df_indiv.columns else None)
        hsr_moy = float(df_indiv[hsr_col].mean()) if hsr_col and not df_indiv[hsr_col].empty else 0.0
        hsr_tot = float(df_indiv[hsr_col].sum()) if hsr_col and not df_indiv[hsr_col].empty else 0.0

        contexte_chiffre = (
            f"Contexte réel de la séance GPS du RCA en cours d'analyse :\n"
            f"- Effectif total : {nb_joueurs} joueurs\n"
            f"- Distance moyenne : {dist_moy:.2f} km/joueur (distance cumulée : {dist_tot:.1f} km)\n"
            f"- Vmax max : {vmax:.1f} km/h\n"
            f"- Volume HSR : moyenne de {int(round(hsr_moy))} m/joueur (total cumulé : {int(round(hsr_tot))} m)\n"
        )

    # 1. Tentative avec le LLM Gemini avec System Instruction
    try:
        resolved_key = (
            api_key
            or os.environ.get("GEMINI_API_KEY")
            or os.environ.get("GOOGLE_API_KEY")
            or (hasattr(st, "secrets") and ("GEMINI_API_KEY" in st.secrets) and st.secrets["GEMINI_API_KEY"])
            or (hasattr(st, "secrets") and ("GOOGLE_API_KEY" in st.secrets) and st.secrets["GOOGLE_API_KEY"])
            or (hasattr(st, "session_state") and st.session_state.get("gemini_api_key"))
        )
        if resolved_key and str(resolved_key).strip():
            from google import genai
            from google.genai import types

            client = genai.Client(api_key=str(resolved_key).strip())

            prompt = f"""{contexte_chiffre}
Colonnes disponibles dans le tableau des joueurs (df_indiv) :
- 'Nom ; Prénom' : Nom complet du joueur
- 'Groupe' : '5 de devant', '3ème Ligne', 'Trois-quarts'
  * Contexte rugby RCA : 'avants' ou 'pack' = ['5 de devant', '3ème Ligne'] ; 'arrières' ou 'trois-quarts' = ['Trois-quarts']
- 'Distance (km)' : distance totale parcourue (km)
- 'Top Speed (km/h)' : vitesse maximale de pointe (km/h)
- ' Distance haute intensité (m)' : distance HSR (> 19.8 km/h en mètres)
- 'VHSR (m)' : distance très haute intensité (> 25.2 km/h en mètres)
- '%Marche ' : pourcentage du temps passé en basse intensité (0 à 100)
- 'Distance Per Min (m/min)' : intensité relative en mètres par minute

Consigne de recherche du staff technique : "{q_clean}"

Retourne UNIQUEMENT un objet JSON valide avec cette structure stricte :
{{
  "filtres": [
    {{
      "colonne": "nom_exact_de_la_colonne",
      "operateur": "> | >= | < | <= | == | != | in | contains",
      "valeur": nombre ou chaîne ou tableau
    }}
  ],
  "tri": {{
    "colonne": "nom_exact_de_la_colonne",
    "ordre": "asc | desc"
  }} ou null,
  "limite": nombre_entier ou null,
  "description_filtre": "Résumé technique et concis en français des critères appliqués",
  "commentaire_expert": "Court commentaire technique d'expert en sciences du sport (1 à 2 phrases) sur le profil athlétique ciblé par cette recherche"
}}
Si la requête demande de réinitialiser ou de fermer la recherche, retourne "filtres": [].
"""

            raw_json = None
            try:
                resp = client.models.generate_content(
                    model="gemini-3.8-flash",
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        system_instruction=SYSTEM_PROMPT_COPILOTE,
                        response_mime_type="application/json",
                        temperature=0.0
                    )
                )
                raw_json = resp.text
            except Exception:
                try:
                    resp = client.models.generate_content(
                        model="gemini-2.5-flash",
                        contents=prompt,
                        config=types.GenerateContentConfig(
                            system_instruction=SYSTEM_PROMPT_COPILOTE,
                            response_mime_type="application/json",
                            temperature=0.0
                        )
                    )
                    raw_json = resp.text
                except Exception:
                    pass

            if raw_json:
                cleaned_text = raw_json.strip()
                if cleaned_text.startswith("```"):
                    cleaned_text = re.sub(r"^```[a-zA-Z]*\n", "", cleaned_text)
                    cleaned_text = re.sub(r"\n```$", "", cleaned_text)
                data = json.loads(cleaned_text)
                if isinstance(data, dict) and "filtres" in data:
                    return data
    except Exception:
        pass

    # 2. Analyseur local de repli sécurisé (déterministe et autonome)
    res_locale = parse_requete_locale(q_clean, cols)
    res_locale["commentaire_expert"] = "Recherche ciblée traitée par le moteur analytique local du Copilote RCA."
    return res_locale


def appliquer_filtres_copilote(
    df: pd.DataFrame,
    config_filtre: Dict[str, Any]
) -> Tuple[pd.DataFrame, str]:
    """
    Applique de manière stricte et déterministe les critères du copilote sur le DataFrame.
    Retourne (df_filtre, description_lisible).
    """
    if df is None or df.empty:
        return df, "Aucune donnée disponible."

    df_out = df.copy()
    filtres = config_filtre.get("filtres", [])
    desc = config_filtre.get("description_filtre", "Filtre copilote appliqué")

    if not filtres:
        return df_out, "Effectif complet (aucun filtre actif)"

    cols_dispo = list(df_out.columns)

    for f in filtres:
        raw_col = f.get("colonne", "")
        op = f.get("operateur", "")
        val = f.get("valeur")

        col = resoudre_nom_colonne(raw_col, cols_dispo)
        if not col or col not in df_out.columns:
            continue

        try:
            if op in [">", ">=", "<", "<="]:
                val_num = float(val) if not isinstance(val, (int, float)) and str(val).replace(".", "", 1).replace("-", "", 1).isdigit() else val
                series_num = pd.to_numeric(df_out[col], errors="coerce")
                if op == ">":
                    df_out = df_out[series_num > val_num]
                elif op == ">=":
                    df_out = df_out[series_num >= val_num]
                elif op == "<":
                    df_out = df_out[series_num < val_num]
                elif op == "<=":
                    df_out = df_out[series_num <= val_num]
            elif op == "==":
                is_num = isinstance(val, (int, float)) or (isinstance(val, str) and str(val).replace(".", "", 1).replace("-", "", 1).isdigit())
                if is_num and pd.api.types.is_numeric_dtype(df_out[col]):
                    val_num = float(val)
                    series_num = pd.to_numeric(df_out[col], errors="coerce")
                    df_out = df_out[series_num == val_num]
                else:
                    df_out = df_out[df_out[col].astype(str).str.strip().str.lower() == str(val).strip().lower()]
            elif op == "!=":
                is_num = isinstance(val, (int, float)) or (isinstance(val, str) and str(val).replace(".", "", 1).replace("-", "", 1).isdigit())
                if is_num and pd.api.types.is_numeric_dtype(df_out[col]):
                    val_num = float(val)
                    series_num = pd.to_numeric(df_out[col], errors="coerce")
                    df_out = df_out[series_num != val_num]
                else:
                    df_out = df_out[df_out[col].astype(str).str.strip().str.lower() != str(val).strip().lower()]
            elif op == "in" and isinstance(val, list):
                val_clean = [str(x).strip().lower() for x in val]
                df_out = df_out[df_out[col].astype(str).str.strip().str.lower().isin(val_clean)]
            elif op == "contains" and isinstance(val, str):
                df_out = df_out[df_out[col].astype(str).str.contains(val, case=False, na=False)]
        except Exception:
            continue

    # Tri optionnel
    tri = config_filtre.get("tri")
    if tri and isinstance(tri, dict):
        col_tri = resoudre_nom_colonne(tri.get("colonne", ""), cols_dispo)
        if col_tri and col_tri in df_out.columns:
            ascending = str(tri.get("ordre", "desc")).lower() == "asc"
            df_out = df_out.sort_values(by=col_tri, ascending=ascending)

    # Limite optionnelle
    limite = config_filtre.get("limite")
    if limite and isinstance(limite, int) and limite > 0:
        df_out = df_out.head(limite)

    return df_out, desc


# ==============================================================================
# 8. APPLICATION STREAMLIT (INTERFACE UTILISATEUR COMPLÈTE)
# ==============================================================================

def main() -> None:
    st.set_page_config(
        page_title="Dashboard & Suivi des Charges GPS | Rugby Club Amiénois",
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
            margin-bottom: 1.4rem;
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
            margin-bottom: 1.1rem;
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
    st.markdown('<div class="main-title">🏉 Rugby Club Amiénois | Suivi des Charges GPS</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-title">Dashboard dynamique d\'analyse athlétique GPS & Injection automatisée dans le classeur de la préparatrice physique</div>',
        unsafe_allow_html=True,
    )

    # ==========================================================================
    # ÉTAPE 1 : UPLOAD DU FICHIER EN HAUT DE PAGE (.CSV OU .XLSX)
    # ==========================================================================
    uploaded_file = st.file_uploader(
        "📥 Déposer l'export brut Catapult One (.csv ou .xlsx) :",
        type=["csv", "xlsx", "xls"],
        help="Prend en charge les exports Catapult One, PlayerTek et OpenField au format CSV ou Excel avec détection automatique du vrai tableau.",
    )

    df_raw: Optional[pd.DataFrame] = None
    df_indiv: Optional[pd.DataFrame] = None

    if uploaded_file is None:
        col_msg, col_demo = st.columns([4, 1])
        with col_msg:
            st.info(
                "👋 **Bienvenue sur le portail d'analyse GPS du RCA.**\n\n"
                "Veuillez glisser-déposer votre export brut Catapult (`.csv` ou `.xlsx`) ci-dessus "
                "pour générer instantanément le rapport **Retour Journée**, le **Podium** et la synthèse de séance."
            )
        with col_demo:
            use_demo = st.checkbox("🧪 Mode Démo", value=False, help="Active des données de démonstration avec l'effectif réel RCA.")

        if use_demo:
            demo_rows = []
            for d in MOCK_DEMO_DATA:
                demo_rows.append({
                    "Date": "08/09/2026",
                    "Player Name": d["Player Name"],
                    "Position": d["Position"],
                    "Distance (km)": d["Distance (km)"],
                    "Distance in Speed Zone 4 (km)": d["Zone 4 (km)"],
                    "Distance in Speed Zone 5 (km)": d["Zone 5 (km)"],
                    "Top Speed (km/h)": d["Top Speed"],
                    "Duration": d["Duration"],
                    "Distance in Speed Zone 1 (km)": d["Z1 (km)"],
                    "Distance in Speed Zone 2 (km)": d["Z2 (km)"],
                    "Distance in Speed Zone 3 (km)": d["Z3 (km)"],
                    "Accelerations Zone Count: 3 - 4 m/s/s": d["Acc34"],
                    "Accelerations Zone Count: > 4 m/s/s": d["Acc4"],
                    "Deceleration Zone Count: 3 - 4 m/s/s": d["Dec34"],
                    "Deceleration Zone Count: > 4 m/s/s": d["Dec4"],
                })
            df_raw = pd.DataFrame(demo_rows)
            df_indiv = process_catapult_data(df_raw)
            st.toast("Mode démonstration actif avec les données de simulation RCA.")
        else:
            return

    else:
        try:
            with st.spinner("Lecture et analyse de l'export Catapult One..."):
                df_raw = load_catapult_file(uploaded_file)

            # ==================================================================
            # ÉTAPE 1b : MAPPING GPS DYNAMIQUE (BASÉ SUR LES BOÎTIERS RÉELS)
            # ==================================================================
            with st.expander("📡 Attribution des Boîtiers GPS aux Joueurs (Effectif RCA)", expanded=True):
                mapping_gps = ui_attribution_gps(df_raw=df_raw)
                if mapping_gps:
                    df_raw = appliquer_mapping(df_raw, mapping_gps)
                    int_keys = [k for k in mapping_gps if isinstance(k, int)]
                    joueurs_uniques = set(mapping_gps[k] for k in int_keys)
                    if len(joueurs_uniques) < len(int_keys):
                        st.warning(
                            f"⚠️ **Attention :** {len(int_keys)} boîtiers attribués pour {len(joueurs_uniques)} joueur(s) unique(s). "
                            "Des doublons existent et les distances des boîtiers d'un même joueur seront fusionnées."
                        )
                    else:
                        st.success(f"✅ Mapping appliqué : {len(int_keys)} boîtier(s) attribué(s) à l'effectif.")

            # ==================================================================
            # ÉTAPE 1c : GESTION DES SPLITS & SÉQUENCES CATAPULT
            # ==================================================================
            split_col_detected = find_column(df_raw, [
                "Split Name", "Split", "Period Name", "Period", "Tags", "Tag", "Activity Name", "Session Name", "Session Title", "Segment"
            ])
            selected_splits_val = None
            excl_warmup_val = False
            split_mode_val = "splits"

            if split_col_detected and split_col_detected in df_raw.columns:
                all_splits = [
                    s for s in df_raw[split_col_detected].dropna().astype(str).str.strip().unique().tolist()
                    if s and s.lower() != "nan"
                ]
                if len(all_splits) > 1:
                    with st.expander("⏱️ Gestion des Séquences & Splits de Séance (Catapult)", expanded=False):
                        st.info("💡 **Méthode Staff RCA (par défaut) :** Les volumes sont calculés en additionnant uniquement les séquences d'ateliers réelles (`Split 1..N`), excluant ainsi la marche des temps morts et transitions.")
                        c_sp1, c_sp2 = st.columns([3, 2])
                        with c_sp1:
                            mode_split = st.radio(
                                "Sélection des données à analyser :",
                                options=[
                                    "Somme des Splits (Méthode Staff RCA - sans temps morts)",
                                    "Ligne globale 'all' (Séance brute complète)",
                                    "Filtrer manuellement par Split",
                                ],
                                index=0,
                                key="radio_split_mode",
                                help="Basculer entre la méthode de la préparatrice physique (somme des ateliers réels) et la ligne globale 'all' (séance continue)."
                            )
                        with c_sp2:
                            excl_warmup_val = st.checkbox(
                                "Exclure l'échauffement",
                                value=False,
                                help="Exclut les séquences contenant 'échauffement' ou 'warmup'.",
                                key="chk_excl_warmup",
                            )
                        if mode_split == "Somme des Splits (Méthode Staff RCA - sans temps morts)":
                            split_mode_val = "splits"
                            selected_splits_val = None
                        elif mode_split == "Ligne globale 'all' (Séance brute complète)":
                            split_mode_val = "all"
                            selected_splits_val = None
                        elif mode_split == "Filtrer manuellement par Split":
                            split_mode_val = "manual"
                            selected_splits_val = st.multiselect(
                                "Choisir les splits à conserver :",
                                options=all_splits,
                                default=all_splits,
                                key="ms_splits_selected",
                            )

            with st.spinner("Traitement des données GPS..."):
                df_indiv = process_catapult_data(
                    df_raw,
                    split_mode=split_mode_val,
                    split_filter=selected_splits_val,
                    exclude_warmup=excl_warmup_val,
                )
            st.success(f"✅ Fichier chargé avec succès ! {len(df_indiv)} joueurs identifiés sur la séance.")
        except Exception as e:
            st.error(f"❌ Erreur lors du chargement du fichier Catapult : {str(e)}")
            with st.expander("Détails techniques de l'erreur"):
                st.exception(e)
            return

    # ==========================================================================
    # ÉTAPE 2 : COPILOTE RCA (ASSISTANT INTELLIGENT VOCAL & TEXTE) & FILTRES
    # ==========================================================================
    st.markdown("---")

    # Initialisation des états de session du Copilote
    if "copilote_active" not in st.session_state:
        st.session_state["copilote_active"] = False
    if "copilote_resultat" not in st.session_state:
        st.session_state["copilote_resultat"] = None
    if "copilote_desc" not in st.session_state:
        st.session_state["copilote_desc"] = ""
    if "copilote_commentaire" not in st.session_state:
        st.session_state["copilote_commentaire"] = ""
    if "copilote_last_voice" not in st.session_state:
        st.session_state["copilote_last_voice"] = ""

    # Conteneur d'interface du Copilote RCA (Recherche isolée experte)
    with st.container():
        st.markdown(
            """
            <div style="background-color: #F8FAFC; border: 1.5px solid #2E5A88; border-radius: 12px; padding: 14px 18px; margin-bottom: 12px;">
                <div style="display: flex; align-items: center; justify-content: space-between;">
                    <div style="font-weight: 800; font-size: 1.12rem; color: #1F4E79;">
                        🤖 Copilote RCA | Assistant Intelligent de Séance (Vocal & Clavier)
                    </div>
                </div>
                <div style="font-size: 0.88rem; color: #64748B; margin-top: 3px;">
                    Agent expert officiel du RCA : effectuez des recherches ciblées en langage naturel (au clavier ou au micro) en toute indépendance des données globales de la séance.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        col_q_text, col_q_btn, col_q_mic = st.columns([5.5, 1.5, 2.0])
        query_text = col_q_text.text_input(
            "Consigne pour le copilote :",
            placeholder="Ex: Affiche les avants avec plus de 150m en HSR, ou sous 20 km/h de top speed...",
            key="input_copilote_text",
            label_visibility="collapsed",
        )
        btn_apply_copilote = col_q_btn.button("🚀 Rechercher", key="btn_apply_copilote", use_container_width=True)

        voice_text = None
        if HAS_MIC_RECORDER:
            with col_q_mic:
                voice_text = speech_to_text(
                    language="fr",
                    start_prompt="🎙️ Dicter (Micro)",
                    stop_prompt="⏹️ Arrêter",
                    just_once=True,
                    key="copilote_voice_recorder",
                    use_container_width=True,
                )

        # Suggestions rapides d'analyse (5 boutons d'action rapide)
        st.caption("💡 **Suggestions rapides :**")
        chip_cols = st.columns(5)
        chip_query = None
        if chip_cols[0].button("⚡ Vitesse < 20 km/h", key="chip_v20", use_container_width=True):
            chip_query = "Affiche-moi uniquement les joueurs sous 20 km/h de top speed"
        if chip_cols[1].button("🏉 Avants HSR > 150m", key="chip_av_hsr", use_container_width=True):
            chip_query = "Filtre les avants qui ont fait plus de 150m en HSR"
        if chip_cols[2].button("🔥 Top Speed > 27 km/h", key="chip_v27", use_container_width=True):
            chip_query = "Joueurs avec vitesse de pointe supérieure à 27 km/h"
        if chip_cols[3].button("🏃 Distance > 5 km", key="chip_dist5", use_container_width=True):
            chip_query = "Joueurs ayant parcouru plus de 5 km de distance"
        if chip_cols[4].button("🛡️ 3ème Ligne", key="chip_3eme", use_container_width=True):
            chip_query = "Affiche uniquement les joueurs de 3ème Ligne"

        # Traitement de la requête active
        active_query = None
        if chip_query:
            active_query = chip_query
        elif btn_apply_copilote and query_text.strip():
            active_query = query_text.strip()
        elif voice_text and voice_text.strip() and voice_text != st.session_state.get("copilote_last_voice"):
            st.session_state["copilote_last_voice"] = voice_text
            active_query = voice_text.strip()
            st.toast(f"🎙️ Commande vocale détectée : « {active_query} »")

        if active_query:
            with st.spinner("Analyse et recherche ciblée par le Copilote RCA..."):
                cfg = interpreter_requete_copilote(
                    active_query,
                    df_indiv=df_indiv,
                    colonnes_disponibles=list(df_indiv.columns)
                )
                if not cfg.get("filtres"):
                    st.session_state["copilote_active"] = False
                    st.session_state["copilote_resultat"] = None
                    st.session_state["copilote_desc"] = ""
                    st.session_state["copilote_commentaire"] = ""
                    st.info("🔄 Recherche ciblée réinitialisée.")
                else:
                    df_copilote_resultat, desc = appliquer_filtres_copilote(df_indiv, cfg)
                    st.session_state["copilote_active"] = True
                    st.session_state["copilote_resultat"] = df_copilote_resultat
                    st.session_state["copilote_desc"] = desc
                    st.session_state["copilote_commentaire"] = cfg.get("commentaire_expert", "")
                    st.toast(f"🎯 Copilote : {len(df_copilote_resultat)} joueur(s) identifié(s)")

    # Affichage isolé et propre du résultat du Copilote (Dissociation stricte)
    if st.session_state.get("copilote_active") and st.session_state.get("copilote_resultat") is not None:
        df_copilote_resultat = st.session_state["copilote_resultat"]
        desc_res = st.session_state.get("copilote_desc", "Critères spécifiques")
        com_expert = st.session_state.get("copilote_commentaire", "")
        nb_trouves = len(df_copilote_resultat)
        nb_total = len(df_indiv)

        with st.container():
            st.markdown(
                f"""
                <div style="background-color: #FFFFFF; border: 1.5px solid #1F4E79; border-radius: 12px; padding: 16px 20px; margin-top: 8px; margin-bottom: 16px; box-shadow: 0 4px 14px rgba(31, 78, 121, 0.08);">
                    <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px;">
                        <div style="display: flex; align-items: center; gap: 8px;">
                            <span style="font-size: 1.25rem;">🎯</span>
                            <span style="font-size: 1.15rem; font-weight: 800; color: #1F4E79;">
                                Résultats de la recherche ciblée du Copilote
                            </span>
                        </div>
                        <span style="background-color: {'#EFF6FF' if nb_trouves > 0 else '#FEF2F2'}; color: {'#1D4ED8' if nb_trouves > 0 else '#B91C1C'}; border: 1px solid {'#BFDBFE' if nb_trouves > 0 else '#FECACA'}; padding: 4px 12px; border-radius: 20px; font-weight: 700; font-size: 0.9rem;">
                            {nb_trouves} joueur{'s' if nb_trouves > 1 else ''} correspondant{'s' if nb_trouves > 1 else ''} trouvé{'s' if nb_trouves > 1 else ''} sur {nb_total}
                        </span>
                    </div>
                    <div style="margin-top: 8px; font-size: 0.92rem; color: #334155;">
                        <strong>Critères appliqués :</strong> <span style="color: #0F172A;">{desc_res}</span>
                    </div>
                    {f'<div style="margin-top: 8px; font-size: 0.88rem; color: #475569; font-style: italic; background-color: #F8FAFC; padding: 8px 12px; border-radius: 6px; border-left: 3px solid #1F4E79;">🏉 <strong>Analyse Copilote RCA :</strong> {com_expert}</div>' if com_expert else ''}
                </div>
                """,
                unsafe_allow_html=True,
            )

            if df_copilote_resultat.empty:
                st.warning(f"⚠️ Aucun joueur de l'effectif RCA ne correspond aux critères '{desc_res}'.")
            else:
                df_cop_display = df_copilote_resultat.copy()
                num_cols = df_cop_display.select_dtypes(include=[np.number]).columns
                df_cop_display[num_cols] = df_cop_display[num_cols].round(2)
                cols_order_cop = ["Nom ; Prénom", "Groupe"] + (["Poste_Detaille"] if "Poste_Detaille" in df_cop_display.columns else []) + [c for c in FINAL_COLUMNS if c != "Nom ; Prénom" and c in df_cop_display.columns]
                df_cop_display = df_cop_display[[c for c in cols_order_cop if c in df_cop_display.columns]]
                st.dataframe(df_cop_display, use_container_width=True, hide_index=True)

            col_btn_close_cop, _ = st.columns([2.5, 4.5])
            with col_btn_close_cop:
                if st.button("❌ Fermer la recherche ciblée", key="btn_close_copilote_result", use_container_width=True):
                    st.session_state["copilote_active"] = False
                    st.session_state["copilote_resultat"] = None
                    st.session_state["copilote_desc"] = ""
                    st.session_state["copilote_commentaire"] = ""
                    st.rerun()

    # Filtre positionnel classique et mode de calcul (Dashboard global RCA)
    col_filter1, col_filter2 = st.columns([3, 2])
    with col_filter1:
        groupe_selectionne = st.selectbox(
            "🏉 Affiner par groupe positionnel :",
            options=["Tous les groupes", "5 de devant", "3ème Ligne", "Trois-quarts"],
            index=0,
            help="Ce filtre adapte l'ensemble des graphiques de charge, du podium et de la synthèse au groupe choisi.",
        )
    with col_filter2:
        mode_calcul = st.radio(
            "Mode de calcul des jauges :",
            options=["Moyenne / joueur", "Somme cumulée"],
            horizontal=True,
            help="'Moyenne / joueur' permet de comparer la performance moyenne du groupe à son objectif athlétique individuel.",
        )

    if groupe_selectionne == "Tous les groupes":
        df_groupe_actif = df_indiv
    else:
        df_groupe_actif = df_indiv[df_indiv["Groupe"] == groupe_selectionne]

    # ==========================================================================
    # ÉTAPE 3 : SECTION 1 : RETOUR JOURNÉE (JAUGES CONNECTÉES AUX VRAIES DONNÉES)
    # ==========================================================================
    st.markdown('<div class="section-header">📊 RETOUR JOURNÉE</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-desc">'
        'Rapport athlétique calculé en temps réel à partir des métriques GPS : '
        '<span class="legend-badge-obj">■ Objectif Séance</span> &nbsp;'
        '<span class="legend-badge-real">■ Réalisé GPS</span> '
        '&nbsp;(Les pourcentages d\'accomplissement sont intégrés dans les barres)'
        '</div>',
        unsafe_allow_html=True,
    )

    if groupe_selectionne == "Tous les groupes":
        tab_avants, tab_3eme, tab_arrieres = st.tabs([
            "🏉 5 de devant",
            "🏉 3ème Ligne",
            "🏉 Trois-quarts"
        ])
        with tab_avants:
            sub = df_indiv[df_indiv["Groupe"] == "5 de devant"]
            metrics = calculer_metriques_jauges(sub, "5 de devant", mode_calc=mode_calcul)
            render_dashboard_objectifs("5 de devant", metrics)
            st.caption(f"Effectif analysé : {len(sub)} joueurs dans les 5 de devant.")

        with tab_3eme:
            sub = df_indiv[df_indiv["Groupe"] == "3ème Ligne"]
            metrics = calculer_metriques_jauges(sub, "3ème Ligne", mode_calc=mode_calcul)
            render_dashboard_objectifs("3ème Ligne", metrics)
            st.caption(f"Effectif analysé : {len(sub)} joueurs en 3ème Ligne.")

        with tab_arrieres:
            sub = df_indiv[df_indiv["Groupe"] == "Trois-quarts"]
            metrics = calculer_metriques_jauges(sub, "Trois-quarts", mode_calc=mode_calcul)
            render_dashboard_objectifs("Trois-quarts", metrics)
            st.caption(f"Effectif analysé : {len(sub)} joueurs chez les Trois-quarts.")
    else:
        metrics = calculer_metriques_jauges(df_groupe_actif, groupe_selectionne, mode_calc=mode_calcul)
        render_dashboard_objectifs(groupe_selectionne, metrics)
        st.caption(f"Effectif analysé : {len(df_groupe_actif)} joueur(s) dans le groupe '{groupe_selectionne}'.")

    # ==========================================================================
    # ÉTAPE 4 : SECTION 2 : PODIUM (CLASSEMENT BASÉ SUR LES VRAIES DONNÉES)
    # ==========================================================================
    podium_dynamique = generer_podium_dynamique(df_groupe_actif)
    render_podium_section(podium_dynamique)

    st.markdown("---")

    # ==========================================================================
    # ÉTAPE 5 : SECTION 3 : SYNTHÈSE INDIVIDUELLE & EXPORT EXCEL OFFICIEL
    # ==========================================================================
    st.markdown('<div class="section-header">📁 Synthèse Complète de Séance & Export Excel</div>', unsafe_allow_html=True)

    template_path = get_template_path()
    if not template_path:
        st.warning(
            f"⚠️ **Fichier modèle absent :** `{TEMPLATE_FILENAME}` est introuvable dans le dossier. "
            "L'export s'effectuera en mode de secours sans les formules et onglets annexes du club."
        )

    kpi_cols = st.columns(4)
    with kpi_cols[0]:
        st.metric("Joueurs Analysés", len(df_groupe_actif))
    with kpi_cols[1]:
        st.metric("Distance Cumulée", f"{df_groupe_actif['Distance (km)'].sum():.2f} km")
    with kpi_cols[2]:
        st.metric("Vitesse Max Séance", f"{df_groupe_actif['Top Speed (km/h)'].max():.1f} km/h")
    with kpi_cols[3]:
        st.metric("Total Haute Intensité", f"{df_groupe_actif[' Distance haute intensité (m)'].sum():.0f} m")

    st.markdown("### 📋 Tableau de Bord Joueurs (`df_indiv`)")

    df_display = df_groupe_actif.copy()
    numeric_cols = df_display.select_dtypes(include=[np.number]).columns
    df_display[numeric_cols] = df_display[numeric_cols].round(2)

    cols_order = ["Nom ; Prénom", "Groupe"] + (["Poste_Detaille"] if "Poste_Detaille" in df_display.columns else []) + [c for c in FINAL_COLUMNS if c != "Nom ; Prénom" and c in df_display.columns]
    df_display = df_display[[c for c in cols_order if c in df_display.columns]]

    st.dataframe(df_display, use_container_width=True, hide_index=True)

    st.markdown("### 📥 Télécharger les Fichiers de Séance")
    col_excel, col_pdf = st.columns(2)

    with col_excel:
        st.markdown("#### 📊 Classeur Excel Complet")
        if template_path:
            excel_buffer = populate_template_excel(df_indiv=df_indiv, df_raw=df_raw)
            st.download_button(
                label=f"📥 Télécharger {OUTPUT_DOWNLOAD_FILENAME}",
                data=excel_buffer.getvalue(),
                file_name=OUTPUT_DOWNLOAD_FILENAME,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Télécharge le classeur officiel complet avec tous les onglets d'origine et les formules actualisées.",
                key="btn_download_excel",
            )
        else:
            fallback_buffer = create_fallback_excel(df_indiv=df_indiv)
            st.download_button(
                label=f"📥 Télécharger {OUTPUT_DOWNLOAD_FILENAME} (Format de secours)",
                data=fallback_buffer.getvalue(),
                file_name=OUTPUT_DOWNLOAD_FILENAME,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Fichier de secours généré sans les onglets modèles annexes.",
                key="btn_download_excel",
            )

    with col_pdf:
        st.markdown("#### 🏉 Rapport PDF pour les Coachs")
        # Récupération de la date de séance si disponible
        session_date_val = None
        if df_raw is not None:
            c_date = find_column(df_raw, ["Date", "Session Date", "Date de séance", "Date de seance"])
            if c_date and c_date in df_raw.columns:
                s_d = df_raw[c_date].dropna().astype(str).str.strip()
                if not s_d.empty:
                    session_date_val = s_d.iloc[0]

        # Gestion et mise en cache de la note narrative basée sur l'effectif global de la séance (df_indiv)
        cache_key = f"synthese_global_{len(df_indiv)}_{session_date_val}"
        if "synthese_cache_key" not in st.session_state or st.session_state["synthese_cache_key"] != cache_key:
            st.session_state["synthese_note_coachs"] = generer_synthese_llm(df_indiv)
            st.session_state["synthese_cache_key"] = cache_key

        synthese_note_active = st.session_state.get("synthese_note_coachs", "")

        with st.expander("📝 Note du Préparateur Physique (Génération IA / Synthèse)", expanded=True):
            st.caption("Synthèse narrative insérée automatiquement sur la Page 1 du rapport PDF pour le staff technique.")
            edited_note = st.text_area(
                "Contenu de la synthèse (modifiable avant export) :",
                value=synthese_note_active,
                height=90,
                key="text_area_synthese_note",
            )
            if edited_note != synthese_note_active:
                st.session_state["synthese_note_coachs"] = edited_note
                synthese_note_active = edited_note

            if st.button("🔄 Régénérer la synthèse par IA", key="btn_regen_synthese"):
                st.session_state["synthese_note_coachs"] = generer_synthese_llm(df_indiv)
                st.rerun()

        pdf_bytes = generer_rapport_pdf_coachs(
            df_indiv=df_indiv,
            date_seance=session_date_val,
            synthese_note=synthese_note_active,
        )
        clean_date_tag = re.sub(r"[^0-9a-zA-Z]", "_", str(session_date_val or datetime.date.today().strftime("%d_%m_%Y")))
        pdf_filename = f"Rapport_Seance_GPS_RCA_{clean_date_tag}.pdf"

        st.download_button(
            label="📄 Télécharger le Rapport PDF Coachs",
            data=pdf_bytes,
            file_name=pdf_filename,
            mime="application/pdf",
            help="Télécharge le rapport récapitulatif PDF imprimable format paysage (A4) pour le staff technique RCA.",
            key="btn_download_pdf",
        )


if __name__ == "__main__":
    main()
