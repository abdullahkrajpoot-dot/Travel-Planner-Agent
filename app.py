import base64
import json
import html
import os
import re
import tempfile
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path
import time

import requests
import streamlit as st
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from fpdf import FPDF
from PIL import Image, UnidentifiedImageError
from streamlit_quill import st_quill


st.set_page_config(page_title="Travel Planner", layout="wide")

st.markdown(
    """
    <style>
    /* Force dark color scheme and hide Streamlit theme menu to prevent light mode */
    :root, html, body { color-scheme: dark !important; background-color: #07111f !important; }
    .stApp { color-scheme: dark !important; }
    #MainMenu { display: none !important; }
    footer { display: none !important; }
    header[data-testid="stHeader"] { visibility: hidden !important; height: 0 !important; }
    [data-testid="stHeader"] { display: none !important; }
    [data-testid="stToolbar"] { display: none !important; }
    [data-testid="stDecoration"] { display: none !important; }
    [data-testid="stStatusWidget"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    [data-testid="stSidebarCollapseButton"] { display: none !important; }
    button[kind="header"] { display: none !important; }

    .stApp {
        background: linear-gradient(180deg, #07111f 0%, #0f172a 100%);
        color: #e2e8f0;
    }
    [data-testid="stAppViewContainer"] {
        background: linear-gradient(180deg, #07111f 0%, #0f172a 100%);
    }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0b1325 0%, #111827 100%);
        border-right: 1px solid rgba(148, 163, 184, 0.2);
    }
    [data-testid="stSidebar"] * {
        color: #dbeafe;
    }
    .stTextInput>div>div>input, .stDateInput input {
        background: rgba(15, 23, 42, 0.92) !important;
        color: #f8fafc !important;
        border: 1px solid rgba(148, 163, 184, 0.35) !important;
        border-radius: 10px !important;
    }
    .stTextInput>label, .stDateInput>label {
        color: #bfdbfe !important;
        font-weight: 600 !important;
    }
    .stAlert {
        background: rgba(30, 41, 59, 0.75) !important;
        border: 1px solid rgba(148, 163, 184, 0.25) !important;
        color: #f8fafc !important;
    }
    .main-header { color: #f8fafc; font-weight: 800; font-size: 2.5rem; margin-bottom: 0.15rem; }
    .sub-header { color: #cbd5e1; margin-bottom: 1rem; }
    .hero-card {
        background: linear-gradient(135deg, rgba(37,99,235,0.26), rgba(14,165,233,0.12));
        border: 1px solid rgba(148, 163, 184, 0.25);
        border-radius: 18px;
        padding: 1.2rem 1.3rem;
        margin-bottom: 1rem;
    }
    .summary-pill {
        display: inline-block;
        margin: 0.2rem 0.35rem 0 0;
        padding: 0.28rem 0.7rem;
        border-radius: 999px;
        background: rgba(30, 41, 59, 0.88);
        color: #e2e8f0;
        border: 1px solid rgba(148, 163, 184, 0.22);
        font-size: 0.88rem;
    }
    .day-card {
        background: rgba(15, 23, 42, 0.78);
        border: 1px solid rgba(148, 163, 184, 0.18);
        border-radius: 18px;
        padding: 1rem 1.1rem;
        margin-bottom: 1rem;
        color: #e5e7eb;
        box-shadow: 0 18px 40px rgba(0, 0, 0, 0.18);
    }
    .day-title { color: #f8fafc; font-weight: 700; font-size: 1.1rem; margin-bottom: 0.2rem; }
    .day-date { color: #93c5fd; font-weight: 700; margin-bottom: 0.7rem; }
    .section-title { color: #38bdf8; font-weight: 700; margin-top: 0.45rem; }
    .place-chip {
        display: inline-block;
        margin: 0.2rem 0.35rem 0 0;
        padding: 0.22rem 0.62rem;
        border-radius: 999px;
        background: rgba(37, 99, 235, 0.17);
        color: #dbeafe;
        border: 1px solid rgba(96, 165, 250, 0.25);
        font-size: 0.82rem;
    }
    .stButton>button, .stDownloadButton>button {
        background: #2563eb;
        color: white;
        border-radius: 10px;
        border: none;
        font-weight: 700;
        width: 100%;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def safe_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value or "")
    return normalized.encode("latin-1", "ignore").decode("latin-1").strip()


def get_openrouter_api_key() -> str:
    """Read the AI key from either the process environment or Streamlit secrets."""
    api_key = os.getenv("OPENROUTER_API_KEY", "").strip()
    if api_key:
        return api_key
    try:
        return str(st.secrets.get("OPENROUTER_API_KEY", "")).strip()
    except Exception:
        return ""


def html_text(value: str) -> str:
    return safe_text(value).replace("\n", "<br>")


def slugify(value: str) -> str:
    cleaned = re.sub(r"[^a-zA-Z0-9]+", "_", value or "")
    return cleaned.strip("_").lower() or "travel_plan"


def format_day_label(day_date: date) -> str:
    return day_date.strftime("%B %d - %A")


def compact_date(day_date: date) -> str:
    return day_date.strftime("%b %d, %Y")


def extract_json_block(raw_text: str) -> str:
    start = raw_text.find("{")
    end = raw_text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("JSON block not found")
    return raw_text[start : end + 1]


def normalize_place_list(value) -> list[str]:
    if not isinstance(value, list):
        return []
    places = []
    for item in value:
        text = safe_text(str(item))
        if text and text not in places:
            places.append(text)
    return places[:4]


def clean_day_title(title: str, idx: int) -> str:
    title = safe_text(title)
    title = re.sub(r"\bSignature\s+Highlights?\b", "Highlights", title, flags=re.IGNORECASE)
    title = re.sub(r"\s+", " ", title).strip(" -")
    return title or f"Day {idx + 1} Highlights"


def apply_trip_day_rules(days: list[dict], destination: str) -> list[dict]:
    if not days:
        return days

    for idx, day in enumerate(days):
        day["title"] = clean_day_title(day.get("title", ""), idx)

    days[0].update(
        {
            "title": "Arrival & Rest",
            "morning": f"Arrive in {destination}, complete transfer formalities, and check in comfortably.",
            "afternoon": "Take time to rest, unpack, and recover from the journey.",
            "evening": "Keep the evening relaxed with no visits or sightseeing planned.",
            "places": [],
        }
    )

    if len(days) > 1:
        days[-1].update(
            {
                "title": "Departure",
                "morning": "Pack, check out, and prepare for the return journey.",
                "afternoon": "Depart for the return trip with no visits or sightseeing planned.",
                "evening": "Continue the journey back home or onward as per schedule.",
                "places": [],
            }
        )

    return days


WIKI_HEADERS = {
    "User-Agent": "AI-Travel-Planner/1.0 (+https://github.com/abdullahkrajpoot-dot/Ai-Travel-Planner-Agent)",
    "Accept": "application/json",
}


BAD_PLACE_WORDS = {
    "seizure", "attack", "battle", "war", "incident", "massacre", "bombing",
    "riot", "siege", "uprising", "crash", "fire", "disaster", "politics",
    "election", "team", "club", "season", "fc", "football", "soccer", "hockey",
}

ANIMAL_PLACE_WORDS = {
    "animal", "animals", "bear", "bears", "bison", "buffalo", "camel", "camels",
    "deer", "dolphin", "dolphins", "eagle", "eagles", "elk", "fish", "fox",
    "giraffe", "goat", "goats", "horse", "horses", "kangaroo", "lion", "lions",
    "monkey", "monkeys", "penguin", "penguins", "sheep", "tiger", "tigers",
    "whale", "whales", "wildlife", "wolf", "wolves",
}

VISITABLE_PLACE_WORDS = {
    "aquarium", "beach", "bridge", "castle", "cave", "center", "centre", "church",
    "fort", "gallery", "garden", "heritage", "lake", "mall", "market", "memorial",
    "monument", "mosque", "mount", "mountain", "museum", "park", "palace", "reserve",
    "sanctuary", "square", "temple", "tower", "trail", "viewpoint", "zoo",
}


def is_valid_place_name(place_name: str) -> bool:
    tokens = set(slugify(place_name).split("_"))
    if not place_name or BAD_PLACE_WORDS.intersection(tokens):
        return False
    if ANIMAL_PLACE_WORDS.intersection(tokens) and not VISITABLE_PLACE_WORDS.intersection(tokens):
        return False
    return True


DESTINATION_LANDMARKS = {
    "makkah": [
        ["Masjid al-Haram, Makkah", "Kaaba, Makkah", "Maqam Ibrahim, Makkah", "Safa and Marwa, Makkah"],
        ["Jabal al-Nour, Makkah", "Hira Cave, Makkah", "Jannat al-Mu'alla, Makkah", "Masjid al-Jinn, Makkah"],
        ["Mount Arafat, Makkah", "Jabal al-Rahmah, Makkah", "Mina, Makkah", "Muzdalifah, Makkah"],
        ["Masjid Aisha, Makkah", "Abraj Al Bait, Makkah", "Clock Tower Museum, Makkah", "Makkah Museum"],
        ["Thawr Cave, Makkah", "Jabal Thawr, Makkah", "Al Kiswa Factory, Makkah", "Hudaibiyah, Makkah"],
        ["Masjid al-Khayf, Mina", "Jamaraat Bridge, Mina", "Bay'ah Mosque, Makkah", "Al Diyafa Mall, Makkah"],
    ],
    "mecca": [
        ["Masjid al-Haram, Makkah", "Kaaba, Makkah", "Maqam Ibrahim, Makkah", "Safa and Marwa, Makkah"],
        ["Jabal al-Nour, Makkah", "Hira Cave, Makkah", "Jannat al-Mu'alla, Makkah", "Masjid al-Jinn, Makkah"],
        ["Mount Arafat, Makkah", "Jabal al-Rahmah, Makkah", "Mina, Makkah", "Muzdalifah, Makkah"],
        ["Masjid Aisha, Makkah", "Abraj Al Bait, Makkah", "Clock Tower Museum, Makkah", "Makkah Museum"],
        ["Thawr Cave, Makkah", "Jabal Thawr, Makkah", "Al Kiswa Factory, Makkah", "Hudaibiyah, Makkah"],
        ["Masjid al-Khayf, Mina", "Jamaraat Bridge, Mina", "Bay'ah Mosque, Makkah", "Al Diyafa Mall, Makkah"],
    ],
    "madinah": [
        ["Al-Masjid an-Nabawi, Madinah", "Green Dome, Madinah", "Jannat al-Baqi, Madinah", "Quba Mosque, Madinah"],
        ["Mount Uhud, Madinah", "Uhud Martyrs Cemetery, Madinah", "The Seven Mosques, Madinah", "Masjid al-Qiblatayn, Madinah"],
        ["Dar Al Madinah Museum", "Hejaz Railway Museum, Madinah", "Quran Exhibition, Madinah", "Al Noor Mall, Madinah"],
    ],
    "medina": [
        ["Al-Masjid an-Nabawi, Madinah", "Green Dome, Madinah", "Jannat al-Baqi, Madinah", "Quba Mosque, Madinah"],
        ["Mount Uhud, Madinah", "Uhud Martyrs Cemetery, Madinah", "The Seven Mosques, Madinah", "Masjid al-Qiblatayn, Madinah"],
        ["Dar Al Madinah Museum", "Hejaz Railway Museum, Madinah", "Quran Exhibition, Madinah", "Al Noor Mall, Madinah"],
    ],
    "paris": [
        ["Eiffel Tower, Paris", "Louvre Museum, Paris", "Tuileries Garden, Paris", "Place de la Concorde, Paris"],
        ["Notre-Dame de Paris", "Sainte-Chapelle, Paris", "Latin Quarter, Paris", "Pantheon, Paris"],
        ["Arc de Triomphe, Paris", "Champs-Elysees, Paris", "Grand Palais, Paris", "Seine River, Paris"],
        ["Sacred Heart Basilica of Montmartre, Paris", "Montmartre, Paris", "Moulin Rouge, Paris", "Galeries Lafayette, Paris"],
    ],
    "london": [
        ["Tower Bridge, London", "Tower of London", "St Paul's Cathedral, London", "Borough Market, London"],
        ["Buckingham Palace, London", "Westminster Abbey, London", "Big Ben, London", "London Eye"],
        ["British Museum, London", "Covent Garden, London", "Trafalgar Square, London", "National Gallery, London"],
    ],
    "dubai": [
        ["Burj Khalifa, Dubai", "Dubai Mall", "Dubai Fountain", "Dubai Opera"],
        ["Museum of the Future, Dubai", "Dubai Frame", "Zabeel Park, Dubai", "Al Seef, Dubai"],
        ["Palm Jumeirah, Dubai", "Atlantis The Palm, Dubai", "Dubai Marina", "Jumeirah Beach Residence, Dubai"],
    ],
    "istanbul": [
        ["Hagia Sophia, Istanbul", "Blue Mosque, Istanbul", "Topkapi Palace, Istanbul", "Basilica Cistern, Istanbul"],
        ["Grand Bazaar, Istanbul", "Suleymaniye Mosque, Istanbul", "Galata Tower, Istanbul", "Karakoy, Istanbul"],
        ["Dolmabahce Palace, Istanbul", "Ortakoy Mosque, Istanbul", "Bosphorus Bridge, Istanbul", "Taksim Square, Istanbul"],
    ],
    "new york": [
        ["Statue of Liberty, New York", "Ellis Island, New York", "Battery Park, New York", "One World Trade Center, New York"],
        ["Times Square, New York", "Bryant Park, New York", "New York Public Library", "Grand Central Terminal, New York"],
        ["Central Park, New York", "Metropolitan Museum of Art", "Fifth Avenue, New York", "Rockefeller Center, New York"],
    ],
    "rome": [
        ["Colosseum, Rome", "Roman Forum, Rome", "Palatine Hill, Rome", "Capitoline Hill, Rome"],
        ["Trevi Fountain, Rome", "Pantheon, Rome", "Piazza Navona, Rome", "Spanish Steps, Rome"],
        ["St. Peter's Basilica, Vatican City", "Vatican Museums", "Castel Sant'Angelo, Rome", "Trastevere, Rome"],
    ],
    "tokyo": [
        ["Senso-ji, Tokyo", "Tokyo Skytree", "Ueno Park, Tokyo", "Ameya-Yokocho, Tokyo"],
        ["Meiji Shrine, Tokyo", "Shibuya Crossing, Tokyo", "Harajuku, Tokyo", "Omotesando, Tokyo"],
        ["Tokyo Tower", "Imperial Palace, Tokyo", "Ginza, Tokyo", "Hamarikyu Gardens, Tokyo"],
    ],
    "bangkok": [
        ["Grand Palace, Bangkok", "Wat Phra Kaew, Bangkok", "Wat Pho, Bangkok", "Wat Arun, Bangkok"],
        ["Jim Thompson House, Bangkok", "Bangkok Art and Culture Centre", "Siam Paragon, Bangkok", "Lumphini Park, Bangkok"],
        ["Chatuchak Weekend Market, Bangkok", "Asiatique The Riverfront, Bangkok", "Yaowarat Road, Bangkok", "Chao Phraya River, Bangkok"],
    ],
    "singapore": [
        ["Marina Bay Sands, Singapore", "Gardens by the Bay, Singapore", "Merlion Park, Singapore", "ArtScience Museum, Singapore"],
        ["Sentosa Island, Singapore", "Universal Studios Singapore", "S.E.A. Aquarium, Singapore", "Siloso Beach, Singapore"],
        ["Singapore Botanic Gardens", "Orchard Road, Singapore", "Chinatown, Singapore", "Clarke Quay, Singapore"],
    ],
    "barcelona": [
        ["Sagrada Familia, Barcelona", "Park Guell, Barcelona", "Casa Batllo, Barcelona", "Casa Mila, Barcelona"],
        ["Gothic Quarter, Barcelona", "Barcelona Cathedral", "La Rambla, Barcelona", "Boqueria Market, Barcelona"],
        ["Montjuic, Barcelona", "Magic Fountain of Montjuic", "Poble Espanyol, Barcelona", "Barceloneta Beach, Barcelona"],
    ],
    "kuala lumpur": [
        ["Petronas Twin Towers, Kuala Lumpur", "KLCC Park, Kuala Lumpur", "Aquaria KLCC, Kuala Lumpur", "Menara Kuala Lumpur"],
        ["Batu Caves, Selangor", "Merdeka Square, Kuala Lumpur", "Sultan Abdul Samad Building", "Central Market Kuala Lumpur"],
        ["Thean Hou Temple, Kuala Lumpur", "Perdana Botanical Gardens", "National Mosque of Malaysia", "Bukit Bintang, Kuala Lumpur"],
    ],
    "lahore": [
        ["Badshahi Mosque, Lahore", "Lahore Fort, Lahore", "Minar-e-Pakistan, Lahore", "Hazuri Bagh, Lahore"],
        ["Shalimar Gardens, Lahore", "Wazir Khan Mosque, Lahore", "Delhi Gate, Lahore", "Lahore Museum"],
        ["Tomb of Jahangir, Lahore", "Tomb of Nur Jahan, Lahore", "Akbari Sarai, Lahore", "Kamran's Baradari, Lahore"],
        ["Lahore Museum", "Anarkali Bazaar, Lahore", "Tollinton Market, Lahore", "Mall Road, Lahore"],
        ["Data Darbar, Lahore", "Masjid Wazir Khan, Lahore", "Shahi Hammam, Lahore", "Sunehri Mosque, Lahore"],
        ["Greater Iqbal Park, Lahore", "Lahore Zoo", "Bagh-e-Jinnah, Lahore", "Jilani Park, Lahore"],
        ["Emporium Mall, Lahore", "Packages Mall, Lahore", "Liberty Market, Lahore", "Fortress Stadium, Lahore"],
    ],
    "islamabad": [
        ["Faisal Mosque, Islamabad", "Pakistan Monument, Islamabad", "Lok Virsa Museum, Islamabad", "Daman-e-Koh, Islamabad"],
        ["Saidpur Village, Islamabad", "Margalla Hills National Park", "Rawal Lake, Islamabad", "Shakarparian, Islamabad"],
        ["Trail 5, Islamabad", "Pir Sohawa, Islamabad", "The Monal, Islamabad", "Japanese Park, Islamabad"],
        ["Centaurus Mall, Islamabad", "F-9 Park, Islamabad", "Pakistan National Council of Arts", "Rose and Jasmine Garden, Islamabad"],
    ],
    "karachi": [
        ["Mazar-e-Quaid, Karachi", "Mohatta Palace, Karachi", "Frere Hall, Karachi", "Clifton Beach, Karachi"],
        ["Empress Market, Karachi", "Port Grand, Karachi", "National Museum of Pakistan", "Dolmen Mall Clifton, Karachi"],
        ["Pakistan Maritime Museum, Karachi", "PAF Museum, Karachi", "Charna Island, Karachi", "Turtle Beach, Karachi"],
        ["Chaukhandi Tombs, Karachi", "Quaid-e-Azam House Museum", "Tooba Mosque, Karachi", "Burns Road, Karachi"],
    ],
    "pakistan": [
        ["Badshahi Mosque, Lahore", "Lahore Fort, Lahore", "Minar-e-Pakistan, Lahore", "Shalimar Gardens, Lahore"],
        ["Faisal Mosque, Islamabad", "Pakistan Monument, Islamabad", "Lok Virsa Museum, Islamabad", "Daman-e-Koh, Islamabad"],
        ["Mazar-e-Quaid, Karachi", "Mohatta Palace, Karachi", "Frere Hall, Karachi", "Clifton Beach, Karachi"],
        ["Altit Fort, Hunza", "Baltit Fort, Hunza", "Attabad Lake, Hunza", "Passu Cones, Hunza"],
        ["Derawar Fort, Bahawalpur", "Noor Mahal, Bahawalpur", "Lal Suhanra National Park, Bahawalpur", "Uch Sharif, Bahawalpur"],
    ],
    "belgium": [
        ["Grand Place, Brussels", "Galeries Royales Saint-Hubert, Brussels", "Mont des Arts, Brussels", "Atomium, Brussels"],
        ["Belfry of Bruges, Bruges", "Market Square, Bruges", "Canals of Bruges, Bruges", "Basilica of the Holy Blood, Bruges"],
        ["Gravensteen, Ghent", "Saint Bavo's Cathedral, Ghent", "Graslei, Ghent", "Korenlei, Ghent"],
        ["Antwerp Central Station, Antwerp", "Cathedral of Our Lady, Antwerp", "Grote Markt, Antwerp", "Museum aan de Stroom, Antwerp"],
    ],
}


def split_destination_names(destination: str) -> list[str]:
    cleaned = safe_text(destination)
    if not cleaned:
        return []
    parts = re.split(r"\s*(?:,|;|/|\+|&|\band\b|\bthen\b)\s*", cleaned, flags=re.IGNORECASE)
    unique = []
    for part in parts:
        item = part.strip()
        if item and item.lower() not in {value.lower() for value in unique}:
            unique.append(item)
    return unique or [cleaned]


def _landmarks_for_one_destination(destination: str) -> list[list[str]]:
    key = slugify(destination).replace("_", " ")
    for name, landmarks in DESTINATION_LANDMARKS.items():
        if name in key:
            return landmarks

    discovered = discover_destination_landmarks(destination)
    if discovered:
        return chunk_places(discovered)

    normalized_destination = safe_text(destination) or "the destination"
    return [[normalized_destination, normalized_destination, normalized_destination, normalized_destination]]


def get_destination_landmarks(destination: str) -> list[list[str]]:
    destinations = split_destination_names(destination)
    if len(destinations) <= 1:
        return _landmarks_for_one_destination(destination)

    combined_places = []
    for item in destinations[:4]:
        for group in _landmarks_for_one_destination(item):
            for place in group:
                if place not in combined_places:
                    combined_places.append(place)
    return chunk_places(combined_places)


def chunk_places(places: list[str], size: int = 4) -> list[list[str]]:
    chunks = [places[idx : idx + size] for idx in range(0, len(places), size)]
    if not chunks:
        return []
    while len(chunks[-1]) < size:
        chunks[-1].append(chunks[-1][len(chunks[-1]) % len(chunks[-1])])
    return chunks


def flat_unique_places(groups: list[list[str]]) -> list[str]:
    places = []
    for group in groups:
        for place in group:
            clean_place = safe_text(place)
            if clean_place and clean_place not in places and is_valid_place_name(clean_place):
                places.append(clean_place)
    return places


def get_landmark_groups_for_days(destination: str, total_days: int) -> list[list[str]]:
    groups = get_destination_landmarks(destination)
    places = flat_unique_places(groups)
    sightseeing_days = max(total_days - 2, 1)
    target_places = max(sightseeing_days * 4, 4)

    if len(places) < target_places:
        for item in split_destination_names(destination):
            for place in discover_destination_landmarks(item):
                if len(places) >= target_places:
                    break
                if place not in places and is_valid_place_name(place):
                    places.append(place)

    return chunk_places(places[:target_places])


def _wiki_api(params: dict, timeout: int = 20) -> dict:
    response = requests.get(
        "https://en.wikipedia.org/w/api.php",
        params={"format": "json", "utf8": 1, **params},
        timeout=timeout,
        headers=WIKI_HEADERS,
    )
    response.raise_for_status()
    return response.json()


def _wiki_search_titles(query: str, limit: int = 8) -> list[str]:
    try:
        payload = _wiki_api({
            "action": "query",
            "list": "search",
            "srsearch": query,
            "srlimit": limit,
        })
        titles = []
        for row in payload.get("query", {}).get("search", []):
            title = safe_text(row.get("title", ""))
            lowered = title.lower()
            if title and "disambiguation" not in lowered and not lowered.startswith("list of"):
                titles.append(title)
        return titles
    except Exception:
        return []


def _wiki_page_summary(title: str) -> dict | None:
    try:
        response = requests.get(
            f"https://en.wikipedia.org/api/rest_v1/page/summary/{requests.utils.quote(title.replace(' ', '_'))}",
            timeout=15,
            headers=WIKI_HEADERS,
        )
        if response.status_code != 200:
            return None
        return response.json()
    except Exception:
        return None


def _destination_coordinates(destination: str) -> tuple[float, float] | None:
    titles = [destination] + _wiki_search_titles(destination, limit=3)
    for title in titles:
        summary = _wiki_page_summary(title)
        coords = (summary or {}).get("coordinates") or {}
        lat = coords.get("lat")
        lon = coords.get("lon")
        if isinstance(lat, (int, float)) and isinstance(lon, (int, float)):
            return float(lat), float(lon)
    return None


def _nominatim_place(destination: str) -> dict | None:
    try:
        response = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={
                "q": destination,
                "format": "jsonv2",
                "limit": 1,
                "addressdetails": 1,
            },
            timeout=20,
            headers={"User-Agent": WIKI_HEADERS["User-Agent"]},
        )
        response.raise_for_status()
        results = response.json()
        if not results:
            return None
        return results[0]
    except Exception:
        return None


def _destination_coordinates_any(destination: str) -> tuple[float, float] | None:
    coords = _destination_coordinates(destination)
    if coords:
        return coords
    place = _nominatim_place(destination)
    if not place:
        return None
    try:
        return float(place["lat"]), float(place["lon"])
    except Exception:
        return None


def _osm_landmarks(destination: str, target_count: int = 60) -> list[str]:
    coords = _destination_coordinates_any(destination)
    if not coords:
        return []

    radius = 18000
    lat, lon = coords
    query = f"""
    [out:json][timeout:35];
    (
      node(around:{radius},{lat},{lon})["name"]["tourism"];
      way(around:{radius},{lat},{lon})["name"]["tourism"];
      relation(around:{radius},{lat},{lon})["name"]["tourism"];
      node(around:{radius},{lat},{lon})["name"]["historic"];
      way(around:{radius},{lat},{lon})["name"]["historic"];
      relation(around:{radius},{lat},{lon})["name"]["historic"];
      node(around:{radius},{lat},{lon})["name"]["amenity"~"place_of_worship|theatre|arts_centre|marketplace|library"];
      way(around:{radius},{lat},{lon})["name"]["amenity"~"place_of_worship|theatre|arts_centre|marketplace|library"];
      relation(around:{radius},{lat},{lon})["name"]["amenity"~"place_of_worship|theatre|arts_centre|marketplace|library"];
      node(around:{radius},{lat},{lon})["name"]["leisure"~"park|garden|nature_reserve"];
      way(around:{radius},{lat},{lon})["name"]["leisure"~"park|garden|nature_reserve"];
      relation(around:{radius},{lat},{lon})["name"]["leisure"~"park|garden|nature_reserve"];
      node(around:{radius},{lat},{lon})["name"]["building"~"mosque|church|cathedral|temple|synagogue|museum"];
      way(around:{radius},{lat},{lon})["name"]["building"~"mosque|church|cathedral|temple|synagogue|museum"];
      relation(around:{radius},{lat},{lon})["name"]["building"~"mosque|church|cathedral|temple|synagogue|museum"];
    );
    out tags center {target_count};
    """
    try:
        response = requests.post(
            "https://overpass-api.de/api/interpreter",
            data={"data": query},
            timeout=45,
            headers={"User-Agent": WIKI_HEADERS["User-Agent"]},
        )
        response.raise_for_status()
        elements = response.json().get("elements", [])
    except Exception:
        return []

    city_label = safe_text(destination)
    scored: list[tuple[int, str]] = []
    priority_tags = {
        "museum": 80,
        "attraction": 75,
        "viewpoint": 70,
        "artwork": 65,
        "gallery": 65,
        "castle": 80,
        "monument": 75,
        "memorial": 60,
        "archaeological_site": 70,
        "place_of_worship": 72,
        "park": 55,
        "garden": 55,
        "marketplace": 50,
        "library": 45,
    }
    banned_tokens = {*BAD_PLACE_WORDS, "hotel", "station", "airport", "school", "hospital", "clinic"}

    for element in elements:
        tags = element.get("tags", {}) or {}
        name = safe_text(tags.get("name", ""))
        if not name or not is_valid_place_name(name):
            continue
        tokens = set(slugify(name).split("_"))
        if banned_tokens.intersection(tokens):
            continue
        score = 10
        for tag_key in ("tourism", "historic", "amenity", "leisure", "building"):
            tag_value = safe_text(tags.get(tag_key, "")).lower()
            score += priority_tags.get(tag_value, 0)
        if city_label and city_label.lower() not in name.lower():
            name = f"{name}, {city_label}"
        scored.append((score, name))

    scored.sort(reverse=True, key=lambda item: item[0])
    unique = []
    seen = set()
    for _, name in scored:
        key = slugify(name)
        if key not in seen:
            unique.append(name)
            seen.add(key)
        if len(unique) >= target_count:
            break
    return unique


def _nearby_wiki_titles(destination: str) -> list[str]:
    coords = _destination_coordinates(destination)
    if not coords:
        return []
    try:
        payload = _wiki_api({
            "action": "query",
            "list": "geosearch",
            "gscoord": f"{coords[0]}|{coords[1]}",
            "gsradius": 10000,
            "gslimit": 35,
        })
        banned = {
            "station", "metro", "subway", "railway", "airport", "hotel", "school",
            *BAD_PLACE_WORDS,
        }
        titles = []
        for row in payload.get("query", {}).get("geosearch", []):
            title = safe_text(row.get("title", ""))
            lowered = title.lower()
            if title and not any(word in lowered for word in banned):
                titles.append(title)
        return titles
    except Exception:
        return []


@st.cache_data(ttl=60 * 60 * 12, show_spinner=False)
def discover_destination_landmarks(destination: str) -> list[str]:
    normalized_destination = safe_text(destination)
    if not normalized_destination:
        return []

    candidates: list[str] = []
    candidates.extend(_osm_landmarks(normalized_destination, target_count=80))
    search_queries = [
        f"{normalized_destination} landmarks tourist attractions",
        f"{normalized_destination} famous places",
        f"{normalized_destination} museums historic sites",
    ]
    for query in search_queries:
        candidates.extend(_wiki_search_titles(query, limit=15))
    candidates.extend(_nearby_wiki_titles(normalized_destination))

    unique = []
    seen = set()
    destination_tokens = {token for token in slugify(normalized_destination).split("_") if len(token) > 2}
    attraction_words = {
        "museum", "palace", "castle", "cathedral", "mosque", "church", "temple", "fort",
        "tower", "square", "park", "garden", "monument", "bridge", "gallery", "market",
        "old", "historic", "gate", "wall", "harbour", "waterfront", "viewpoint",
        "shrine", "basilica", "beach", "lake", "cave", "mount", "mountain", "mall",
        "zoo", "aquarium", "library", "theatre", "opera", "citadel", "fortress",
    }
    banned_words = {*BAD_PLACE_WORDS, "people", "list"}

    for title in candidates:
        clean_title = safe_text(title)
        key = slugify(clean_title)
        if not clean_title or key in seen:
            continue
        title_tokens = set(key.split("_"))
        if banned_words.intersection(title_tokens) or not is_valid_place_name(clean_title):
            continue
        has_context = bool(destination_tokens.intersection(title_tokens))
        has_attraction_word = bool(attraction_words.intersection(title_tokens))
        if not has_context and not has_attraction_word:
            continue
        seen.add(key)
        if normalized_destination.lower() not in clean_title.lower():
            unique.append(f"{clean_title}, {normalized_destination}")
        else:
            unique.append(clean_title)
        if len(unique) >= 80:
            break

    return unique


def fallback_plan(destination: str, start_date: date, end_date: date):
    total_days = max((end_date - start_date).days + 1, 1)
    templates = get_landmark_groups_for_days(destination, total_days)
    days = []
    for idx in range(total_days):
        current_date = start_date + timedelta(days=idx)
        places = templates[idx % len(templates)]
        if idx == 0:
            title = "Arrival & Rest"
        elif idx == total_days - 1:
            title = "Departure"
        else:
            title = f"Day {idx + 1} Highlights"
        days.append(
            {
                "date": current_date,
                "title": title,
                "morning": f"Begin with {places[0]} and {places[1]}, two well-known highlights connected to {destination}.",
                "afternoon": f"Continue toward {places[2]} for local culture, photos, lunch, and neighborhood exploration.",
                "evening": f"End the day around {places[3]} with sunset views, dinner, and a comfortable transfer back to the hotel.",
                "places": places,
            }
        )
    return apply_trip_day_rules(days, destination)


def parse_ai_plan(raw_text: str, destination: str, start_date: date, end_date: date):
    total_days = max((end_date - start_date).days + 1, 1)
    expected_dates = [start_date + timedelta(days=i) for i in range(total_days)]
    landmark_templates = get_landmark_groups_for_days(destination, total_days)
    payload = json.loads(extract_json_block(raw_text))
    raw_days = payload.get("days", [])
    if not isinstance(raw_days, list) or not raw_days:
        return fallback_plan(destination, start_date, end_date)

    normalized = []
    for idx in range(total_days):
        item = raw_days[idx] if idx < len(raw_days) and isinstance(raw_days[idx], dict) else {}
        normalized.append(
            {
                "date": expected_dates[idx],
                "title": clean_day_title(item.get("title") or f"Day {idx + 1} Highlights", idx),
                "morning": safe_text(item.get("morning", "")),
                "afternoon": safe_text(item.get("afternoon", "")),
                "evening": safe_text(item.get("evening", "")),
                "places": normalize_place_list(item.get("places", [])),
            }
        )

    used_plan_places: set[str] = set()
    for idx, day in enumerate(normalized):
        fresh_places = []
        for place in day["places"]:
            if is_valid_place_name(place) and place not in used_plan_places:
                fresh_places.append(place)
                used_plan_places.add(place)
        day["places"] = fresh_places
        fallback_places = landmark_templates[idx % len(landmark_templates)] if landmark_templates else []
        for place in fallback_places:
            if len(day["places"]) >= 4:
                break
            if place not in used_plan_places and is_valid_place_name(place):
                day["places"].append(place)
                used_plan_places.add(place)

    return apply_trip_day_rules(normalized, destination)


def get_ai_plan(client: str, destination: str, start_date: date, end_date: date):
    api_key = get_openrouter_api_key()
    if not api_key:
        return fallback_plan(destination, start_date, end_date)

    total_days = max((end_date - start_date).days + 1, 1)
    prompt = f"""
Create a professional travel itinerary in strict JSON.

Client: {client}
Destination: {destination}
Trip length: {total_days} days
Dates: {start_date} to {end_date}

Rules:
- Return JSON only.
- Root object must be {{ "days": [ ... ] }}.
- Include exactly {total_days} day objects.
- First day must be arrival and rest only, with no sightseeing visits.
- Last day must be departure/return only, with no sightseeing visits.
- First and last day "places" must be an empty array.
- Each day object must contain:
  - "title"
  - "morning"
  - "afternoon"
  - "evening"
  - "places"
- For middle sightseeing days, "places" must be an array of 3 or 4 real, specific place names strongly related to that day's activities.
- Place names should be exact enough for image search, like "Badshahi Mosque, Lahore" not generic terms like "city center".
- Do not repeat the same place on another day unless the trip is longer than the destination's available attractions.
- Each day's morning, afternoon, and evening text must describe that same day's places.
- Use visitor attractions only: mosques, museums, mountains, caves, parks, markets, towers, monuments, heritage sites, viewpoints, and malls.
- Every item in "places" must be a physical, visitable attraction or landmark. Never use an animal, person, species, organization, event, or general cultural topic as a place.
- Wildlife is allowed only when the place is a named zoo, aquarium, park, reserve, or sanctuary—not an animal name such as "Big Horn Sheep".
- Do not include events, incidents, attacks, battles, seizures, wars, disasters, hotels, airports, or railway stations as places.
- Keep the itinerary realistic and tourist-friendly.
- Do not use the phrase "Signature Highlight" or "Signature Highlights".
"""

    try:
        response = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={"Authorization": f"Bearer {api_key}"},
            json={
                "model": os.getenv("OPENROUTER_MODEL", "openrouter/auto"),
                "messages": [
                    {"role": "system", "content": "You are a precise travel-plan editor. Follow every user requirement and return valid JSON only."},
                    {"role": "user", "content": prompt},
                ],
                "temperature": 0.15,
            },
            timeout=90,
        )
        response.raise_for_status()
        data = response.json()
        raw_text = data["choices"][0]["message"]["content"]
        return parse_ai_plan(raw_text, destination, start_date, end_date)
    except Exception:
        return fallback_plan(destination, start_date, end_date)


def revise_plan_with_prompt(
    days: list[dict],
    instruction: str,
    destination: str,
    reference_image: dict | None = None,
) -> list[dict]:
    instruction = (instruction or "").strip()
    if not instruction:
        return days
    api_key = get_openrouter_api_key()
    if not api_key:
        updated = [dict(day) for day in days]
        ordinal_map = {
            "first": 0, "1st": 0, "pehla": 0, "pehli": 0,
            "second": 1, "2nd": 1, "dusra": 1, "doosra": 1,
            "third": 2, "3rd": 2, "teesra": 2,
            "fourth": 3, "4th": 3, "chautha": 3,
            "fifth": 4, "5th": 4, "panchwa": 4,
        }
        target_index = 0
        requested_day = extract_day_number(instruction)
        if requested_day is not None:
            target_index = max(requested_day - 1, 0)
        else:
            lowered = instruction.lower()
            for word, index in ordinal_map.items():
                if re.search(rf"\b{re.escape(word)}\b", lowered):
                    target_index = index
                    break
        target_index = min(target_index, len(updated) - 1)
        value_match = re.search(r"(?:to|with)\s+(.+?)\s*$", instruction, re.IGNORECASE)
        requested_value = safe_text(value_match.group(1)) if value_match else ""
        if requested_value and re.search(r"\b(?:heading|title)\b", instruction, re.IGNORECASE):
            updated[target_index]["title"] = requested_value
            return updated
        for field in ("morning", "afternoon", "evening"):
            if requested_value and re.search(rf"\b{field}\b", instruction, re.IGNORECASE):
                updated[target_index][field] = requested_value
                return updated
        change_match = re.search(r"(?:change|replace)\s+(.+?)\s+(?:with|to)\s+(.+)", instruction, re.IGNORECASE)
        remove_match = re.search(r"remove\s+(.+)", instruction, re.IGNORECASE)
        fields = ("title", "morning", "afternoon", "evening")
        for day in updated:
            if change_match:
                old_text, new_text = change_match.group(1).strip(), change_match.group(2).strip()
                for field in fields:
                    day[field] = re.sub(re.escape(old_text), new_text, str(day.get(field, "")), flags=re.IGNORECASE)
                day["places"] = [re.sub(re.escape(old_text), new_text, place, flags=re.IGNORECASE) for place in day.get("places", [])]
            elif remove_match:
                target = remove_match.group(1).strip()
                for field in fields:
                    day[field] = re.sub(re.escape(target), "", str(day.get(field, "")), flags=re.IGNORECASE).strip()
                day["places"] = [place for place in day.get("places", []) if target.lower() not in place.lower()]
        if not change_match and not remove_match:
            editable = updated[1:-1] if len(updated) > 2 else updated
            if editable:
                theme = safe_text(instruction)
                if re.search(r"new\s*year|naya\s+saal", instruction, re.IGNORECASE):
                    theme = "New Year Celebration"
                target = editable[0]
                target["title"] = theme[:80]
                target["morning"] = safe_text(f"Begin the day with destination highlights connected to {theme}.")
                target["afternoon"] = safe_text(f"Continue with sightseeing and local experiences planned around {theme}.")
                target["evening"] = safe_text(f"Enjoy an evening experience that meaningfully reflects {theme}.")
        return updated

    prompt = f"""
Revise this travel plan according to the user's instruction. The instruction is an order, not a suggestion.
Return JSON only as {{"days": [...]}} and preserve the same number of days and each date.
Keep every day object fields: title, morning, afternoon, evening, places.
Implement every requested text/content change exactly and do not silently ignore any part.
If the user asks to add a theme or occasion (for example New Year), rewrite the relevant title and activities so the plan genuinely includes it; do not merely append the instruction as raw text.
If the user says "add/create a travel plan", integrate that requested plan or theme into the existing dated itinerary without changing the number of days.
When activities change, update "places" with real, specific locations mentioned or required by those activities.
Only keep first-day arrival and last-day departure constraints when they do not conflict with the explicit instruction.
Destination: {destination}
User instruction: {instruction}
Current plan: {json.dumps(days, default=str)}
"""
    try:
        user_content: str | list[dict] = prompt
        if reference_image and reference_image.get("bytes"):
            mime_type = reference_image.get("mime_type") or "image/jpeg"
            encoded = base64.b64encode(reference_image["bytes"]).decode("ascii")
            user_content = [
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{encoded}"}},
            ]

        response = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={"Authorization": f"Bearer {api_key}"},
            json={
                "model": os.getenv("OPENROUTER_MODEL", "openrouter/auto"),
                "messages": [
                    {"role": "system", "content": "You are a precise document editor. Execute every requested change and return valid JSON only."},
                    {"role": "user", "content": user_content},
                ],
                "temperature": 0.1,
            },
            timeout=90,
        )
        response.raise_for_status()
        raw_text = response.json()["choices"][0]["message"]["content"]
        payload = json.loads(extract_json_block(raw_text))
        revised = payload.get("days", [])
        if len(revised) != len(days):
            raise ValueError("AI response changed the itinerary length")
        normalized = []
        for index, original in enumerate(days):
            item = revised[index]
            normalized.append(
                {
                    **original,
                    "title": clean_day_title(str(item.get("title", original["title"])), index),
                    "morning": safe_text(str(item.get("morning", original.get("morning", "")))),
                    "afternoon": safe_text(str(item.get("afternoon", original.get("afternoon", "")))),
                    "evening": safe_text(str(item.get("evening", original.get("evening", "")))),
                    "places": [
                        place for place in normalize_place_list(item.get("places", original.get("places", [])))
                        if is_valid_place_name(place)
                    ],
                }
            )
        original_signature = [
            (day.get("title"), day.get("morning"), day.get("afternoon"), day.get("evening"), day.get("places", []))
            for day in days
        ]
        revised_signature = [
            (day.get("title"), day.get("morning"), day.get("afternoon"), day.get("evening"), day.get("places", []))
            for day in normalized
        ]
        if revised_signature == original_signature:
            raise ValueError("AI returned the plan without applying the instruction")
        return normalized
    except Exception:
        updated = [{**day, "places": list(day.get("places", []))} for day in days]
        editable = updated[1:-1] if len(updated) > 2 else updated
        if editable:
            theme = "New Year Celebration" if re.search(r"new\s*year|naya\s+saal", instruction, re.IGNORECASE) else safe_text(instruction)
            target = editable[0]
            target["title"] = theme[:80]
            target["morning"] = safe_text(f"Begin the day with destination highlights connected to {theme}.")
            target["afternoon"] = safe_text(f"Continue with sightseeing and local experiences planned around {theme}.")
            target["evening"] = safe_text(f"Enjoy an evening experience that meaningfully reflects {theme}.")
        return updated


IMAGE_INSTRUCTION_PATTERN = re.compile(
    r"\b(?:image|images|im[a-z]{1,6}ge|im[a-z]{1,6}ges|img|picture|pictures|photo|photos|pic|pics|"
    r"tasveer|tasveeren|tasveerain|tasvir|tasweer)\b",
    re.IGNORECASE,
)


def is_image_change_instruction(instruction: str) -> bool:
    return bool(IMAGE_INSTRUCTION_PATTERN.search(instruction or ""))


def is_image_only_instruction(instruction: str) -> bool:
    """Keep image commands away from the itinerary text editor.

    A mixed command such as "change the first heading and replace its image" is
    still allowed to edit both. Commands that only ask for images (including
    common Roman-Urdu wording) must never be written into plan paragraphs.
    """
    if not is_image_change_instruction(instruction):
        return False
    text_targets = re.compile(
        r"\b(?:text|word|words|sentence|paragraph|para|heading|title|morning|"
        r"afternoon|evening|activity|activities|place|places|itinerary|plan\s+text)\b",
        re.IGNORECASE,
    )
    return not bool(text_targets.search(instruction or ""))


def extract_day_number(instruction: str) -> int | None:
    """Return a 1-based day number from common English/Roman-Urdu phrasing."""
    lowered = (instruction or "").lower()
    for pattern in (
        r"\bday\s*(\d+)(?:st|nd|rd|th)?\b",
        r"\b(\d+)(?:st|nd|rd|th)?\s*day\b",
        r"\bdin\s*(\d+)\b",
        r"\b(\d+)\s*(?:wa|wan|th)\s*din\b",
    ):
        match = re.search(pattern, lowered)
        if match:
            return max(int(match.group(1)), 1)
    ordinal_days = {
        "first": 1, "pehla": 1, "pehli": 1,
        "second": 2, "dusra": 2, "doosra": 2,
        "third": 3, "teesra": 3, "teesri": 3,
        "fourth": 4, "chautha": 4, "chauthi": 4,
        "fifth": 5, "panchwa": 5, "panchwi": 5,
        "sixth": 6,
        "seventh": 7, "eighth": 8, "ninth": 9, "tenth": 10,
    }
    if re.search(r"\b(?:day|din)\b", lowered):
        for word, number in ordinal_days.items():
            if re.search(rf"\b{re.escape(word)}\b", lowered):
                return number
    return None


def extract_image_position(instruction: str) -> int | None:
    """Return a zero-based image position when the prompt names one."""
    lowered = (instruction or "").lower()
    image_term = r"(?:image|picture|photo|pic|tasveer|tasvir|tasweer)"
    ordinal_positions = {
        "first": 0, "1st": 0, "pehli": 0, "pehla": 0,
        "second": 1, "2nd": 1, "dusri": 1, "doosri": 1, "dusra": 1,
        "third": 2, "3rd": 2, "teesri": 2, "teesra": 2,
        "fourth": 3, "4th": 3, "chauthi": 3, "chautha": 3,
    }
    for word, position in ordinal_positions.items():
        if re.search(rf"\b{re.escape(word)}\s+{image_term}\b", lowered):
            return position
        if re.search(rf"\b{image_term}\s+{re.escape(word)}\b", lowered):
            return position
    numeric = re.search(rf"\b{image_term}\s*(\d+)\b", lowered)
    if numeric:
        return max(int(numeric.group(1)) - 1, 0)
    return None


def requests_multiple_images(instruction: str) -> bool:
    return bool(re.search(
        r"\b(?:images|pictures|photos|pics|tasveeren|tasveerain)\b",
        instruction or "",
        re.IGNORECASE,
    ))


def build_unique_replacement_image(
    result: dict,
    queries: list[str],
    retired_urls: set[str] | None = None,
    retired_fingerprints: set[str] | None = None,
) -> dict | None:
    """Find a relevant photo that is not already used anywhere in the plan."""
    existing_entries = [
        entry for entries in result.get("day_images", []) for entry in entries
    ]
    excluded_urls = {
        (entry.get("url") or "").split("?", 1)[0]
        for entry in existing_entries if entry.get("url")
    }
    excluded_urls.update(retired_urls or set())
    existing_fingerprints = {
        fingerprint for fingerprint in (
            image_fingerprint(entry.get("path")) for entry in existing_entries
        ) if fingerprint
    }
    existing_fingerprints.update(retired_fingerprints or set())
    destination = result.get("destination", "")
    for query in queries:
        for _ in range(3):
            candidate = build_place_image(query, destination, excluded_urls)
            if not candidate or not candidate.get("path"):
                break
            candidate_url = (candidate.get("url") or "").split("?", 1)[0]
            candidate_fingerprint = image_fingerprint(candidate.get("path"))
            if candidate_url:
                excluded_urls.add(candidate_url)
            if candidate_fingerprint and candidate_fingerprint in existing_fingerprints:
                continue
            return candidate
    return None


def apply_prompt_image_changes(result: dict, instruction: str, uploaded_path: str | None = None) -> bool:
    lowered = (instruction or "").lower()
    image_words = r"(?:image|images|im[a-z]{1,6}ge|im[a-z]{1,6}ges|img|picture|pictures|photo|photos|pic|pics|tasveer|tasveeren|tasveerain|tasvir|tasweer)"
    if not is_image_change_instruction(lowered):
        return False

    destination = result["destination"]
    replacement_match = re.search(r"(?:to|with)\s+(.+?)\s*$", instruction, re.IGNORECASE)
    replacement_query = safe_text(replacement_match.group(1)) if replacement_match else ""
    replacement_query = re.sub(rf"\b{image_words}\b", "", replacement_query, flags=re.IGNORECASE).strip(" -")

    if "cover" in lowered:
        if uploaded_path:
            result["cover_path"] = uploaded_path
            return True
        query = replacement_query or destination
        new_entry = build_place_image(query, destination)
        if new_entry and new_entry.get("path"):
            result["cover_path"] = new_entry["path"]
            return True

    requested_day = extract_day_number(lowered)
    if requested_day is not None:
        if requested_day > len(result["day_images"]):
            return False
        day_index = requested_day - 1
        entries = result["day_images"][day_index]
        day_places = result["days"][day_index].get("places") or [destination]
        requested_position = extract_image_position(lowered)

        if requests_multiple_images(lowered) and not uploaded_path:
            changed = False
            retired_urls = {
                (entry.get("url") or "").split("?", 1)[0]
                for entry in entries if entry.get("url")
            }
            retired_fingerprints = {
                fingerprint for fingerprint in (
                    image_fingerprint(entry.get("path")) for entry in entries
                ) if fingerprint
            }
            target_count = min(max(len(entries), 2), len(day_places), 2)
            for image_index in range(target_count):
                ordered_places = day_places[image_index:] + day_places[:image_index]
                replacement = build_unique_replacement_image(
                    result,
                    ordered_places,
                    retired_urls,
                    retired_fingerprints,
                )
                if replacement and replacement.get("path"):
                    if image_index < len(entries):
                        entries[image_index] = replacement
                    else:
                        entries.append(replacement)
                    changed = True
            return changed

        image_index = min(requested_position or 0, max(len(entries) - 1, 0))
        place_index = min(image_index, len(day_places) - 1)
        query = replacement_query or day_places[place_index]
        ordered_queries = [query, *day_places[place_index + 1:], *day_places[:place_index]]
        replacement = (
            {"place": query, "path": uploaded_path}
            if uploaded_path
            else build_unique_replacement_image(result, ordered_queries)
        )
        if replacement and replacement.get("path"):
            if entries:
                entries[image_index] = replacement
            else:
                result["day_images"][day_index] = [replacement]
            return True

    ordinal_map = {"first": 0, "1st": 0, "second": 1, "2nd": 1, "third": 2, "3rd": 2, "fourth": 3, "4th": 3}
    requested_position = None
    for word, position in ordinal_map.items():
        if re.search(rf"\b{re.escape(word)}\b", lowered):
            requested_position = position
            break
    if requested_position is not None:
        flattened = []
        for day_index, entries in enumerate(result["day_images"]):
            for image_index, entry in enumerate(entries):
                flattened.append((day_index, image_index, entry))
        if flattened:
            day_index, image_index, old_entry = flattened[min(requested_position, len(flattened) - 1)]
            query = replacement_query or old_entry.get("place") or destination
            replacement = (
                {"place": query, "path": uploaded_path}
                if uploaded_path
                else build_unique_replacement_image(result, [query])
            )
            if replacement and replacement.get("path"):
                result["day_images"][day_index][image_index] = replacement
                return True

    if uploaded_path:
        for day_index, entries in enumerate(result["day_images"]):
            if entries:
                old_place = entries[0].get("place") or destination
                result["day_images"][day_index][0] = {"place": old_place, "path": uploaded_path}
                return True
        result["cover_path"] = uploaded_path
        return True

    new_cover, new_day_images = build_image_bundle(result["days"], destination)
    result["cover_path"] = new_cover
    result["day_images"] = new_day_images
    return True


def save_pdf_ready_image(content: bytes, max_width: int = 1000, max_height: int = 750) -> str | None:
    try:
        image = Image.open(BytesIO(content))
        image.load()
        if image.width < 120 or image.height < 120:
            return None
        if image.mode not in ("RGB", "L"):
            image = image.convert("RGB")
        elif image.mode == "L":
            image = image.convert("RGB")
        image.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as temp_file:
            image.save(temp_file.name, format="JPEG", quality=88, optimize=True, progressive=True)
            return temp_file.name
    except (UnidentifiedImageError, OSError, ValueError):
        return None


def download_image(url: str, suffix: str = ".jpg"):
    decoded_url = requests.utils.unquote(url).lower()
    rejected_markers = (
        "logo", "emblem", "crest", "seal", "badge", "wordmark", "symbol",
        "icon", "coat_of_arms", "coat-of-arms", "portrait", "headshot",
    )
    clean_path = decoded_url.split("?", 1)[0]
    if any(marker in decoded_url for marker in rejected_markers) or clean_path.endswith((".svg", ".png")):
        return None
    # Retry downloads a few times with backoff to reduce intermittent failures.
    headers = {
        "User-Agent": "AI-Travel-Planner/1.0 (+https://github.com/abdullahkrajpoot-dot/Ai-Travel-Planner-Agent)",
        "Accept": "image/jpeg,image/png,image/*;q=0.8,*/*;q=0.5",
    }
    for attempt in range(1, 4):
        try:
            response = requests.get(url, timeout=30, allow_redirects=True, headers=headers)
            response.raise_for_status()
            content_type = (response.headers.get("Content-Type", "") or "").lower()
            if "image/" not in content_type:
                return None

            if len(response.content) < 1024:
                return None

            return save_pdf_ready_image(response.content)
        except Exception:
            if attempt < 3:
                time.sleep(0.8 * attempt)
                continue
            return None


def _candidate_place_queries(place_name: str, destination: str | None) -> list[str]:
    candidates = [safe_text(place_name)]
    if destination and destination.lower() not in place_name.lower():
        candidates.append(safe_text(f"{place_name}, {destination}"))
    if "," in place_name:
        candidates.append(safe_text(place_name.split(",", 1)[0]))
    unique = []
    for item in candidates:
        if item and item not in unique:
            unique.append(item)
    return unique


def _relevance_tokens(value: str) -> set[str]:
    ignored = {
        "the", "and", "for", "with", "from", "city", "travel", "tourist",
        "tourism", "photo", "image", "view", "landmark", "attraction",
    }
    return {
        token for token in slugify(safe_text(value)).split("_")
        if len(token) > 2 and token not in ignored
    }


def _is_relevant_image_result(query: str, label: str) -> bool:
    query_tokens = _relevance_tokens(query)
    label_tokens = _relevance_tokens(label)
    if not query_tokens or not label_tokens:
        return False
    overlap = len(query_tokens.intersection(label_tokens))
    required = 1 if len(query_tokens) <= 2 else 2
    if overlap < required:
        return False

    primary_name = query.split(",", 1)[0]
    primary_tokens = _relevance_tokens(primary_name)
    primary_overlap = len(primary_tokens.intersection(label_tokens))
    primary_required = 1 if len(primary_tokens) <= 1 else 2
    banned_result_tokens = {
        "logo", "basketball", "football", "baseball", "hockey", "team", "club",
        "season", "jersey", "mascot", "album", "film", "television", "game",
        "emblem", "crest", "seal", "badge", "wordmark", "symbol", "icon",
        "portrait", "headshot", "painting", "drawing", "illustration", "poster",
    }
    return primary_overlap >= primary_required and not banned_result_tokens.intersection(label_tokens)


def _image_url_matches_query(query: str, url: str) -> bool:
    """Require the actual image filename to describe the requested place."""
    decoded = requests.utils.unquote(url).split("?", 1)[0]
    filename = decoded.rsplit("/", 1)[-1]
    filename = re.sub(r"^\d+px-", "", filename, flags=re.IGNORECASE)
    return _is_relevant_image_result(query, filename)


def _wikipedia_summary_image(query: str) -> str | None:
    title = requests.utils.quote(query.strip().replace(" ", "_"))
    url = f"https://en.wikipedia.org/api/rest_v1/page/summary/{title}"
    try:
        response = requests.get(
            url,
            timeout=20,
            headers={
                "User-Agent": "AI-Travel-Planner/1.0",
                "Accept": "application/json",
            },
        )
        if response.status_code != 200:
            return None
        payload = response.json()
        context = " ".join(
            str(payload.get(field, ""))
            for field in ("title", "description", "extract")
        )
        if not _is_relevant_image_result(query, context):
            return None
        image_obj = payload.get("originalimage") or payload.get("thumbnail") or {}
        image_url = image_obj.get("source")
        if image_url and image_url.startswith("http"):
            return image_url
        return None
    except Exception:
        return None


def _wikipedia_page_image(query: str) -> str | None:
    try:
        payload = _wiki_api({
            "action": "query",
            "prop": "pageimages",
            "piprop": "original|thumbnail",
            "pithumbsize": 1200,
            "redirects": 1,
            "titles": query,
        })
        pages = payload.get("query", {}).get("pages", {})
        for page in pages.values():
            if not _is_relevant_image_result(query, str(page.get("title", ""))):
                continue
            image_obj = page.get("original") or page.get("thumbnail") or {}
            image_url = image_obj.get("source")
            if image_url and image_url.startswith("http"):
                return image_url
        return None
    except Exception:
        return None


def _wikipedia_search_image(query: str) -> str | None:
    search_url = "https://en.wikipedia.org/w/api.php"
    try:
        response = requests.get(
            search_url,
            timeout=25,
            params={
                "action": "query",
                "list": "search",
                "format": "json",
                "utf8": 1,
                "srlimit": 5,
                "srsearch": query,
            },
            headers=WIKI_HEADERS,
        )
        response.raise_for_status()
        results = response.json().get("query", {}).get("search", [])
        if not results:
            return None

        ranked_titles = []
        for row in results:
            title = safe_text(row.get("title", ""))
            overlap = len(_relevance_tokens(query).intersection(_relevance_tokens(title)))
            ranked_titles.append((overlap, title))
        ranked_titles.sort(reverse=True, key=lambda item: item[0])

        for overlap, title in ranked_titles:
            if not _is_relevant_image_result(query, title):
                continue
            image_url = _wikipedia_page_image(title)
            if image_url:
                return image_url
            image_url = _wikipedia_summary_image(title)
            if image_url:
                return image_url
        return None
    except Exception:
        return None


def _commons_search_image(query: str, excluded_urls: set[str] | None = None) -> str | None:
    """Search Wikimedia Commons for a matching file and return a direct image URL."""
    try:
        search_params = {
            "action": "query",
            "format": "json",
            "list": "search",
            "srsearch": query,
            "srnamespace": 6,  # File namespace on Commons
            "srlimit": 12,
        }
        resp = requests.get("https://commons.wikimedia.org/w/api.php", params=search_params, timeout=20, headers=WIKI_HEADERS)
        resp.raise_for_status()
        data = resp.json()
        results = data.get("query", {}).get("search", [])
        if not results:
            return None

        query_tokens = _relevance_tokens(query)
        ranked_results = []
        for row in results:
            title = safe_text(row.get("title", ""))
            title_tokens = _relevance_tokens(title)
            overlap = len(query_tokens.intersection(title_tokens))
            ranked_results.append((overlap, row))
        ranked_results.sort(reverse=True, key=lambda item: item[0])

        for overlap, row in ranked_results:
            if not _is_relevant_image_result(query, safe_text(row.get("title", ""))):
                continue
            title = row.get("title")
            if not title:
                continue
            # Get imageinfo for this file title
            info_params = {
                "action": "query",
                "format": "json",
                "prop": "imageinfo",
                "iiprop": "url",
                "iiurlwidth": 1200,
                "titles": title,
            }
            info_resp = requests.get("https://commons.wikimedia.org/w/api.php", params=info_params, timeout=20, headers=WIKI_HEADERS)
            info_resp.raise_for_status()
            pages = info_resp.json().get("query", {}).get("pages", {})
            for page in pages.values():
                imageinfo = page.get("imageinfo") or []
                if imageinfo:
                    url = imageinfo[0].get("thumburl") or imageinfo[0].get("url")
                    normalized_url = (url or "").split("?", 1)[0]
                    if url and url.startswith("http") and normalized_url not in (excluded_urls or set()):
                        return url
        return None
    except Exception:
        return None


def _google_custom_search_images(query: str) -> list[str]:
    api_key = os.getenv("GOOGLE_API_KEY", "").strip()
    cse_id = os.getenv("GOOGLE_CSE_ID", "").strip()
    if not api_key or not cse_id:
        return []
    try:
        response = requests.get(
            "https://www.googleapis.com/customsearch/v1",
            params={
                "key": api_key,
                "cx": cse_id,
                "q": query,
                "searchType": "image",
                "num": 6,
                "safe": "active",
                "imgSize": "large",
                "rights": "cc_publicdomain,cc_attribute,cc_sharealike",
            },
            timeout=20,
            headers={"User-Agent": WIKI_HEADERS["User-Agent"]},
        )
        response.raise_for_status()
        urls = []
        for item in response.json().get("items", []):
            link = item.get("link", "")
            if link.startswith("http"):
                urls.append(link)
        return urls
    except Exception:
        return []


def _duckduckgo_image_search_urls(query: str) -> list[str]:
    try:
        page = requests.get(
            "https://duckduckgo.com/",
            params={"q": query},
            timeout=20,
            headers={"User-Agent": WIKI_HEADERS["User-Agent"]},
        )
        page.raise_for_status()
        token_match = re.search(r"vqd=['\"]([^'\"]+)['\"]", page.text)
        if not token_match:
            return []
        response = requests.get(
            "https://duckduckgo.com/i.js",
            params={
                "l": "us-en",
                "o": "json",
                "q": query,
                "vqd": token_match.group(1),
                "f": ",,,",
                "p": "1",
            },
            timeout=25,
            headers={
                "User-Agent": WIKI_HEADERS["User-Agent"],
                "Referer": "https://duckduckgo.com/",
            },
        )
        response.raise_for_status()
        urls = []
        for item in response.json().get("results", [])[:8]:
            image_url = item.get("image") or item.get("thumbnail")
            if image_url and image_url.startswith("http"):
                urls.append(image_url)
        return urls
    except Exception:
        return []


def _web_image_search_urls(query: str) -> list[str]:
    clean_query = f"{safe_text(query)} landmark tourist place travel photo"
    urls = []
    for provider_urls in (
        _google_custom_search_images(clean_query),
        _duckduckgo_image_search_urls(clean_query),
    ):
        for url in provider_urls:
            if url not in urls:
                urls.append(url)
    return urls


def build_place_image(
    place_name: str,
    destination: str | None = None,
    excluded_urls: set[str] | None = None,
):
    # Prefer file-level Commons results. Any Wikipedia fallback must also have
    # an image filename that explicitly matches the requested location.
    queries = _candidate_place_queries(place_name, destination)

    for query in queries:
        commons_url = _commons_search_image(query, excluded_urls)
        if commons_url and _image_url_matches_query(query, commons_url):
            path = download_image(commons_url, suffix=".jpg")
            if path:
                return {"place": place_name, "url": commons_url, "path": path}

        for candidate_url in (
            _wikipedia_page_image(query),
            _wikipedia_summary_image(query),
            _wikipedia_search_image(query),
        ):
            normalized_candidate = (candidate_url or "").split("?", 1)[0]
            if (
                not candidate_url
                or normalized_candidate in (excluded_urls or set())
                or not _image_url_matches_query(query, candidate_url)
            ):
                continue
            path = download_image(candidate_url, suffix=".jpg")
            if path:
                return {"place": place_name, "url": candidate_url, "path": path}

    return {"place": place_name, "url": None, "path": None}


def build_image_bundle(days: list[dict], destination: str):
    cover_candidates = []
    for day in days[:2]:
        cover_candidates.extend(day.get("places", [])[:2])
    if not cover_candidates:
        cover_candidates = [destination]

    unique_places = []
    planned_places = [place for day in days for place in day.get("places", [])[:4]]
    for place in cover_candidates + planned_places:
        if place not in unique_places:
            unique_places.append(place)

    results = {}
    with ThreadPoolExecutor(max_workers=12) as executor:
        future_map = {executor.submit(build_place_image, place, destination): place for place in unique_places}
        for future in as_completed(future_map):
            place = future_map[future]
            try:
                results[place] = future.result()
            except Exception:
                results[place] = {"place": place, "url": None, "path": None}

    cover_path = None
    used_sources: set[str] = set()
    used_fingerprints: set[str] = set()
    used_places: set[str] = set()
    for place in cover_candidates:
        entry = results.get(place, {})
        path = entry.get("path")
        if path:
            cover_path = path
            source = (entry.get("url") or "").split("?", 1)[0]
            if source:
                used_sources.add(source)
            fingerprint = image_fingerprint(path)
            if fingerprint:
                used_fingerprints.add(fingerprint)
            break

    day_images: list[list[dict]] = []
    for day in days:
        entries: list[dict] = []
        for place in day.get("places", [])[:4]:
            entry = results.get(place, {"place": place, "url": None, "path": None})
            path = entry.get("path")
            source = (entry.get("url") or "").split("?", 1)[0] or path
            fingerprint = image_fingerprint(path) if path else None
            is_duplicate = source in used_sources or (fingerprint and fingerprint in used_fingerprints)
            if source and not is_duplicate and place not in used_places:
                entries.append(entry)
                used_sources.add(source)
                if fingerprint:
                    used_fingerprints.add(fingerprint)
                used_places.add(place)
                if len(entries) == 2:
                    break

        day_images.append(entries[:2])

    return cover_path, day_images


def image_fingerprint(path: str | None) -> str | None:
    """Return a compact visual fingerprint so repeated photos are not shown twice."""
    if not path:
        return None
    try:
        with Image.open(path) as image:
            grayscale = image.convert("L").resize((16, 16))
            pixels = list(grayscale.getdata())
            average = sum(pixels) / len(pixels)
            return "".join("1" if pixel >= average else "0" for pixel in pixels)
    except (OSError, UnidentifiedImageError):
        return None


class TravelPlanPDF(FPDF):
    def header(self):
        if self.page_no() >= 3:
            self.set_font("Helvetica", "B", 8)
            self.set_text_color(0, 0, 0)
            self.cell(0, 6, f"Page {self.page_no()} of {{nb}}", align="R")
            self.ln(3)


def draw_cover_page(pdf: TravelPlanPDF, client: str, title: str, destination: str, start_date: date, end_date: date, cover_path: str | None):
    pdf.add_page()
    pdf.set_fill_color(255, 255, 255)
    pdf.rect(0, 0, 210, 297, "F")
    pdf.set_xy(0, 18)
    pdf.set_font("Helvetica", "B", 18)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(210, 10, safe_text(client.upper()), align="C")
    if cover_path:
        try:
            pdf.image(cover_path, x=10, y=38, w=190, h=86)
        except Exception:
            pass
    pdf.set_xy(0, 132)
    pdf.set_font("Helvetica", "B", 16)
    pdf.set_text_color(14, 44, 85)
    pdf.cell(210, 10, safe_text(title), align="C")
    pdf.set_xy(0, 146)
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(53, 90, 129)
    pdf.cell(210, 8, safe_text(f"{compact_date(start_date)} - {compact_date(end_date)}"), align="C")
    pdf.set_xy(0, 156)
    pdf.set_font("Helvetica", "", 11)
    pdf.set_text_color(70, 85, 110)
    pdf.cell(210, 7, safe_text(destination), align="C")


def draw_summary_page(pdf: TravelPlanPDF, days: list[dict]):
    pdf.add_page()
    pdf.set_xy(14, 18)
    pdf.set_font("Helvetica", "B", 18)
    pdf.set_text_color(20, 54, 109)
    pdf.cell(0, 10, "Trip Summary")
    pdf.ln(14)
    for item in days:
        pdf.set_draw_color(210, 220, 230)
        pdf.line(14, pdf.get_y() + 2, 196, pdf.get_y() + 2)
        pdf.ln(5)
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(35, 68, 115)
        pdf.cell(110, 8, safe_text(item["title"]))
        pdf.set_font("Helvetica", "", 11)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(0, 8, safe_text(format_day_label(item["date"])), align="R")
        pdf.ln(7)
        pdf.set_x(14)
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(90, 100, 120)
        pdf.multi_cell(182, 5, safe_text(", ".join(item.get("places", []))))


def draw_timeline_marker(pdf: TravelPlanPDF, y_top: float):
    x = 20
    pdf.set_draw_color(190, 200, 210)
    pdf.line(x, y_top, x, y_top + 24)
    pdf.ellipse(x - 4, y_top, 8, 8, "D")
    pdf.set_xy(x - 2, y_top + 1.2)
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_text_color(50, 50, 50)
    pdf.cell(4, 4, "+", align="C")


def draw_image_grid(pdf: TravelPlanPDF, image_entries: list[dict], x: float, y: float):
    if not image_entries:
        return y
    positions = [
        (x, y, 68, 40),
        (x + 70, y, 34, 19),
        (x + 106, y, 34, 19),
        (x + 70, y + 21, 34, 19),
    ]
    for idx, image_entry in enumerate(image_entries[:4]):
        px, py, pw, ph = positions[idx]
        image_path = image_entry.get("path")
        if not image_path:
            continue
        try:
            pdf.image(image_path, x=px, y=py, w=pw, h=ph)
        except Exception:
            continue
    return y + 44


def draw_day_block(pdf: TravelPlanPDF, item: dict, image_entries: list[dict]):
    if pdf.get_y() > 230:
        pdf.add_page()
    start_y = pdf.get_y() + 4
    draw_timeline_marker(pdf, start_y)
    content_x = 32
    pdf.set_xy(content_x, start_y - 1)
    pdf.set_font("Helvetica", "B", 15)
    pdf.set_text_color(52, 86, 129)
    pdf.cell(0, 8, safe_text(format_day_label(item["date"])))
    pdf.ln(8)
    pdf.set_x(content_x)
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 7, safe_text(item["title"]))
    pdf.ln(7)
    for label, content in [("Morning", item["morning"]), ("Afternoon", item["afternoon"]), ("Evening", item["evening"])]:
        if not safe_text(content):
            continue
        pdf.set_x(content_x)
        pdf.set_font("Helvetica", "B", 10)
        pdf.write(5, f"{label}: ")
        pdf.set_font("Helvetica", "", 10)
        pdf.multi_cell(150, 5, safe_text(content))
    if item.get("places"):
        pdf.set_x(content_x)
        pdf.set_font("Helvetica", "B", 10)
        pdf.write(5, "Places: ")
        pdf.set_font("Helvetica", "", 10)
        pdf.multi_cell(150, 5, safe_text(", ".join(item["places"])))
    if image_entries:
        image_y = pdf.get_y() + 2
        last_y = draw_image_grid(pdf, image_entries, content_x, image_y)
        pdf.set_y(last_y + 7)
    else:
        pdf.ln(7)
    pdf.set_draw_color(220, 228, 236)
    pdf.line(14, pdf.get_y(), 196, pdf.get_y())
    pdf.ln(5)


def build_pdf(client: str, title: str, destination: str, start_date: date, end_date: date, days: list[dict], cover_path: str | None, day_images: list[list[dict]]) -> bytes:
    pdf = TravelPlanPDF("P", "mm", "A4")
    pdf.set_compression(True)
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=16)
    draw_cover_page(pdf, client, title, destination, start_date, end_date, cover_path)
    draw_summary_page(pdf, days)
    pdf.add_page()
    pdf.set_y(12)
    for idx, item in enumerate(days):
        draw_day_block(pdf, item, day_images[idx] if idx < len(day_images) else [])
    output = pdf.output()
    if isinstance(output, bytearray):
        return bytes(output)
    if isinstance(output, str):
        return output.encode("latin-1", errors="ignore")
    return output


def build_docx(client: str, title: str, destination: str, start_date: date, end_date: date, days: list[dict], cover_path: str | None, day_images: list[list[dict]]) -> bytes:
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    styles = document.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(10.5)
    styles["Title"].font.name = "Aptos Display"
    styles["Title"].font.size = Pt(24)
    styles["Title"].font.color.rgb = None
    for style_name in ("Heading 1", "Heading 2"):
        styles[style_name].font.name = "Aptos Display"
        styles[style_name].font.color.rgb = None

    title_paragraph = document.add_paragraph(style="Title")
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_paragraph.add_run(safe_text(title))
    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.add_run(safe_text(client)).bold = True
    subtitle.add_run(f"\n{safe_text(destination)}")
    subtitle.add_run(f"\n{compact_date(start_date)} - {compact_date(end_date)}")

    if cover_path and Path(cover_path).exists():
        cover = document.add_paragraph()
        cover.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cover.add_run().add_picture(cover_path, width=Inches(6.25))

    document.add_heading("Trip Summary", level=1)
    for item in days:
        summary = document.add_paragraph(style="List Bullet")
        summary.add_run(f"{format_day_label(item['date'])}: ").bold = True
        summary.add_run(safe_text(item.get("title", "")))

    for index, item in enumerate(days):
        document.add_heading(f"Day {index + 1}  {format_day_label(item['date'])}", level=1)
        document.add_heading(safe_text(item.get("title", "")), level=2)
        for label, field in (("Morning", "morning"), ("Afternoon", "afternoon"), ("Evening", "evening")):
            paragraph = document.add_paragraph()
            paragraph.add_run(f"{label}: ").bold = True
            paragraph.add_run(safe_text(item.get(field, "")))
        if item.get("places"):
            paragraph = document.add_paragraph()
            paragraph.add_run("Places: ").bold = True
            paragraph.add_run(safe_text(", ".join(item["places"])))
        entries = day_images[index] if index < len(day_images) else []
        for entry in entries[:2]:
            image_path = entry.get("path")
            if image_path and Path(image_path).exists():
                image_paragraph = document.add_paragraph()
                image_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                image_paragraph.add_run().add_picture(image_path, width=Inches(5.8))
                caption = document.add_paragraph(safe_text(entry.get("place", "Travel image")))
                caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
                caption.runs[0].italic = True

    output = BytesIO()
    document.save(output)
    return output.getvalue()


def cleanup_temp_images(cover_path: str | None, day_images: list[list[dict]]):
    paths = set()
    if cover_path:
        paths.add(cover_path)
    for entries in day_images:
        for entry in entries:
            path = entry.get("path")
            if path:
                paths.add(path)
    for file_path in paths:
        try:
            path_obj = Path(file_path)
            if path_obj.exists():
                path_obj.unlink()
        except Exception:
            continue


def save_uploaded_image(uploaded_file) -> str | None:
    """Normalize an uploaded replacement image for reliable PDF rendering."""
    if uploaded_file is None:
        return None
    try:
        return save_pdf_ready_image(uploaded_file.getvalue())
    except Exception:
        return None


def rebuild_result_documents(result: dict) -> None:
    result["pdf_bytes"] = build_pdf(
        client=result["client_name"],
        title=result["plan_title"],
        destination=result["destination"],
        start_date=result["start_date"],
        end_date=result["end_date"],
        days=result["days"],
        cover_path=result["cover_path"],
        day_images=result["day_images"],
    )
    result["docx_bytes"] = build_docx(
        client=result["client_name"],
        title=result["plan_title"],
        destination=result["destination"],
        start_date=result["start_date"],
        end_date=result["end_date"],
        days=result["days"],
        cover_path=result["cover_path"],
        day_images=result["day_images"],
    )


def repair_non_visitable_places(result: dict) -> bool:
    """Replace species/topics in an existing session with real attractions."""
    days = result.get("days", [])
    if not any(not is_valid_place_name(place) for day in days for place in day.get("places", [])):
        return False

    groups = get_landmark_groups_for_days(result.get("destination", ""), len(days))
    used = {
        place for day in days for place in day.get("places", [])
        if is_valid_place_name(place)
    }
    for index, day in enumerate(days):
        if index == 0 or index == len(days) - 1:
            day["places"] = []
            continue
        valid_places = [place for place in day.get("places", []) if is_valid_place_name(place)]
        candidates = groups[index % len(groups)] if groups else []
        for candidate in candidates:
            if len(valid_places) >= 4:
                break
            if candidate not in used and is_valid_place_name(candidate):
                valid_places.append(candidate)
                used.add(candidate)
        day["places"] = valid_places[:4]

    result["cover_path"], result["day_images"] = build_image_bundle(days, result["destination"])
    rebuild_result_documents(result)
    result["ui_version"] = time.time_ns()
    return True


def day_to_editor_html(day: dict) -> str:
    def escaped(value: str) -> str:
        return html.escape(str(value or ""))

    places = ", ".join(day.get("places", []))
    return (
        f"<h2>{escaped(day.get('title', ''))}</h2>"
        f"<p><strong>Morning</strong></p><p>{escaped(day.get('morning', ''))}</p>"
        f"<p><strong>Afternoon</strong></p><p>{escaped(day.get('afternoon', ''))}</p>"
        f"<p><strong>Evening</strong></p><p>{escaped(day.get('evening', ''))}</p>"
        f"<p><strong>Places</strong></p><p>{escaped(places)}</p>"
    )


def editor_html_to_day(editor_html: str, original: dict) -> dict:
    soup = BeautifulSoup(editor_html or "", "html.parser")
    updated = dict(original)
    heading = soup.find(["h1", "h2", "h3"])
    if heading and heading.get_text(" ", strip=True):
        updated["title"] = heading.get_text(" ", strip=True)

    values = {}
    for strong in soup.find_all("strong"):
        label = strong.get_text(" ", strip=True).lower()
        if label not in {"morning", "afternoon", "evening", "places"}:
            continue
        label_block = strong.find_parent(["p", "div"])
        value_block = label_block.find_next_sibling(["p", "div"]) if label_block else None
        if value_block:
            values[label] = value_block.get_text(" ", strip=True)

    for field in ("morning", "afternoon", "evening"):
        if values.get(field):
            updated[field] = values[field]
    if "places" in values:
        updated["places"] = [item.strip() for item in values["places"].split(",") if item.strip()][:4]
    updated["editor_html"] = editor_html
    return updated


st.markdown(
    """
    <style>
    :root, html, body { color-scheme: light !important; background: #f4f9fd !important; }
    .stApp, [data-testid="stAppViewContainer"] {
        background: linear-gradient(135deg, #edf7ff 0%, #f8fbff 52%, #edf9fb 100%) !important;
        color: #12365a !important;
    }
    [data-testid="stMainBlockContainer"] {
        padding-top: 2.5rem !important;
    }
    [data-testid="stDateInputQuickSelect"] {
        display: none !important;
    }
    [data-testid="stSidebar"] {
        background: #f7fbff !important;
        border-right: 1px solid #b8d8ed !important;
    }
    [data-testid="stSidebar"] * { color: #103e64 !important; }
    .main-header {
        color: white !important; font-size: 2.15rem !important; padding: 2.15rem 2.5rem;
        min-height: 150px; display: flex; align-items: center;
        border-radius: 0 0 22px 22px; margin: -1rem 0 1.4rem 0 !important;
        background: linear-gradient(110deg, #1768aa, #178cc0 64%, #2bb0b3);
        box-shadow: 0 14px 28px rgba(27, 105, 156, .16);
    }
    .sub-header { color: #3b607d !important; margin-bottom: 1.5rem !important; }
    .stTextInput input, .stDateInput input, .stTextArea textarea {
        background: white !important; color: #173f61 !important;
        border: 1px solid #a8cce4 !important; border-radius: 10px !important;
    }
    div[data-baseweb="input"], div[data-baseweb="base-input"] {
        background: white !important; border-radius: 10px !important;
    }
    [data-testid="stSidebar"] .stTextInput > div > div > input,
    [data-testid="stSidebar"] .stTextInput input {
        background: #ffffff !important;
        color: #123f63 !important;
        -webkit-text-fill-color: #123f63 !important;
        border: 1px solid #9fc8e2 !important;
        box-shadow: none !important;
    }
    [data-testid="stSidebar"] .stTextInput > div > div > input::placeholder,
    [data-testid="stSidebar"] .stTextInput input::placeholder {
        color: #6d8ba1 !important;
        -webkit-text-fill-color: #6d8ba1 !important;
        opacity: 1 !important;
    }
    [data-testid="stSidebar"] div[data-baseweb="input"],
    [data-testid="stSidebar"] div[data-baseweb="base-input"] {
        background: #ffffff !important;
        border-color: #9fc8e2 !important;
    }
    [data-testid="stSidebar"] [data-testid="stTextInput"] div[data-baseweb="input"],
    [data-testid="stSidebar"] [data-testid="stTextInput"] div[data-baseweb="input"] > div,
    [data-testid="stSidebar"] [data-testid="stTextInput"] div[data-baseweb="base-input"],
    [data-testid="stSidebar"] [data-testid="stTextInputRootElement"],
    [data-testid="stSidebar"] [data-testid="stTextInputRootElement"]:focus-within,
    [data-testid="stSidebar"] [data-testid="stTextInput"] div:has(> input:focus),
    [data-testid="stSidebar"] [data-testid="stTextInput"] input:focus,
    [data-testid="stSidebar"] [data-testid="InputInstructions"] {
        background: #ffffff !important;
        background-color: #ffffff !important;
        color: #176895 !important;
        -webkit-text-fill-color: #176895 !important;
    }
    [data-testid="stSidebar"] [data-testid="stTextInputRootElement"] {
        border-color: #9fc8e2 !important;
    }
    [data-testid="stSidebar"] [data-testid="stTextInputRootElement"]:focus-within {
        border-color: #268fbd !important;
        box-shadow: 0 0 0 1px #268fbd !important;
    }
    [data-testid="stSidebar"] [data-testid="stTextInput"] div[data-baseweb="input"]:focus-within {
        background: #ffffff !important;
        border-color: #268fbd !important;
        box-shadow: 0 0 0 1px #268fbd !important;
    }
    [data-testid="stSidebar"] .stButton > button {
        color: #ffffff !important;
    }
    [data-testid="stDateInput"] div[data-baseweb="input"],
    [data-testid="stDateInput"] div[data-baseweb="base-input"],
    [data-testid="stDateInput"] > div > div {
        background: #eef7ff !important;
        border-color: #82b9db !important;
    }
    [data-testid="stDateInput"] input {
        background: #eef7ff !important;
        color: #123f63 !important;
        -webkit-text-fill-color: #123f63 !important;
        opacity: 1 !important;
        font-weight: 600 !important;
    }
    [data-testid="stDateInput"] svg,
    [data-testid="stDateInput"] button,
    [data-testid="stDateInput"] span {
        color: #176895 !important;
        fill: #176895 !important;
    }
    [data-testid="stMain"] .stTextInput > div > div > input,
    [data-testid="stMain"] .stTextArea textarea {
        background: #ffffff !important;
        color: #163f61 !important;
        -webkit-text-fill-color: #163f61 !important;
        border: 1px solid #c4dceb !important;
        box-shadow: none !important;
    }
    [data-testid="stMain"] .stTextInput > label,
    [data-testid="stMain"] .stTextArea > label {
        color: #34617e !important;
        font-weight: 600 !important;
    }
    [data-testid="stTabs"] [data-baseweb="tab-list"] {
        gap: 1.25rem;
        border-bottom: 1px solid #b8cddd;
    }
    [data-testid="stTabs"] [data-baseweb="tab"] {
        color: #174a6d !important;
        font-weight: 650;
    }
    [data-testid="stTabs"] [aria-selected="true"] {
        color: #0c7da7 !important;
    }
    .stButton > button, .stDownloadButton > button,
    [data-testid="stFormSubmitButton"] > button,
    .stFormSubmitButton > button {
        background: linear-gradient(90deg, #1976cf, #1598bd) !important;
        color: white !important; min-height: 46px; border-radius: 12px !important;
        border: none !important;
        box-shadow: 0 8px 18px rgba(25, 118, 207, .16);
    }
    [data-testid="stExpander"] {
        background: rgba(255,255,255,.9); border: 1px solid #b9d8ed !important;
        border-radius: 14px !important; overflow: hidden;
    }
    [data-testid="stExpander"] details,
    [data-testid="stExpander"] summary,
    [data-testid="stExpander"] [data-testid="stExpanderDetails"] {
        background: #ffffff !important;
        color: #145e8c !important;
    }
    [data-testid="stExpander"] summary:hover {
        background: #eef8fd !important;
    }
    [data-testid="stExpander"] svg {
        color: #1688b0 !important;
        fill: #1688b0 !important;
    }
    [data-testid="stFileUploaderDropzone"] {
        background: #f1f8fc !important;
        border: 1px dashed #82bddc !important;
        color: #174f73 !important;
    }
    [data-testid="stFileUploaderDropzone"] * {
        color: #174f73 !important;
    }
    [data-testid="stFileUploaderDropzone"] button {
        background: #ffffff !important;
        border: 1px solid #58a9d1 !important;
        color: #176895 !important;
    }
    .preview-heading { color: #145e8c; font-size: 1.45rem; font-weight: 800; margin: .8rem 0; }
    .preview-card {
        background: white; border: 1px solid #bddced; border-left: 5px solid #1999bd;
        border-radius: 14px; padding: 1rem 1.1rem; margin: .6rem 0 1rem;
    }
    .preview-card strong { color: #135d8c; }
    [class*="st-key-document_page_"] {
        background: #ffffff !important;
        border: 1px solid #b8c8d4 !important;
        border-radius: 8px !important;
        padding: 1.25rem 1.4rem 1.4rem !important;
        margin: 1rem 0 1.5rem !important;
        box-shadow: 0 8px 24px rgba(32, 72, 102, .08);
    }
    [class*="st-key-document_page_"] .stTextInput > div > div > input,
    [class*="st-key-document_page_"] .stTextArea textarea {
        border: 1px solid transparent !important;
        border-radius: 4px !important;
        background: transparent !important;
        padding-left: .35rem !important;
        font-size: 1rem !important;
        line-height: 1.65 !important;
    }
    [class*="st-key-document_page_"] .stTextInput > div > div > input:hover,
    [class*="st-key-document_page_"] .stTextInput > div > div > input:focus,
    [class*="st-key-document_page_"] .stTextArea textarea:hover,
    [class*="st-key-document_page_"] .stTextArea textarea:focus {
        border-color: #8cc6e5 !important;
        background: #f8fcff !important;
    }
    [class*="st-key-document_page_"] .stTextInput > label,
    [class*="st-key-document_page_"] .stTextArea > label {
        color: #174f73 !important;
        font-size: .82rem !important;
        text-transform: uppercase;
        letter-spacing: .04em;
    }
    .document-date {
        color: #167aa4; font-weight: 700; font-size: .92rem;
        border-bottom: 1px solid #d6e5ee; padding-bottom: .65rem; margin-bottom: .55rem;
    }
    [data-stale="true"] {
        display: none !important;
    }
    .st-key-download_section {
        margin-top: 2rem !important;
        padding: 1.35rem !important;
        background: #ffffff !important;
        border: 1px solid #b9d8ed !important;
        border-radius: 16px !important;
        box-shadow: 0 10px 28px rgba(32, 72, 102, .08);
    }
    .download-title {
        color: #123f63;
        font-size: 1.55rem;
        font-weight: 800;
        line-height: 1.25;
        margin-bottom: .2rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ALIBABA_THEME_OVERRIDE_START
# This block is intentionally isolated. Revert the dedicated theme commit to
# restore the preserved blue frontend without removing any product features.
st.markdown(
    """
    <style>
    :root, html, body {
        color-scheme: light !important;
        background: #f6f6f6 !important;
    }
    .stApp, [data-testid="stAppViewContainer"] {
        background:
            radial-gradient(circle at 82% 12%, rgba(215, 25, 32, .075), transparent 24rem),
            radial-gradient(circle at 35% 82%, rgba(20, 20, 20, .035), transparent 30rem),
            linear-gradient(145deg, #fffafa 0%, #f5f5f5 48%, #ffffff 100%) !important;
        color: #171717 !important;
    }
    [data-testid="stMainBlockContainer"] {
        padding-top: 1rem !important;
    }
    [data-testid="stSidebar"] {
        background:
            linear-gradient(180deg, #ffffff 0%, #fff8f8 58%, #f5f5f5 100%) !important;
        border-right: 1px solid #dedede !important;
        box-shadow: 10px 0 32px rgba(0, 0, 0, .055);
    }
    [data-testid="stSidebar"] * {
        color: #171717 !important;
    }
    .main-header {
        color: #111111 !important;
        min-height: 132px !important;
        margin: 0 0 1.25rem 0 !important;
        padding: 1.75rem 2.35rem !important;
        border: 1px solid #e1e1e1;
        border-top: 5px solid #d71920;
        border-radius: 0 0 18px 18px !important;
        background:
            radial-gradient(circle at 92% 18%, rgba(215,25,32,.13) 0 74px, transparent 75px),
            linear-gradient(115deg, #ffffff 0%, #fffafa 68%, #f7f7f7 100%) !important;
        box-shadow: 0 14px 34px rgba(0, 0, 0, .08) !important;
        position: relative;
        overflow: hidden;
    }
    .main-header::before {
        content: "";
        position: absolute;
        width: 220px;
        height: 220px;
        right: -92px;
        bottom: -150px;
        border: 24px solid rgba(215, 25, 32, .08);
        border-radius: 50%;
    }
    .main-header::after {
        content: "PLAN • EDIT • DOWNLOAD";
        display: block;
        margin-left: auto;
        color: #d71920;
        font-size: .72rem;
        font-weight: 800;
        letter-spacing: .16em;
    }
    .preview-heading, .download-title {
        color: #171717 !important;
    }
    .preview-heading {
        border-left: 5px solid #d71920;
        padding-left: .8rem;
    }
    .stTextInput input, .stDateInput input, .stTextArea textarea,
    [data-testid="stSidebar"] .stTextInput input,
    [data-testid="stMain"] .stTextInput input,
    [data-testid="stMain"] .stTextArea textarea {
        background: #ffffff !important;
        color: #171717 !important;
        -webkit-text-fill-color: #171717 !important;
        border-color: #cfcfcf !important;
    }
    .stTextInput input:focus, .stDateInput input:focus, .stTextArea textarea:focus,
    [data-testid="stSidebar"] [data-testid="stTextInputRootElement"]:focus-within {
        border-color: #d71920 !important;
        box-shadow: 0 0 0 1px #d71920 !important;
    }
    [data-testid="stSidebar"] .stTextInput input,
    [data-testid="stSidebar"] [data-testid="stDateInput"] > div {
        box-shadow: 0 5px 14px rgba(0, 0, 0, .045) !important;
    }
    div[data-baseweb="input"], div[data-baseweb="base-input"],
    [data-testid="stSidebar"] div[data-baseweb="input"],
    [data-testid="stSidebar"] div[data-baseweb="base-input"] {
        background: #ffffff !important;
        border-color: #cfcfcf !important;
    }
    [data-testid="stDateInput"] div[data-baseweb="input"],
    [data-testid="stDateInput"] div[data-baseweb="base-input"],
    [data-testid="stDateInput"] > div > div,
    [data-testid="stDateInput"] input {
        background: #fff7f7 !important;
        color: #171717 !important;
        -webkit-text-fill-color: #171717 !important;
        border-color: #dfb4b6 !important;
    }
    [data-testid="stDateInput"] svg,
    [data-testid="stDateInput"] button,
    [data-testid="stDateInput"] span {
        color: #d71920 !important;
        fill: #d71920 !important;
    }
    .stButton > button, .stDownloadButton > button,
    [data-testid="stFormSubmitButton"] > button,
    .stFormSubmitButton > button {
        background: linear-gradient(90deg, #b90f17, #e0242b) !important;
        color: #ffffff !important;
        border: 1px solid #b90f17 !important;
        box-shadow: 0 8px 18px rgba(215, 25, 32, .18) !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover,
    [data-testid="stFormSubmitButton"] > button:hover {
        background: #a90c13 !important;
        border-color: #8f090f !important;
    }
    [data-testid="stExpander"] {
        background: #ffffff !important;
        border: 1px solid #d8d8d8 !important;
        box-shadow: 0 7px 20px rgba(0, 0, 0, .04);
    }
    [data-testid="stExpander"] details,
    [data-testid="stExpander"] summary,
    [data-testid="stExpander"] [data-testid="stExpanderDetails"] {
        background: #ffffff !important;
        color: #171717 !important;
    }
    [data-testid="stExpander"] summary:hover {
        background: #fff5f5 !important;
    }
    [data-testid="stExpander"] svg {
        color: #d71920 !important;
        fill: #d71920 !important;
    }
    [data-testid="stFileUploaderDropzone"] {
        background: #fffafa !important;
        border-color: #d8a9ab !important;
    }
    [data-testid="stFileUploaderDropzone"] *,
    [data-testid="stFileUploaderDropzone"] button {
        color: #171717 !important;
    }
    [class*="st-key-document_page_"] {
        background: #ffffff !important;
        border: 1px solid #d9d9d9 !important;
        border-top: 3px solid #d71920 !important;
        box-shadow: 0 14px 34px rgba(0, 0, 0, .075) !important;
    }
    [class*="st-key-document_page_"] .stTextInput > label,
    [class*="st-key-document_page_"] .stTextArea > label,
    .document-date {
        color: #b90f17 !important;
    }
    [class*="st-key-document_page_"] .stTextInput input:hover,
    [class*="st-key-document_page_"] .stTextInput input:focus,
    [class*="st-key-document_page_"] .stTextArea textarea:hover,
    [class*="st-key-document_page_"] .stTextArea textarea:focus {
        border-color: #d71920 !important;
        background: #fffafa !important;
    }
    [class*="st-key-document_page_"] [data-testid="stImage"] img {
        width: 100% !important;
        height: 340px !important;
        max-height: 340px !important;
        object-fit: cover !important;
        border-radius: 12px !important;
    }
    .st-key-cover_image_card [data-testid="stImage"] img {
        width: 100% !important;
        height: 320px !important;
        max-height: 320px !important;
        object-fit: cover !important;
        border-radius: 12px !important;
    }
    .st-key-download_section {
        background:
            linear-gradient(90deg, #d71920 0 6px, transparent 6px),
            #ffffff !important;
        border: 1px solid #dedede !important;
        box-shadow: 0 12px 30px rgba(0, 0, 0, .07) !important;
        padding: 1.4rem 1.5rem 1.5rem 1.75rem !important;
    }
    .st-key-download_section [data-testid="stCaptionContainer"] {
        color: #626262 !important;
    }
    [data-testid="stFeedback"] svg {
        color: #d71920 !important;
        fill: #d71920 !important;
    }
    [data-testid="stProgress"] {
        margin: -.55rem 0 1.25rem 0 !important;
    }
    [data-testid="stProgress"] > div > div > div > div {
        background: linear-gradient(90deg, #b90f17, #ef3038) !important;
    }
    @media (max-width: 720px) {
        .main-header {
            min-height: 112px !important;
            padding: 1.35rem !important;
            font-size: 1.75rem !important;
        }
        .main-header::after { display: none; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)
# ALIBABA_THEME_OVERRIDE_END


st.markdown("<div class='main-header'>Travel Planner</div>", unsafe_allow_html=True)
progress_placeholder = st.empty()

with st.sidebar:
    st.markdown("### Trip Details")
    client_name = st.text_input("Client Name", "", placeholder="Enter client name")
    destination = st.text_input(
        "Destination City / Country",
        "",
        placeholder="New York, USA",
        help="Enter both the destination city and country. Example: New York, USA",
    )
    plan_title = st.text_input("Plan Title", "Travel Plan")
    travel_dates = st.date_input(
        "Travel Dates",
        value=(datetime.now().date(), datetime.now().date() + timedelta(days=5)),
        min_value=datetime.now().date(),
    )
    if isinstance(travel_dates, (tuple, list)) and len(travel_dates) == 2:
        start_date, end_date = travel_dates
    else:
        start_date = end_date = travel_dates[0] if isinstance(travel_dates, (tuple, list)) else travel_dates
    generate = st.button("Create Travel Plan")

if end_date < start_date:
    st.error("Return date cannot be before departure date.")
    st.stop()

if "trip_result" not in st.session_state:
    st.session_state["trip_result"] = None

if generate:
    if not client_name.strip() or not destination.strip():
        st.error("Please enter client name and destination.")
        st.stop()
    progress = progress_placeholder.progress(5, text="Creating your itinerary...")
    itinerary_days = get_ai_plan(client_name, destination, start_date, end_date)
    progress.progress(35, text="Finding verified destination images...")
    cover_path, day_images = build_image_bundle(itinerary_days, destination)
    progress.progress(75, text="Building your PDF...")
    pdf_bytes = build_pdf(
        client=client_name,
        title=plan_title,
        destination=destination,
        start_date=start_date,
        end_date=end_date,
        days=itinerary_days,
        cover_path=cover_path,
        day_images=day_images,
    )
    progress.progress(88, text="Building your Word document...")
    docx_bytes = build_docx(
        client=client_name,
        title=plan_title,
        destination=destination,
        start_date=start_date,
        end_date=end_date,
        days=itinerary_days,
        cover_path=cover_path,
        day_images=day_images,
    )
    progress.progress(98, text="Finalizing your travel plan...")
    progress.progress(100, text="Travel plan ready")
    progress.empty()
    st.session_state["trip_result"] = {
        "client_name": client_name,
        "destination": destination,
        "plan_title": plan_title,
        "start_date": start_date,
        "end_date": end_date,
        "days": itinerary_days,
        "cover_path": cover_path,
        "day_images": day_images,
        "pdf_bytes": pdf_bytes,
        "docx_bytes": docx_bytes,
        "rating": None,
        "ui_version": time.time_ns(),
    }

result = st.session_state.get("trip_result")

if result:
    if repair_non_visitable_places(result):
        st.session_state["trip_result"] = result
    ai_change_notice = st.session_state.pop("ai_change_notice", None)
    if ai_change_notice:
        st.toast(ai_change_notice, icon="✅")
    ui_version = result.get("ui_version", 0)
    st.markdown("<div class='preview-heading'>Travel Plan</div>", unsafe_allow_html=True)

    meta_left, meta_right = st.columns(2)
    with meta_left:
        edited_client = st.text_input("Client Name", result["client_name"], key=f"preview_client_{ui_version}")
        edited_title = st.text_input("Plan Title", result["plan_title"], key=f"preview_title_{ui_version}")
    with meta_right:
        edited_destination = st.text_input(
            "Destination City / Country",
            result["destination"],
            key=f"preview_destination_{ui_version}",
        )
        st.text_input(
            "Travel Dates",
            f"{compact_date(result['start_date'])} - {compact_date(result['end_date'])}",
            disabled=True,
            key=f"preview_dates_{ui_version}",
        )

    cover_upload = None
    if result.get("cover_path"):
        cover_col, _ = st.columns(2)
        with cover_col:
            with st.container(key="cover_image_card"):
                st.image(result["cover_path"], width="stretch")
                with st.expander("Replace Image"):
                    cover_upload = st.file_uploader(
                        "Upload a new cover image",
                        type=["png", "jpg", "jpeg", "webp"],
                        key=f"cover_replacement_{ui_version}",
                        label_visibility="collapsed",
                    )

    edited_days = []
    image_uploads = []
    for idx, day in enumerate(result["days"]):
        with st.container(border=True, key=f"document_page_{idx}"):
            st.markdown(
                f"<div class='document-date'>Day {idx + 1} · {format_day_label(day['date'])}</div>",
                unsafe_allow_html=True,
            )
            editor_value = st_quill(
                value=day.get("editor_html") or day_to_editor_html(day),
                html=True,
                toolbar=[["bold", "italic", "underline"], [{"list": "ordered"}, {"list": "bullet"}], ["clean"]],
                key=f"day_editor_{ui_version}_{idx}",
            )
            current_images = result["day_images"][idx] if idx < len(result["day_images"]) else []
            visible_images = [(original_index, entry) for original_index, entry in enumerate(current_images) if entry.get("path")]
            if visible_images:
                image_columns = st.columns(2)
                for image_idx, (original_index, entry) in enumerate(visible_images):
                    with image_columns[image_idx % len(image_columns)]:
                        st.image(entry["path"], caption=entry.get("place", "Travel image"), width="stretch")
                        with st.expander("Replace Image"):
                            uploaded = st.file_uploader(
                                "Upload a replacement",
                                type=["png", "jpg", "jpeg", "webp"],
                                key=f"replace_image_{ui_version}_{idx}_{original_index}",
                                label_visibility="collapsed",
                            )
                            image_uploads.append((idx, original_index, uploaded))
        edited_days.append(editor_html_to_day(editor_value, day))

    if cover_upload:
        new_cover = save_uploaded_image(cover_upload)
        if new_cover:
            result["cover_path"] = new_cover
            result["ui_version"] = time.time_ns()
            rebuild_result_documents(result)
            st.session_state["trip_result"] = result
            st.rerun()

    for day_index, image_index, uploaded in image_uploads:
        if not uploaded:
            continue
        image_path = save_uploaded_image(uploaded)
        if image_path and day_index < len(result["day_images"]) and image_index < len(result["day_images"][day_index]):
            old_place = result["day_images"][day_index][image_index].get("place", "Travel image")
            result["day_images"][day_index][image_index] = {"place": old_place, "path": image_path}
            result["ui_version"] = time.time_ns()
            rebuild_result_documents(result)
            st.session_state["trip_result"] = result
            st.rerun()

    def day_signature(item: dict):
        return tuple(str(item.get(field, "")) for field in ("title", "morning", "afternoon", "evening")) + (tuple(item.get("places", [])),)

    manual_changed = (
        edited_client.strip() != result["client_name"]
        or edited_title.strip() != result["plan_title"]
        or edited_destination.strip() != result["destination"]
        or any(day_signature(edited) != day_signature(original) for edited, original in zip(edited_days, result["days"]))
    )
    if manual_changed:
        result["client_name"] = edited_client.strip() or result["client_name"]
        result["plan_title"] = edited_title.strip() or result["plan_title"]
        result["destination"] = edited_destination.strip() or result["destination"]
        result["days"] = edited_days
        rebuild_result_documents(result)
        st.session_state["trip_result"] = result
        st.rerun()

    with st.expander("AI Change Assistant", expanded=False):
        prompt_response = st.chat_input(
            placeholder="",
            key=f"travel_plan_change_{ui_version}",
            accept_file=True,
            file_type=["png", "jpg", "jpeg", "webp"],
            max_upload_size=10,
        )

    if prompt_response:
        change_prompt = (
            prompt_response.strip()
            if isinstance(prompt_response, str)
            else (prompt_response.text or "").strip()
        )
        if change_prompt:
            processing_notice = st.info("Applying your requested changes...", icon="⏳")
            reference_payload = None
            uploaded_path = None
            prompt_files = [] if isinstance(prompt_response, str) else list(prompt_response.files or [])
            if prompt_files:
                reference_image = prompt_files[0]
                reference_bytes = reference_image.getvalue()
                reference_payload = {
                    "bytes": reference_bytes,
                    "mime_type": reference_image.type or "image/jpeg",
                }
                uploaded_path = save_pdf_ready_image(reference_bytes)
            image_only_command = is_image_only_instruction(change_prompt)
            previous_places = [list(day.get("places", [])) for day in result["days"]]
            if not image_only_command:
                result["days"] = revise_plan_with_prompt(
                    result["days"], change_prompt, result["destination"], reference_payload
                )
            image_changed = apply_prompt_image_changes(result, change_prompt, uploaded_path)
            if image_only_command and not image_changed:
                processing_notice.empty()
                st.warning("A different relevant image could not be found. Please name the place or attach an image.")
            else:
                revised_places = [list(day.get("places", [])) for day in result["days"]]
                if not image_changed and revised_places != previous_places:
                    result["cover_path"], result["day_images"] = build_image_bundle(result["days"], result["destination"])
                for day in result["days"]:
                    day.pop("editor_html", None)
                result["ui_version"] = time.time_ns()
                rebuild_result_documents(result)
                st.session_state["trip_result"] = result
                st.session_state["ai_change_notice"] = "Requested changes applied successfully."
                processing_notice.empty()
                st.rerun()
        else:
            st.warning("Please enter the change you want.")

    with st.container(border=True, key="download_section"):
        st.markdown("<div class='download-title'>Download Your Travel Plan</div>", unsafe_allow_html=True)
        pdf_col, word_col = st.columns(2)
        with pdf_col:
            st.download_button(
                label="Download PDF",
                data=result["pdf_bytes"],
                file_name=f"{slugify(result['destination'])}_travel_plan.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        with word_col:
            st.download_button(
                label="Download Word",
                data=result["docx_bytes"],
                file_name=f"{slugify(result['destination'])}_travel_plan.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

    st.markdown("---")
    st.caption("Rate this travel plan")
    rating = st.feedback("stars", key="travel_plan_rating")
    if rating is not None:
        result["rating"] = rating + 1
        st.session_state["trip_result"] = result
        st.caption(f"Thank you — {rating + 1}/5")
