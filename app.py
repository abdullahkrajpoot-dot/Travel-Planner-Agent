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
from PIL import Image, ImageDraw, ImageFont, UnidentifiedImageError
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
    header { visibility: hidden !important; height: 0 !important; }
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
    "election", "team", "club", "season",
}


def is_valid_place_name(place_name: str) -> bool:
    tokens = set(slugify(place_name).split("_"))
    return bool(place_name) and not BAD_PLACE_WORDS.intersection(tokens)


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
    return [[
        f"{normalized_destination} historic center",
        f"{normalized_destination} main square",
        f"{normalized_destination} national museum",
        f"{normalized_destination} scenic viewpoint",
    ]]


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
    target_places = max(total_days * 4, 4)

    if len(places) < target_places:
        for item in split_destination_names(destination):
            for place in discover_destination_landmarks(item):
                if len(places) >= target_places:
                    break
                if place not in places and is_valid_place_name(place):
                    places.append(place)

    if len(places) < target_places:
        base = safe_text(destination) or "Destination"
        generic_types = [
            "Old Town",
            "Central Market",
            "Heritage Museum",
            "Main Mosque",
            "City Park",
            "Waterfront",
            "Scenic Viewpoint",
            "Cultural Center",
            "Historic Quarter",
            "National Museum",
            "Main Square",
            "Botanical Garden",
            "Old Bazaar",
            "City Museum",
            "Grand Mosque",
            "Riverside Promenade",
            "Art Gallery",
            "Fortress",
            "Clock Tower",
            "Public Garden",
            "Heritage Street",
            "Traditional Market",
            "Archaeological Site",
            "Panoramic Lookout",
        ]
        for place_type in generic_types:
            if len(places) >= target_places:
                break
            place = f"{base} {place_type}"
            if place not in places:
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
    api_key = os.getenv("OPENROUTER_API_KEY", "").strip()
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
- Do not include events, incidents, attacks, battles, seizures, wars, disasters, hotels, airports, or railway stations as places.
- Keep the itinerary realistic and tourist-friendly.
- Do not use the phrase "Signature Highlight" or "Signature Highlights".
"""

    try:
        response = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={"Authorization": f"Bearer {api_key}"},
            json={
                "model": "anthropic/claude-3-haiku",
                "messages": [{"role": "user", "content": prompt}],
            },
            timeout=90,
        )
        response.raise_for_status()
        data = response.json()
        raw_text = data["choices"][0]["message"]["content"]
        return parse_ai_plan(raw_text, destination, start_date, end_date)
    except Exception:
        return fallback_plan(destination, start_date, end_date)


def revise_plan_with_prompt(days: list[dict], instruction: str, destination: str) -> list[dict]:
    instruction = (instruction or "").strip()
    if not instruction:
        return days
    api_key = os.getenv("OPENROUTER_API_KEY", "").strip()
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
        day_match = re.search(r"day\s*(\d+)", instruction, re.IGNORECASE)
        if day_match:
            target_index = max(int(day_match.group(1)) - 1, 0)
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
            for day in updated[1:-1] if len(updated) > 2 else updated:
                day["evening"] = safe_text(f"{day.get('evening', '')} {instruction}")
        return updated

    prompt = f"""
Revise this travel plan according to the user's instruction.
Return JSON only as {{"days": [...]}} and preserve the same number of days and each date.
Keep every day object fields: title, morning, afternoon, evening, places.
First day remains arrival/rest only and last day remains departure only.
Destination: {destination}
User instruction: {instruction}
Current plan: {json.dumps(days, default=str)}
"""
    try:
        response = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={"Authorization": f"Bearer {api_key}"},
            json={
                "model": "anthropic/claude-3-haiku",
                "messages": [{"role": "user", "content": prompt}],
            },
            timeout=90,
        )
        response.raise_for_status()
        raw_text = response.json()["choices"][0]["message"]["content"]
        payload = json.loads(extract_json_block(raw_text))
        revised = payload.get("days", [])
        if len(revised) != len(days):
            return days
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
                    "places": normalize_place_list(item.get("places", original.get("places", []))),
                }
            )
        return normalized
    except Exception:
        return days


def apply_prompt_image_changes(result: dict, instruction: str) -> bool:
    lowered = (instruction or "").lower()
    if not re.search(r"\b(?:image|images|picture|pictures|photo|photos|pic|pics)\b", lowered):
        return False

    destination = result["destination"]
    replacement_match = re.search(r"(?:to|with)\s+(.+?)\s*$", instruction, re.IGNORECASE)
    replacement_query = safe_text(replacement_match.group(1)) if replacement_match else ""
    replacement_query = re.sub(r"\b(?:image|picture|photo|pic)\b", "", replacement_query, flags=re.IGNORECASE).strip(" -")

    if "cover" in lowered:
        query = replacement_query or destination
        new_path = build_place_image(query, destination)
        if new_path:
            result["cover_path"] = new_path
            return True

    day_match = re.search(r"day\s*(\d+)", lowered)
    if day_match:
        day_index = min(max(int(day_match.group(1)) - 1, 0), len(result["day_images"]) - 1)
        entries = result["day_images"][day_index]
        query = replacement_query or (result["days"][day_index].get("places") or [destination])[0]
        new_path = build_place_image(query, destination)
        if new_path:
            replacement = {"place": query, "path": new_path}
            if entries:
                entries[0] = replacement
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
            new_path = build_place_image(query, destination)
            if new_path:
                result["day_images"][day_index][image_index] = {"place": query, "path": new_path}
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


def create_text_image(place_name: str, destination: str | None = None) -> str | None:
    try:
        image = Image.new("RGB", (1000, 667), (235, 239, 244))
        draw = ImageDraw.Draw(image)
        draw.rectangle((0, 0, 1000, 667), fill=(235, 239, 244))
        draw.rectangle((0, 0, 1000, 134), fill=(22, 48, 86))
        draw.rectangle((58, 183, 942, 583), outline=(150, 163, 184), width=4)
    except Exception:
        return None

    try:
        title_font = ImageFont.truetype("arial.ttf", 58)
        body_font = ImageFont.truetype("arial.ttf", 38)
    except Exception:
        title_font = ImageFont.load_default()
        body_font = ImageFont.load_default()

    try:
        safe_place = safe_text(place_name) or "Travel Highlight"
        safe_destination = safe_text(destination or "")
        draw.text((58, 43), "Travel Highlight", fill=(255, 255, 255), font=title_font)
        wrapped_place = wrap_text(safe_place, 34)
        y = 233
        for line in wrapped_place:
            draw.text((92, y), line, fill=(15, 23, 42), font=body_font)
            y += 54
        if safe_destination:
            draw.text((92, 517), safe_destination, fill=(71, 85, 105), font=body_font)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as temp_file:
            image.save(temp_file.name, format="JPEG", quality=88, optimize=True, progressive=True)
            return temp_file.name
    except Exception:
        return None


def wrap_text(value: str, max_chars: int) -> list[str]:
    words = value.split()
    lines = []
    current = ""
    for word in words:
        next_line = f"{current} {word}".strip()
        if len(next_line) > max_chars and current:
            lines.append(current)
            current = word
        else:
            current = next_line
    if current:
        lines.append(current)
    return lines[:5]


def download_image(url: str, suffix: str = ".jpg"):
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
    if "," in place_name:
        candidates.append(safe_text(place_name.split(",", 1)[0]))
    if destination and destination.lower() not in place_name.lower():
        candidates.append(safe_text(f"{place_name}, {destination}"))
    unique = []
    for item in candidates:
        if item and item not in unique:
            unique.append(item)
    return unique


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

        query_tokens = {
            token
            for token in slugify(query).split("_")
            if len(token) > 2 and token not in {"the", "and", "for"}
        }
        ranked_titles = []
        for row in results:
            title = safe_text(row.get("title", ""))
            title_tokens = {token for token in slugify(title).split("_") if len(token) > 2}
            overlap = len(query_tokens.intersection(title_tokens))
            ranked_titles.append((overlap, title))
        ranked_titles.sort(reverse=True, key=lambda item: item[0])

        for overlap, title in ranked_titles:
            if overlap == 0:
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


def _commons_search_image(query: str) -> str | None:
    """Search Wikimedia Commons for a matching file and return a direct image URL."""
    try:
        search_params = {
            "action": "query",
            "format": "json",
            "list": "search",
            "srsearch": query,
            "srnamespace": 6,  # File namespace on Commons
            "srlimit": 4,
        }
        resp = requests.get("https://commons.wikimedia.org/w/api.php", params=search_params, timeout=20, headers=WIKI_HEADERS)
        resp.raise_for_status()
        data = resp.json()
        results = data.get("query", {}).get("search", [])
        if not results:
            return None

        query_tokens = {
            token
            for token in slugify(query).split("_")
            if len(token) > 2 and token not in {"the", "and", "for", "lahore", "makkah", "madinah"}
        }
        ranked_results = []
        for row in results:
            title = safe_text(row.get("title", ""))
            title_tokens = {token for token in slugify(title).split("_") if len(token) > 2}
            overlap = len(query_tokens.intersection(title_tokens))
            ranked_results.append((overlap, row))
        ranked_results.sort(reverse=True, key=lambda item: item[0])

        for overlap, row in ranked_results:
            if overlap == 0 and query_tokens:
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
                    if url and url.startswith("http"):
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


def build_place_image(place_name: str, destination: str | None = None):
    # Strict place matching first to keep image relevance high.
    queries = _candidate_place_queries(place_name, destination)
    best_url = None

    for query in queries:
        page_url = _wikipedia_page_image(query)
        if page_url:
            best_url = best_url or page_url
            path = download_image(page_url, suffix=".jpg")
            if path:
                return {"place": place_name, "url": page_url, "path": path}

        direct_url = _wikipedia_summary_image(query)
        if direct_url:
            best_url = best_url or direct_url
            path = download_image(direct_url, suffix=".jpg")
            if path:
                return {"place": place_name, "url": direct_url, "path": path}

        searched_url = _wikipedia_search_image(query)
        if not searched_url:
            # Try Wikimedia Commons as an additional fallback for image search
            commons_url = _commons_search_image(query)
            if commons_url:
                best_url = best_url or commons_url
                path = download_image(commons_url, suffix=".jpg")
                if path:
                    return {"place": place_name, "url": commons_url, "path": path}
            continue
        best_url = best_url or searched_url
        path = download_image(searched_url, suffix=".jpg")
        if path:
            return {"place": place_name, "url": searched_url, "path": path}

    for query in queries:
        for image_url in _web_image_search_urls(query):
            best_url = best_url or image_url
            path = download_image(image_url, suffix=".jpg")
            if path:
                return {"place": place_name, "url": image_url, "path": path}

    fallback_query = requests.utils.quote(f"{place_name} {destination or ''} landmark travel photo".strip())
    fallback_urls = []
    if best_url:
        fallback_urls.append(best_url)
    fallback_urls.extend([
        f"https://source.unsplash.com/1200x800/?{fallback_query}",
        f"https://loremflickr.com/1200/800/{fallback_query}",
    ])

    for fallback_url in fallback_urls:
        path = download_image(fallback_url, suffix=".jpg")
        if path:
            return {"place": place_name, "url": fallback_url, "path": path}

    placeholder_path = create_text_image(place_name, destination)
    return {"place": place_name, "url": fallback_urls[0] if fallback_urls else None, "path": placeholder_path}


def build_image_bundle(days: list[dict], destination: str):
    cover_candidates = []
    for day in days[:2]:
        cover_candidates.extend(day.get("places", []))
    if not cover_candidates:
        cover_candidates = [destination]

    unique_places = []
    planned_places = [place for day in days for place in day.get("places", [])]
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
    used_places: set[str] = set()
    for place in cover_candidates:
        entry = results.get(place, {})
        path = entry.get("path")
        if path:
            cover_path = path
            break

    day_images: list[list[dict]] = []
    for day in days:
        entries: list[dict] = []
        for place in day.get("places", []):
            entry = results.get(place, {"place": place, "url": None, "path": None})
            path = entry.get("path")
            source = path or entry.get("url")
            if source and source not in used_sources and place not in used_places:
                entries.append(entry)
                used_sources.add(source)
                used_places.add(place)

        day_images.append(entries[:4])

    return cover_path, day_images


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
    output = pdf.output(dest="S")
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
    .stButton > button, .stDownloadButton > button {
        background: linear-gradient(90deg, #1976cf, #1598bd) !important;
        color: white !important; min-height: 46px; border-radius: 12px !important;
        box-shadow: 0 8px 18px rgba(25, 118, 207, .16);
    }
    [data-testid="stExpander"] {
        background: rgba(255,255,255,.9); border: 1px solid #b9d8ed !important;
        border-radius: 14px !important; overflow: hidden;
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
    </style>
    """,
    unsafe_allow_html=True,
)


st.markdown("<div class='main-header'>Travel Planner</div>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### Trip Details")
    client_name = st.text_input("Client Name", "", placeholder="Enter client name")
    destination = st.text_input(
        "Destination",
        "",
        placeholder="New York, Paris",
        help="For multiple cities, separate them with a comma. Example: New York, Paris",
    )
    plan_title = st.text_input("Plan Title", "Travel Plan")
    travel_dates = st.date_input(
        "Travel Dates",
        value=(datetime.now().date(), datetime.now().date() + timedelta(days=5)),
        help="Select the departure date first, then the return date.",
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
    progress = st.progress(0)
    itinerary_days = get_ai_plan(client_name, destination, start_date, end_date)
    progress.progress(35)
    cover_path, day_images = build_image_bundle(itinerary_days, destination)
    progress.progress(80)
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
    progress.progress(100)
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
    }

result = st.session_state.get("trip_result")

if result:
    st.markdown("### AI Change Assistant")
    st.caption("Tell the assistant what to change in the text or images. PDF and Word files update automatically.")
    with st.form("prompt_change_form"):
        change_prompt = st.text_area(
            "What would you like to change?",
            placeholder=(
                "Examples: First paragraph heading change to Welcome to New York | "
                "Day 2 morning change to a relaxed city tour | "
                "First picture replace with Eiffel Tower | Cover image change to Manhattan skyline"
            ),
            height=95,
        )
        apply_prompt = st.form_submit_button("Apply AI Changes")
    if apply_prompt:
        if change_prompt.strip():
            revised_days = revise_plan_with_prompt(result["days"], change_prompt, result["destination"])
            result["days"] = revised_days
            image_changed = apply_prompt_image_changes(result, change_prompt)
            for day in result["days"]:
                day.pop("editor_html", None)
            rebuild_result_documents(result)
            st.session_state["trip_result"] = result
            if image_changed:
                st.success("AI changes applied to the travel plan and images.")
            else:
                st.success("AI changes applied to the travel plan.")
            st.rerun()
        else:
            st.warning("Please enter the change you want.")

    download_tab, preview_tab = st.tabs(["Download", "Document Preview"])

    with download_tab:
        pdf_col, word_col = st.columns(2)
        with pdf_col:
            st.download_button(
                label="Download PDF",
                data=result["pdf_bytes"],
                file_name=f"{slugify(result['destination'])}_travel_plan.pdf",
                mime="application/pdf",
            )
        with word_col:
            st.download_button(
                label="Download Word",
                data=result["docx_bytes"],
                file_name=f"{slugify(result['destination'])}_travel_plan.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    with preview_tab:
        st.markdown("<div class='preview-heading'>Travel Plan</div>", unsafe_allow_html=True)
        st.caption("Click directly inside the document text to edit it.")
        with st.form("editable_travel_plan"):
            meta_left, meta_right = st.columns(2)
            with meta_left:
                edited_client = st.text_input("Client Name", result["client_name"])
                edited_title = st.text_input("Plan Title", result["plan_title"])
            with meta_right:
                edited_destination = st.text_input("Destination", result["destination"])
                st.text_input(
                    "Travel Dates",
                    f"{compact_date(result['start_date'])} - {compact_date(result['end_date'])}",
                    disabled=True,
                )

            if result.get("cover_path"):
                st.image(result["cover_path"], width=520)

            with st.expander("Change cover image (optional)"):
                cover_upload = st.file_uploader(
                    "Cover image",
                    type=["png", "jpg", "jpeg", "webp"],
                    key="cover_replacement",
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
                        key=f"day_editor_{idx}",
                    )
                    current_images = result["day_images"][idx] if idx < len(result["day_images"]) else []
                    visible_images = [entry for entry in current_images if entry.get("path")]
                    if visible_images:
                        image_columns = st.columns(min(len(visible_images), 2))
                        for image_idx, entry in enumerate(visible_images):
                            with image_columns[image_idx % len(image_columns)]:
                                st.image(entry["path"], caption=entry.get("place", "Travel image"), use_container_width=True)
                    with st.expander("Replace images"):
                        uploads = st.file_uploader(
                            "Upload replacement images",
                            type=["png", "jpg", "jpeg", "webp"],
                            accept_multiple_files=True,
                            key=f"day_images_{idx}",
                            label_visibility="collapsed",
                        )
                edited_days.append(editor_html_to_day(editor_value, day))
                image_uploads.append(uploads)

            apply_edits = st.form_submit_button("Save Changes")

        if apply_edits:
            result["client_name"] = edited_client.strip() or result["client_name"]
            result["plan_title"] = edited_title.strip() or result["plan_title"]
            result["destination"] = edited_destination.strip() or result["destination"]
            result["days"] = edited_days
            new_cover = save_uploaded_image(cover_upload)
            if new_cover:
                result["cover_path"] = new_cover
            for idx, uploads in enumerate(image_uploads):
                if not uploads:
                    continue
                replacements = []
                for image_idx, uploaded in enumerate(uploads[:4]):
                    image_path = save_uploaded_image(uploaded)
                    if image_path:
                        places = result["days"][idx].get("places", [])
                        label = places[image_idx] if image_idx < len(places) else f"Day {idx + 1} image"
                        replacements.append({"place": label, "path": image_path})
                if replacements:
                    result["day_images"][idx] = replacements
            rebuild_result_documents(result)
            st.session_state["trip_result"] = result
            st.success("Travel plan updated.")
            st.rerun()

    st.markdown("---")
    st.caption("Rate this travel plan")
    rating = st.feedback("stars", key="travel_plan_rating")
    if rating is not None:
        result["rating"] = rating + 1
        st.session_state["trip_result"] = result
        st.caption(f"Thank you — {rating + 1}/5")
