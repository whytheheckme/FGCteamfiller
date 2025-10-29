"""FIRST Global 2025 Team Video Slotter GUI.

This module launches a simple Tkinter application with two tabs for
future expansion.
"""

from __future__ import annotations

import importlib.util

if importlib.util.find_spec("tkinter") is None:  # pragma: no cover - import-time guard
    raise ModuleNotFoundError(
        "tkinter is not available in this Python installation. "
        "On macOS with Homebrew Python, install it with 'brew install python-tk@3.13'."
    )

import difflib
import itertools
import json
import math
import re
import threading
import tkinter as tk
import unicodedata
from dataclasses import dataclass
from enum import IntEnum
from datetime import datetime, timezone
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, Dict, Iterable, List, Mapping, Optional, Sequence, Set, Tuple

try:  # pragma: no cover - optional dependency
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
except ModuleNotFoundError:  # pragma: no cover - handled at runtime
    Credentials = None  # type: ignore[assignment]
    InstalledAppFlow = None  # type: ignore[assignment]

try:  # pragma: no cover - optional dependency
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except ModuleNotFoundError:  # pragma: no cover - handled at runtime
    build = None  # type: ignore[assignment]
    HttpError = None  # type: ignore[assignment]

try:  # pragma: no cover - optional dependency
    from google.auth.transport.requests import Request
except ModuleNotFoundError:  # pragma: no cover - handled at runtime
    Request = None  # type: ignore[assignment]


FLAG_EMOJI_PATTERN = re.compile(r"^\s*([\U0001F1E6-\U0001F1FF]{2})")
COUNTRY_NAME_LEADING_NOISE_PATTERN = re.compile(
    r"^(?:(?:first\s+global|fgc)\s+)?(?:team|delegation)\b[\s:-]*",
    re.IGNORECASE,
)
COUNTRY_NAME_TRAILING_NOISE_PATTERN = re.compile(
    r"[\s:-]*(?:team|delegation)\b$",
    re.IGNORECASE,
)
ISO_CODE_IN_TEXT_PATTERN = re.compile(r"\b([A-Z]{3})\b")
PLACEHOLDER_CODE_PATTERN = re.compile(
    r"^TEAM VIDEO PLACEHOLDER ([A-Z]{3})\b",
    re.IGNORECASE,
)
RANKING_MATCH_NUMBER_PATTERN = re.compile(r"RANKING MATCH\s*#?\s*(\d+)", re.IGNORECASE)
GOOGLE_DOCUMENT_ID_PATTERN = re.compile(
    r"/document/(?:u/\d+/)?d/([a-zA-Z0-9_-]+)", re.IGNORECASE
)
BLOCK_START_PATTERN = re.compile(r"\[Block\s+(\d+)\s+start\]", re.IGNORECASE)
BLOCK_END_PATTERN = re.compile(r"\[Block\s+(\d+)\s+end\]", re.IGNORECASE)


def _normalize_header_alias(text: str, *, uppercase: bool) -> str:
    normalized = unicodedata.normalize("NFKC", text)
    normalized = re.sub(r"\s+", " ", normalized).strip()
    if uppercase:
        return normalized.upper()
    return normalized.lower()


def _normalize_alias_set(aliases: Iterable[str], *, uppercase: bool) -> Set[str]:
    return {_normalize_header_alias(alias, uppercase=uppercase) for alias in aliases}


VIDEO_NUMBER_HEADER_ALIASES = _normalize_alias_set(
    [
        "VIDEO #",
        "VIDEO NO",
        "VIDEO N°",
        "VIDEO Nº",
        "VIDEO NUMBER",
        "VIDEO NUM",
        "VIDEO N",
    ],
    uppercase=True,
)


DURATION_HEADER_ALIASES = _normalize_alias_set(
    [
        "DURATION",
        "DURACION",
        "DURACIÓN",
        "LENGTH",
    ],
    uppercase=True,
)


VIDEO_SHEET_HEADER_ALIASES: Dict[str, Set[str]] = {
    "team": _normalize_alias_set(["team"], uppercase=False),
    "value": _normalize_alias_set(["value", "score"], uppercase=False),
    "video_id": _normalize_alias_set(["video id", "id", "video"], uppercase=False),
    "time": _normalize_alias_set(["time", "start time", "scheduled"], uppercase=False),
    "video_number": _normalize_alias_set(
        [
            "video #",
            "video no",
            "video nº",
            "video n°",
            "video number",
            "video num",
        ],
        uppercase=False,
    ),
    "duration": _normalize_alias_set(["duration", "duración", "length"], uppercase=False),
    "match": _normalize_alias_set(
        [
            "match",
            "match #",
            "match number",
            "match no",
            "match nº",
            "match n°",
        ],
        uppercase=False,
    ),
}


def strip_leading_flag_emoji(text: str) -> str:
    """Remove a leading emoji flag from *text*, if present."""

    match = FLAG_EMOJI_PATTERN.match(text)
    if not match:
        return text.strip()
    start, end = match.span(1)
    return (text[:start] + text[end:]).strip()


def _strip_country_name_noise(name: str) -> str:
    value = name.strip()
    if not value:
        return value

    # Remove repeated leading noise phrases such as "FGC Team" or "Delegation".
    while True:
        cleaned = COUNTRY_NAME_LEADING_NOISE_PATTERN.sub("", value)
        if cleaned == value:
            break
        value = cleaned.strip(" -–—,:")

    value = re.sub(r"(?i)^of\b[\s:-]*", "", value).strip(" -–—,:")
    value = COUNTRY_NAME_TRAILING_NOISE_PATTERN.sub("", value).strip(" -–—,:")
    return value


def _extract_numeric_cell_value(cell: Mapping[str, Any]) -> Optional[float]:
    for key in ("effectiveValue", "userEnteredValue"):
        value = cell.get(key)
        if isinstance(value, Mapping):
            number = value.get("numberValue")
            if isinstance(number, (int, float)):
                return float(number)
    text = _extract_cell_text(dict(cell)) if isinstance(cell, dict) else ""
    if text:
        normalized = text.replace(",", "").strip()
        try:
            return float(normalized)
        except ValueError:
            return None
    return None


COUNTRY_CODE_DATA = """
AFG,AF,Afghanistan
ALB,AL,Albania
DZA,DZ,Algeria
AND,AD,Andorra
AGO,AO,Angola
AIA,AI,Anguilla
ATG,AG,Antigua and Barbuda
ARG,AR,Argentina
ARM,AM,Armenia
ABW,AW,Aruba
AUS,AU,Australia
AUT,AT,Austria
AZE,AZ,Azerbaijan
BHS,BS,Bahamas
BHR,BH,Bahrain
BGD,BD,Bangladesh
BRB,BB,Barbados
BLR,BY,Belarus
BEL,BE,Belgium
BLZ,BZ,Belize
BEN,BJ,Benin
BMU,BM,Bermuda
BTN,BT,Bhutan
BOL,BO,Bolivia (Plurinational State of)
BIH,BA,Bosnia and Herzegovina
BWA,BW,Botswana
BRA,BR,Brazil
BRN,BN,Brunei Darussalam
BGR,BG,Bulgaria
BFA,BF,Burkina Faso
BDI,BI,Burundi
KHM,KH,Cambodia
CMR,CM,Cameroon
CAN,CA,Canada
CPV,CV,Cabo Verde
CYM,KY,Cayman Islands
CAF,CF,Central African Republic
TCD,TD,Chad
CHL,CL,Chile
CHN,CN,China
COL,CO,Colombia
COM,KM,Comoros
COG,CG,Congo
COD,CD,Democratic Republic of the Congo
COK,CK,Cook Islands
CRI,CR,Costa Rica
CIV,CI,Côte d’Ivoire
HRV,HR,Croatia
CUB,CU,Cuba
CUW,CW,Curaçao
CYP,CY,Cyprus
CZE,CZ,Czechia
DNK,DK,Denmark
DJI,DJ,Djibouti
DMA,DM,Dominica
DOM,DO,Dominican Republic
ECU,EC,Ecuador
EGY,EG,Egypt
SLV,SV,El Salvador
GNQ,GQ,Equatorial Guinea
ERI,ER,Eritrea
EST,EE,Estonia
SWZ,SZ,Eswatini
ETH,ET,Ethiopia
FJI,FJ,Fiji
FIN,FI,Finland
FRA,FR,France
GUF,GF,French Guiana
PYF,PF,French Polynesia
GAB,GA,Gabon
GMB,GM,Gambia
GEO,GE,Georgia
DEU,DE,Germany
GHA,GH,Ghana
GRC,GR,Greece
GRD,GD,Grenada
GLP,GP,Guadeloupe
GTM,GT,Guatemala
GGY,GG,Guernsey
GIN,GN,Guinea
GNB,GW,Guinea-Bissau
GUY,GY,Guyana
HTI,HT,Haiti
HND,HN,Honduras
HKG,HK,Hong Kong
HUN,HU,Hungary
ISL,IS,Iceland
IND,IN,India
IDN,ID,Indonesia
IRN,IR,Iran (Islamic Republic of)
IRQ,IQ,Iraq
IRL,IE,Ireland
IMN,IM,Isle of Man
ISR,IL,Israel
ITA,IT,Italy
JAM,JM,Jamaica
JPN,JP,Japan
JEY,JE,Jersey
JOR,JO,Jordan
KAZ,KZ,Kazakhstan
KEN,KE,Kenya
KIR,KI,Kiribati
PRK,KP,Democratic People's Republic of Korea
KOR,KR,Republic of Korea
KWT,KW,Kuwait
KGZ,KG,Kyrgyzstan
LAO,LA,Lao People's Democratic Republic
LVA,LV,Latvia
LBN,LB,Lebanon
LSO,LS,Lesotho
LBR,LR,Liberia
LBY,LY,Libya
LIE,LI,Liechtenstein
LTU,LT,Lithuania
LUX,LU,Luxembourg
MAC,MO,Macao
MDG,MG,Madagascar
MWI,MW,Malawi
MYS,MY,Malaysia
MDV,MV,Maldives
MLI,ML,Mali
MLT,MT,Malta
MHL,MH,Marshall Islands
MTQ,MQ,Martinique
MRT,MR,Mauritania
MUS,MU,Mauritius
MYT,YT,Mayotte
MEX,MX,Mexico
FSM,FM,Micronesia (Federated States of)
MDA,MD,Republic of Moldova
MCO,MC,Monaco
MNG,MN,Mongolia
MNE,ME,Montenegro
MSR,MS,Montserrat
MAR,MA,Morocco
MOZ,MZ,Mozambique
MMR,MM,Myanmar
NAM,NA,Namibia
NRU,NR,Nauru
NPL,NP,Nepal
NLD,NL,Netherlands
NCL,NC,New Caledonia
NZL,NZ,New Zealand
NIC,NI,Nicaragua
NER,NE,Niger
NGA,NG,Nigeria
NIU,NU,Niue
MKD,MK,North Macedonia
NOR,NO,Norway
OMN,OM,Oman
PAK,PK,Pakistan
PLW,PW,Palau
PSE,PS,State of Palestine
PAN,PA,Panama
PNG,PG,Papua New Guinea
PRY,PY,Paraguay
PER,PE,Peru
PHL,PH,Philippines
POL,PL,Poland
PRT,PT,Portugal
PRI,PR,Puerto Rico
QAT,QA,Qatar
REU,RE,Réunion
ROU,RO,Romania
RUS,RU,Russian Federation
RWA,RW,Rwanda
BLM,BL,Saint Barthélemy
SHN,SH,Saint Helena, Ascension and Tristan da Cunha
KNA,KN,Saint Kitts and Nevis
LCA,LC,Saint Lucia
MAF,MF,Saint Martin (French part)
SPM,PM,Saint Pierre and Miquelon
VCT,VC,Saint Vincent and the Grenadines
WSM,WS,Samoa
SMR,SM,San Marino
STP,ST,Sao Tome and Principe
SAU,SA,Saudi Arabia
SEN,SN,Senegal
SRB,RS,Serbia
SYC,SC,Seychelles
SLE,SL,Sierra Leone
SGP,SG,Singapore
SXM,SX,Sint Maarten (Dutch part)
SVK,SK,Slovakia
SVN,SI,Slovenia
SLB,SB,Solomon Islands
SOM,SO,Somalia
ZAF,ZA,South Africa
SSD,SS,South Sudan
ESP,ES,Spain
LKA,LK,Sri Lanka
SDN,SD,Sudan
SUR,SR,Suriname
SWE,SE,Sweden
CHE,CH,Switzerland
SYR,SY,Syrian Arab Republic
TWN,TW,Taiwan, Province of China
TJK,TJ,Tajikistan
TZA,TZ,United Republic of Tanzania
THA,TH,Thailand
TLS,TL,Timor-Leste
TGO,TG,Togo
TON,TO,Tonga
TTO,TT,Trinidad and Tobago
TUN,TN,Tunisia
TUR,TR,Turkey
TKM,TM,Turkmenistan
TCA,TC,Turks and Caicos Islands
TUV,TV,Tuvalu
UGA,UG,Uganda
UKR,UA,Ukraine
ARE,AE,United Arab Emirates
GBR,GB,United Kingdom of Great Britain and Northern Ireland
USA,US,United States of America
URY,UY,Uruguay
UZB,UZ,Uzbekistan
VUT,VU,Vanuatu
VAT,VA,Vatican City State
VEN,VE,Bolivarian Republic of Venezuela
VNM,VN,Viet Nam
WLF,WF,Wallis and Futuna
ESH,EH,Western Sahara
YEM,YE,Yemen
ZMB,ZM,Zambia
ZWE,ZW,Zimbabwe
XKX,XK,Kosovo
"""


ADDITIONAL_COUNTRY_ALIASES: Dict[str, Tuple[str, ...]] = {
    "ARE": ("UAE", "United Arab Emirates"),
    "BOL": ("Bolivia",),
    "BRN": ("Brunei",),
    "CIV": ("Ivory Coast", "Cote d'Ivoire"),
    "COD": ("DR Congo", "Democratic Republic of Congo", "Congo DR"),
    "COG": ("Republic of the Congo", "Congo"),
    "CPV": ("Cape Verde",),
    "CZE": ("Czech Republic",),
    "FSM": ("Micronesia",),
    "GBR": ("United Kingdom", "Great Britain", "UK"),
    "IRN": ("Iran",),
    "KOR": ("South Korea", "Republic of Korea"),
    "LAO": ("Laos",),
    "MAC": ("Macau",),
    "MDA": ("Moldova",),
    "MKD": ("Macedonia", "North Macedonia"),
    "MMR": ("Burma", "Myanmar"),
    "PRK": ("North Korea",),
    "RUS": ("Russia",),
    "SRB": ("Serbia",),
    "SWZ": ("Swaziland",),
    "SYR": ("Syria",),
    "TJK": ("Tadjikistan",),
    "TLS": ("East Timor",),
    "TTO": ("Trinidad & Tobago", "Trinidad and Tobago"),
    "TWN": ("Taiwan",),
    "TZA": ("Tanzania",),
    "UKR": ("Ukraine",),
    "USA": ("United States", "USA", "United States of America"),
    "VEN": ("Venezuela",),
    "VNM": ("Vietnam",),
    "XKX": ("Kosovo",),
    "TUR": ("Türkiye", "Republic of Türkiye"),
}


COUNTRY_KEYWORD_MATCHES: Tuple[Tuple[str, str], ...] = (
    ("hongkong", "HKG"),
    ("palestin", "PSE"),
    ("moldov", "MDA"),
    ("micrones", "FSM"),
    ("iran", "IRN"),
    ("china", "CHN"),
    ("turkiye", "TUR"),
)


def _normalize_country_name(name: str) -> str:
    normalized = unicodedata.normalize("NFKD", name)
    normalized = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    normalized = normalized.replace("’", "'")
    normalized = normalized.lower()
    normalized = re.sub(r"[^a-z0-9]+", "", normalized)
    return normalized


def _generate_country_name_variants(name: str) -> Set[str]:
    variants: Set[str] = set()
    stripped = name.strip()
    if not stripped:
        return variants
    variants.add(stripped)
    variants.add(stripped.replace("’", "'"))

    if "(" in stripped and ")" in stripped:
        without_parens = re.sub(r"\s*\(.*?\)", "", stripped).strip()
        if without_parens:
            variants.add(without_parens)

    if "," in stripped:
        parts = [part.strip() for part in stripped.split(",") if part.strip()]
        if len(parts) == 2:
            variants.add(" ".join(reversed(parts)))
        variants.update(parts)

    if " and " in stripped.lower():
        variants.add(stripped.replace(" and ", " & "))

    variants.add(stripped.upper())

    return {variant for variant in variants if variant}


COUNTRY_CODE_TO_INFO: Dict[str, Tuple[str, str]] = {}
ISO2_TO_ISO3: Dict[str, str] = {}
COUNTRY_NAME_TO_CODE: Dict[str, str] = {}

for line in COUNTRY_CODE_DATA.strip().splitlines():
    iso3, iso2, name = [part.strip() for part in line.split(",", 2)]
    iso3 = iso3.upper()
    iso2 = iso2.upper()
    COUNTRY_CODE_TO_INFO[iso3] = (name, iso2)
    if len(iso2) == 2 and iso2.isalpha() and iso2 not in ISO2_TO_ISO3:
        ISO2_TO_ISO3[iso2] = iso3

    variants = _generate_country_name_variants(name)
    variants.update(ADDITIONAL_COUNTRY_ALIASES.get(iso3, ()))
    variants.add(iso3)
    variants.add(iso2)

    for variant in variants:
        normalized = _normalize_country_name(variant)
        if normalized and normalized not in COUNTRY_NAME_TO_CODE:
            COUNTRY_NAME_TO_CODE[normalized] = iso3


def normalize_country_code(code: str) -> Optional[str]:
    """Return the ISO-3166 alpha-3 code for *code* if it can be resolved."""

    if not code:
        return None
    candidate = code.strip().upper()
    if not candidate:
        return None
    if re.fullmatch(r"[A-Z]{3}", candidate) and candidate in COUNTRY_CODE_TO_INFO:
        return candidate
    if re.fullmatch(r"[A-Z]{2}", candidate):
        mapped = ISO2_TO_ISO3.get(candidate)
        if mapped:
            return mapped
    return None


def normalize_country_lookup_value(name: str) -> str:
    cleaned = strip_leading_flag_emoji(name)
    cleaned = _strip_country_name_noise(cleaned)
    return _normalize_country_name(cleaned)


def lookup_country_code(name: str) -> Optional[str]:
    normalized = normalize_country_lookup_value(name)
    if not normalized:
        return None
    code = COUNTRY_NAME_TO_CODE.get(normalized)
    if code:
        return code
    for keyword, keyword_code in COUNTRY_KEYWORD_MATCHES:
        if keyword in normalized:
            return keyword_code
    return None


def get_country_display_name(code: str) -> Optional[str]:
    info = COUNTRY_CODE_TO_INFO.get(code.upper())
    if not info:
        return None
    return info[0]


def country_code_to_flag(code: str) -> str:
    info = COUNTRY_CODE_TO_INFO.get(code.upper())
    if not info:
        return ""
    iso2 = info[1]
    if len(iso2) != 2 or not iso2.isalpha():
        return ""
    try:
        return "".join(chr(0x1F1E6 + ord(char.upper()) - ord("A")) for char in iso2)
    except ValueError:
        return ""


def _format_country_codes_for_log(codes: Sequence[str]) -> str:
    if not codes:
        return "(none)"
    formatted: List[str] = []
    for code in codes:
        normalized = normalize_country_code(code) or code.strip().upper()
        display = get_country_display_name(normalized)
        if display:
            formatted.append(f"{normalized} ({display})")
        elif normalized:
            formatted.append(normalized)
    return ", ".join(formatted) if formatted else "(none)"


def _extract_match_number_from_text(text: str) -> Optional[int]:
    match = RANKING_MATCH_NUMBER_PATTERN.search(text)
    if not match:
        return None
    try:
        return int(match.group(1))
    except ValueError:
        return None


def _normalize_video_number(value: str) -> Optional[str]:
    digits = re.findall(r"\d+", value)
    if not digits:
        return None
    number = digits[-1].lstrip("0")
    return number or (digits[-1] if digits[-1] else None)


def _normalize_booth_key(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", value or "")
    normalized = normalized.upper()
    normalized = re.sub(r"[^A-Z0-9]+", " ", normalized)
    return re.sub(r"\s+", " ", normalized).strip()


def _match_booth_interview(key: str, rows: Sequence[BoothInterviewRow]) -> Optional[BoothInterviewRow]:
    normalized_key = _normalize_booth_key(key)
    if not normalized_key:
        return None

    for row in rows:
        if row.normalized_key == normalized_key:
            return row

    best_row: Optional[BoothInterviewRow] = None
    best_score = 0.0
    for row in rows:
        score = difflib.SequenceMatcher(None, normalized_key, row.normalized_key).ratio()
        if score > best_score:
            best_score = score
            best_row = row

    if best_row and best_score >= 0.7:
        return best_row
    return None


def _extract_host_number(value: str) -> Optional[str]:
    if not value:
        return None
    match = re.search(r"(\d+)", value)
    if match:
        number = match.group(1).lstrip("0")
        return number or match.group(1)
    return None


def find_host_column(sheet_data: Sequence[dict], task_column: int) -> Optional[int]:
    rows = _collect_sheet_rows(sheet_data)
    header_row_index: Optional[int] = None
    for row_index, columns in rows:
        cell = columns.get(task_column)
        if not cell:
            continue
        text = _extract_cell_text(cell)
        if text.strip().upper() == "TASK":
            header_row_index = row_index
            break

    search_limit = header_row_index + 3 if header_row_index is not None else 5
    for row_index, columns in rows:
        if row_index > search_limit:
            break
        for column_index, cell in columns.items():
            if column_index <= task_column:
                continue
            text = _extract_cell_text(cell)
            if not text:
                continue
            normalized = _normalize_header_alias(text, uppercase=True)
            if "HOST" in normalized or "TALENT" in normalized:
                return column_index

    return None


@dataclass
class VideoEntry:
    team_name: str
    normalized_name: str
    value: Optional[float]
    video_id: Optional[str]
    time: Optional[str]
    row_index: int
    video_number: Optional[str]
    duration: Optional[str]
    team_cell_text: Optional[str] = None
    video_number_cell_text: Optional[str] = None
    duration_cell_text: Optional[str] = None


@dataclass
class VideoDataset:
    entries: List[VideoEntry]
    by_code: Dict[str, VideoEntry]
    by_normalized_name: Dict[str, VideoEntry]
    sheet_title: str = "Videos"
    video_number_column: Optional[int] = None
    duration_column: Optional[int] = None
    match_column: Optional[int] = None
    team_column: Optional[int] = None


@dataclass
class PlaceholderSlot:
    sheet_title: str
    row_index: int
    task_column: int
    match_number: int
    placeholder_index: int
    video_number_column: Optional[int] = None
    duration_column: Optional[int] = None


@dataclass
class PlaceholderAssignment:
    slot: PlaceholderSlot
    country_code: str
    video: VideoEntry
    dataset: VideoDataset


@dataclass
class ScriptLine:
    text: str
    bold: bool = False
    alignment: str = "left"


@dataclass
class VideoScriptRow:
    number: str
    label: str
    script: str


@dataclass
class BoothInterviewRow:
    key: str
    script: str
    normalized_key: str



def _get_widget_background(widget: tk.Misc, fallback: tk.Misc) -> str:
    """Return a usable background color for a widget.

    ttk widgets such as ``ttk.Frame`` don't expose a ``-background`` option,
    which results in a ``TclError`` when ``cget("background")`` is called.
    This helper gracefully handles that case by falling back to the top-level
    window's background or, ultimately, the fallback widget's current
    background color.
    """

    for candidate in (widget, widget.winfo_toplevel(), fallback):
        try:
            background = str(candidate.cget("background"))
        except tk.TclError:
            continue
        if background:
            return background
    return str(fallback.cget("background"))


class ApplicationConsole:
    """Display log messages emitted by tools across the application."""

    class LogLevel(IntEnum):
        INFO = 10
        WARN = 20

    def __init__(self, parent: tk.Misc) -> None:
        self.parent = parent
        self._text_widget: Optional[tk.Text] = None
        self._level: "ApplicationConsole.LogLevel" = self.LogLevel.WARN

    def render(self, row: int) -> None:
        if hasattr(self.parent, "columnconfigure"):
            try:
                self.parent.columnconfigure(0, weight=1)
            except tk.TclError:
                pass
        if hasattr(self.parent, "rowconfigure"):
            try:
                self.parent.rowconfigure(row, weight=1)
            except tk.TclError:
                pass

        frame = ttk.Frame(self.parent, padding=(16, 0, 16, 16))
        frame.grid(row=row, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text="Console", font=("Helvetica", 12, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )

        text_widget = tk.Text(
            frame,
            height=8,
            wrap="word",
            borderwidth=1,
            relief="solid",
            cursor="xterm",
        )
        background = _get_widget_background(frame, text_widget)
        text_widget.configure(background=background, state="disabled")
        text_widget.grid(row=1, column=0, sticky="nsew")

        def _on_key(event: tk.Event) -> str | None:  # type: ignore[name-defined]
            control_pressed = bool(event.state & 0x4)
            command_pressed = bool(event.state & 0x200000)
            if control_pressed or command_pressed:
                if event.keysym.lower() in {"a", "c"}:
                    return None
                return "break"
            if event.keysym in {
                "Left",
                "Right",
                "Up",
                "Down",
                "Home",
                "End",
                "Next",
                "Prior",
                "Shift_L",
                "Shift_R",
                "Control_L",
                "Control_R",
                "Meta_L",
                "Meta_R",
            }:
                return None
            if event.keysym in {"BackSpace", "Delete", "Return", "KP_Enter"}:
                return "break"
            if event.char and event.char.isprintable():
                return "break"
            return None

        text_widget.bind("<Key>", _on_key)
        frame.rowconfigure(1, weight=1)
        self._text_widget = text_widget

    def get_level(self) -> "ApplicationConsole.LogLevel":
        return self._level

    def set_level(self, level: "ApplicationConsole.LogLevel") -> None:
        self._level = level

    def log(
        self,
        message: str,
        *,
        level: "ApplicationConsole.LogLevel" | None = None,
    ) -> None:
        if not message:
            return
        widget = self._text_widget
        if widget is None:
            return

        current_level = level if level is not None else self.LogLevel.WARN
        if current_level < self._level:
            return

        def _append() -> None:
            widget.configure(state="normal")
            if widget.index("end-1c") != "1.0":
                widget.insert(tk.END, "\n")
            widget.insert(tk.END, message)
            widget.see(tk.END)
            widget.configure(state="disabled")

        widget.after(0, _append)

    def log_info(self, message: str) -> None:
        self.log(message, level=self.LogLevel.INFO)

    def log_warn(self, message: str) -> None:
        self.log(message, level=self.LogLevel.WARN)


def create_main_window() -> tk.Tk:
    """Create and configure the main application window."""
    root = tk.Tk()
    root.title("FIRST Global 2025 Team Video Slotter")
    root.update_idletasks()
    screen_width = root.winfo_screenwidth() or 1280
    screen_height = root.winfo_screenheight() or 720
    default_width = max(800, int(screen_width * 0.5))
    default_height = max(600, int(screen_height * (3 / 4)))
    root.geometry(f"{default_width}x{default_height}")

    # Configure a grid layout so the notebook expands with the window.
    root.columnconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)

    title_label = ttk.Label(
        root,
        text="FIRST Global 2025 Team Video Slotter",
        font=("Helvetica", 18, "bold"),
        anchor="center",
    )
    title_label.grid(row=0, column=0, padx=16, pady=(16, 8), sticky="ew")

    paned = ttk.PanedWindow(root, orient="vertical")
    paned.grid(row=1, column=0, padx=16, pady=(0, 16), sticky="nsew")

    notebook_container = ttk.Frame(paned)
    notebook_container.columnconfigure(0, weight=1)
    notebook_container.rowconfigure(0, weight=1)

    notebook = ttk.Notebook(notebook_container)
    notebook.grid(row=0, column=0, sticky="nsew")

    team_videos_frame = ttk.Frame(notebook)
    config_frame = ttk.Frame(notebook)
    tools_frame = ttk.Frame(notebook)

    for frame in (team_videos_frame, config_frame, tools_frame):
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

    notebook.add(team_videos_frame, text="Team Videos")
    notebook.add(config_frame, text="Config")
    notebook.add(tools_frame, text="Tools")

    paned.add(notebook_container, weight=3)

    console_container = ttk.Frame(paned)
    console_container.columnconfigure(0, weight=1)
    console_container.rowconfigure(0, weight=1)
    paned.add(console_container, weight=1)

    console = ApplicationConsole(console_container)
    console.render(row=0)
    (
        credentials_manager,
        ros_document_loader,
        match_schedule_importer,
    ) = build_config_tab(config_frame, console)
    build_team_videos_tab(
        team_videos_frame,
        console,
        credentials_manager,
        ros_document_loader,
        match_schedule_importer,
    )
    build_tools_tab(
        tools_frame,
        console,
        credentials_manager,
        ros_document_loader,
        match_schedule_importer,
    )

    return root


def build_team_videos_tab(
    parent: ttk.Frame,
    console: ApplicationConsole,
    credentials_manager: "GoogleDriveCredentialsManager",
    ros_document_loader: "ROSDocumentLoaderUI",
    match_schedule_importer: "MatchScheduleImporterUI",
) -> None:
    """Populate the Team Videos tab with optimization tooling."""

    parent.columnconfigure(0, weight=1)
    parent.rowconfigure(0, weight=1)

    container = ttk.Frame(parent, padding=(12, 12, 12, 12))
    container.grid(row=0, column=0, sticky="nsew")
    container.columnconfigure(1, weight=1)

    ttk.Label(
        container,
        text="Optimize Team Videos",
        font=("Helvetica", 14, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w")

    optimizer = OptimizeTeamVideosUI(
        container,
        credentials_manager,
        ros_document_loader,
        match_schedule_importer,
        console,
    )
    optimizer.render(row=1)

    separator = ttk.Separator(container, orient="horizontal")
    separator.grid(row=4, column=0, columnspan=2, sticky="ew", pady=12)

    fill_script_frame = ttk.Frame(container, padding=(0, 0, 0, 0))
    fill_script_frame.grid(row=5, column=0, columnspan=2, sticky="nsew")
    fill_script_frame.columnconfigure(1, weight=1)

    ttk.Label(
        fill_script_frame,
        text="Fill Script",
        font=("Helvetica", 14, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w")

    fill_script = FillScriptUI(
        fill_script_frame,
        credentials_manager,
        ros_document_loader,
        console,
    )
    fill_script.render(row=1)


def build_config_tab(
    parent: ttk.Frame, console: ApplicationConsole
) -> Tuple[
    "GoogleDriveCredentialsManager",
    "ROSDocumentLoaderUI",
    "MatchScheduleImporterUI",
]:
    """Populate the Config tab with shared application settings."""

    parent.columnconfigure(0, weight=1)

    load_drive_frame = ttk.Frame(parent, padding=(12, 12, 12, 6))
    load_drive_frame.grid(row=0, column=0, sticky="nsew")
    load_drive_frame.columnconfigure(1, weight=1)

    ttk.Label(
        load_drive_frame,
        text="Load Google Drive Credentials",
        font=("Helvetica", 14, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w")

    credentials_manager = GoogleDriveCredentialsManager(load_drive_frame, console)
    credentials_manager.render(row=1)

    separator_one = ttk.Separator(parent, orient="horizontal")
    separator_one.grid(row=1, column=0, sticky="ew", padx=12, pady=6)

    ros_document_frame = ttk.Frame(parent, padding=(12, 6, 12, 12))
    ros_document_frame.grid(row=2, column=0, sticky="nsew")
    ros_document_frame.columnconfigure(1, weight=1)

    ttk.Label(
        ros_document_frame,
        text="Load ROS Document",
        font=("Helvetica", 14, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w")

    ros_document_loader = ROSDocumentLoaderUI(
        ros_document_frame, console, credentials_manager=credentials_manager
    )
    ros_document_loader.render(row=1)

    separator_two = ttk.Separator(parent, orient="horizontal")
    separator_two.grid(row=3, column=0, sticky="ew", padx=12, pady=6)

    match_schedule_frame = ttk.Frame(parent, padding=(12, 6, 12, 12))
    match_schedule_frame.grid(row=4, column=0, sticky="nsew")
    match_schedule_frame.columnconfigure(1, weight=1)

    ttk.Label(
        match_schedule_frame,
        text="Import Match Schedule",
        font=("Helvetica", 14, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w")

    match_schedule_importer = MatchScheduleImporterUI(match_schedule_frame, console)
    match_schedule_importer.render(row=1)

    return credentials_manager, ros_document_loader, match_schedule_importer


def build_tools_tab(
    parent: ttk.Frame,
    console: ApplicationConsole,
    credentials_manager: "GoogleDriveCredentialsManager",
    ros_document_loader: "ROSDocumentLoaderUI",
    match_schedule_importer: "MatchScheduleImporterUI",
) -> None:
    """Populate the Tools tab with utilities that operate on ROS documents."""

    parent.columnconfigure(0, weight=1)

    ros_frame = ttk.Frame(parent, padding=(12, 12, 12, 6))
    ros_frame.grid(row=0, column=0, sticky="nsew")
    ros_frame.columnconfigure(1, weight=1)

    ttk.Label(
        ros_frame,
        text="ROS Placeholder Generator",
        font=("Helvetica", 14, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w")

    ROSPlaceholderGeneratorUI(ros_frame, credentials_manager, ros_document_loader, console).render(
        row=1
    )

    separator = ttk.Separator(parent, orient="horizontal")
    separator.grid(row=1, column=0, sticky="ew", padx=12, pady=6)

    match_frame = ttk.Frame(parent, padding=(12, 6, 12, 12))
    match_frame.grid(row=2, column=0, sticky="nsew")
    match_frame.columnconfigure(1, weight=1)

    ttk.Label(
        match_frame,
        text="Match Number Generator",
        font=("Helvetica", 14, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w")

    MatchNumberGeneratorUI(
        match_frame,
        credentials_manager,
        ros_document_loader,
        match_schedule_importer,
        console,
    ).render(row=1)


class MatchScheduleImporterUI:
    """Import match schedules from JSON files within the Config tab."""

    FIELD_OPTIONS = tuple(str(number) for number in range(1, 6))

    def __init__(self, parent: ttk.Frame, console: ApplicationConsole) -> None:
        self.parent = parent
        self.console = console
        self.file_path_var = tk.StringVar()
        self.field_number_var = tk.StringVar(value=self.FIELD_OPTIONS[0])
        self._status_var = tk.StringVar()
        self._matches: List[Dict[str, Any]] = []
        self._imported_field: Optional[int] = None
        self._matches_by_date: Dict[str, List[Dict[str, Any]]] = {}

    def render(self, row: int) -> None:
        ttk.Label(self.parent, text="Selected file:").grid(
            row=row, column=0, sticky="w", pady=(8, 0)
        )

        file_display = ttk.Entry(
            self.parent,
            textvariable=self.file_path_var,
            state="readonly",
        )
        file_display.grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        ttk.Label(self.parent, text="Field:").grid(
            row=row + 1, column=0, sticky="w", pady=(8, 0)
        )

        field_selector = ttk.Combobox(
            self.parent,
            textvariable=self.field_number_var,
            values=self.FIELD_OPTIONS,
            state="readonly",
            width=5,
        )
        field_selector.grid(row=row + 1, column=1, sticky="w", padx=(8, 0), pady=(8, 0))

        import_button = ttk.Button(
            self.parent,
            text="Import Match Schedule",
            command=self.import_schedule,
        )
        import_button.grid(row=row + 2, column=1, sticky="e", pady=8)

        status_label = ttk.Label(
            self.parent,
            textvariable=self._status_var,
            wraplength=520,
            justify="left",
        )
        status_label.grid(row=row + 3, column=0, columnspan=2, sticky="w")

        self._set_status("No match schedule imported yet.", log=False)

    def import_schedule(self) -> None:
        field_number = self._get_selected_field()
        if field_number is None:
            self._set_status("Select a valid field number between 1 and 5.")
            return

        filename = filedialog.askopenfilename(
            parent=self.parent,
            title="Select match schedule JSON",
            filetypes=[("JSON files", "*.json"), ("All files", "*")],
        )
        if not filename:
            return

        try:
            with open(filename, "r", encoding="utf-8") as stream:
                data = json.load(stream)
        except FileNotFoundError:
            self._set_status("Selected file could not be found.")
            return
        except json.JSONDecodeError as exc:
            self._set_status(f"Failed to parse JSON: {exc}")
            return
        except OSError as exc:
            self._set_status(f"Unable to read file: {exc}")
            return

        matches = data.get("matches") if isinstance(data, dict) else None
        if not isinstance(matches, list):
            self._set_status("JSON file does not contain a 'matches' list.")
            return

        self._matches = matches
        self._imported_field = field_number
        self._matches_by_date = self._group_matches_by_date(matches, field_number)
        self.file_path_var.set(filename)

        field_counts = self._count_field_matches(matches, field_number)
        total_matches = len(matches)
        field_total = sum(field_counts.values())

        self._log_field_counts(field_counts, field_number)

        if field_counts:
            summary = (
                f"Imported {total_matches} matches. {field_total} occur on Field {field_number}."
            )
        else:
            summary = (
                f"Imported {total_matches} matches. No matches found on Field {field_number}."
            )
        self._set_status(summary)

    def _get_selected_field(self) -> Optional[int]:
        try:
            field_number = int(self.field_number_var.get())
        except (TypeError, ValueError):
            return None
        if field_number < 1 or field_number > 5:
            return None
        return field_number

    def _count_field_matches(
        self, matches: Sequence[Dict[str, Any]], field_number: int
    ) -> Dict[str, int]:
        counts: Dict[str, int] = {}
        for match in matches:
            field = match.get("field")
            if field != field_number:
                continue
            scheduled_time = match.get("scheduledTime")
            if not isinstance(scheduled_time, str):
                continue
            date_str = self._extract_date(scheduled_time)
            if not date_str:
                continue
            counts[date_str] = counts.get(date_str, 0) + 1
        return dict(sorted(counts.items()))

    def _group_matches_by_date(
        self, matches: Sequence[Dict[str, Any]], field_number: int
    ) -> Dict[str, List[Dict[str, Any]]]:
        grouped: Dict[str, List[Dict[str, Any]]] = {}
        for match in matches:
            if match.get("field") != field_number:
                continue
            scheduled_time = match.get("scheduledTime")
            if not isinstance(scheduled_time, str):
                continue
            date_str = self._extract_date(scheduled_time)
            if not date_str:
                continue
            grouped.setdefault(date_str, []).append(match)

        for date, items in grouped.items():
            items.sort(key=self._match_sort_key)

        return {date: list(grouped[date]) for date in sorted(grouped)}

    def _extract_date(self, timestamp: str) -> Optional[str]:
        timestamp = timestamp.strip()
        if not timestamp:
            return None
        normalized = timestamp.replace("Z", "+00:00")
        try:
            return datetime.fromisoformat(normalized).date().isoformat()
        except ValueError:
            if "T" in timestamp:
                candidate = timestamp.split("T", 1)[0]
                if re.fullmatch(r"\d{4}-\d{2}-\d{2}", candidate):
                    return candidate
            return None

    def _parse_datetime(self, timestamp: Any) -> Optional[datetime]:
        if not isinstance(timestamp, str):
            return None
        normalized = timestamp.strip()
        if not normalized:
            return None
        normalized = normalized.replace("Z", "+00:00")
        try:
            return datetime.fromisoformat(normalized)
        except ValueError:
            return None

    def _match_sort_key(self, match: Dict[str, Any]) -> Tuple[Any, Any, Any]:
        scheduled_dt = self._parse_datetime(match.get("scheduledTime"))
        if scheduled_dt is not None and scheduled_dt.tzinfo is not None:
            scheduled_dt = scheduled_dt.astimezone(timezone.utc).replace(tzinfo=None)
        if scheduled_dt is not None:
            scheduled_key: Tuple[Any, ...] = (
                0,
                scheduled_dt.toordinal(),
                scheduled_dt.hour,
                scheduled_dt.minute,
                scheduled_dt.second,
                scheduled_dt.microsecond,
            )
        else:
            scheduled_key = (1, 0, 0, 0, 0, 0, 0)

        match_number = self.extract_match_number(match)
        fallback = (
            match.get("matchKey")
            or match.get("description")
            or match.get("name")
            or ""
        )
        return (
            scheduled_key,
            match_number if match_number is not None else float("inf"),
            fallback,
        )

    def _log_field_counts(self, counts: Dict[str, int], field_number: int) -> None:
        if not counts:
            self.console.log(
                f"[Match Schedule Importer] No matches on Field {field_number} were found."
            )
            return
        for date, total in counts.items():
            message = f"{date} has {total} matches on Field {field_number}"
            self.console.log(f"[Match Schedule Importer] {message}")

    def _set_status(self, message: str, *, log: bool = True) -> None:
        self._status_var.set(message)
        if log:
            self.console.log(f"[Match Schedule Importer] {message}")

    def has_loaded_schedule(self) -> bool:
        return bool(self._matches)

    def get_imported_field_number(self) -> Optional[int]:
        return self._imported_field

    def get_matches_by_date_for_selected_field(self) -> List[Tuple[str, List[Dict[str, Any]]]]:
        return [
            (date, list(matches))
            for date, matches in sorted(self._matches_by_date.items())
        ]

    def get_matches_for_selected_field(self) -> List[Dict[str, Any]]:
        field_number = self._imported_field
        if field_number is None:
            return []
        selected: List[Dict[str, Any]] = []
        for match in self._matches:
            if not isinstance(match, dict):
                continue
            if match.get("field") != field_number:
                continue
            selected.append(dict(match))
        return selected

    def extract_match_number(self, match: Dict[str, Any]) -> Optional[int]:
        for key in (
            "matchNumber",
            "match_number",
            "matchNumberDisplay",
            "matchnumber",
            "matchNo",
            "id",
        ):
            number = self._coerce_match_number(match.get(key))
            if number is not None:
                return number

        for fallback_key in ("matchKey", "match"):
            number = self._coerce_match_number(match.get(fallback_key))
            if number is not None:
                return number

        return None

    def describe_match(self, match: Mapping[str, Any]) -> str:
        """Return a human-readable summary of a match entry for diagnostics."""

        def _format_value(value: Any) -> Optional[str]:
            if value is None or isinstance(value, bool):
                return None
            if isinstance(value, (int, float)):
                if isinstance(value, float) and not value.is_integer():
                    return str(value)
                return str(int(value)) if isinstance(value, float) else str(value)
            if isinstance(value, str):
                text = value.strip()
                if text:
                    return text
                return None
            return None

        details: List[str] = []
        for label in ("matchKey", "description", "name", "id"):
            formatted = _format_value(match.get(label))
            if formatted:
                details.append(f"{label}={formatted}")
                break

        scheduled = _format_value(match.get("scheduledTime"))
        if scheduled:
            details.append(f"scheduledTime={scheduled}")

        field_value = match.get("field")
        field_formatted = _format_value(field_value)
        if field_formatted:
            details.append(f"field={field_formatted}")

        if details:
            return ", ".join(details)

        keys = ", ".join(sorted(map(str, match.keys())))
        if keys:
            return f"available keys: {keys}"
        return "no additional details available"

    def _coerce_match_number(self, value: Any) -> Optional[int]:
        if isinstance(value, bool):
            return None
        if isinstance(value, int):
            return value
        if isinstance(value, float) and value.is_integer():
            return int(value)
        if isinstance(value, str):
            stripped = value.strip()
            if not stripped:
                return None
            if stripped.isdigit():
                return int(stripped)
            digits = re.findall(r"\d+", stripped)
            if digits:
                return int(digits[-1])
        return None

class GoogleDriveCredentialsManager:
    """Handle selection and loading of Google Drive OAuth credentials."""

    SCOPES = [
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/documents",
    ]
    STORAGE_PATH = Path.home() / ".fgc_team_filler" / "google_credentials.json"
    USER_POLL_INTERVAL_MS = 60_000

    def __init__(self, parent: ttk.Frame, console: ApplicationConsole) -> None:
        self.parent = parent
        self.console = console
        self.credentials_path_var = tk.StringVar()
        self.logged_in_user_var = tk.StringVar(value="No credentials loaded")
        self._credentials: Optional[Credentials] = None
        self._user_poll_after_id: Optional[str] = None
        self._last_user_message: Optional[str] = None
        self._credential_listeners: List[Callable[[Optional[Credentials]], None]] = []

    @property
    def credentials(self) -> Optional[Credentials]:
        return self._credentials

    def get_valid_credentials(
        self, *, log_status: bool = True
    ) -> Tuple[Optional[Credentials], Optional[str]]:
        """Return credentials that are refreshed and valid, or an error message."""

        credentials = self._credentials
        if credentials is None:
            return None, None

        if credentials.expired and getattr(credentials, "refresh_token", None):
            if Request is None:
                message = (
                    "Stored credentials have expired and google-auth-transport-requests is missing. "
                    "Install it with 'pip install google-auth'."
                )
                if log_status:
                    self.set_status(message)
                return None, message
            try:
                credentials.refresh(Request())
            except Exception as exc:  # pragma: no cover - network interaction
                message = f"Failed to refresh credentials: {exc}"[:500]
                if log_status:
                    self.set_status(message)
                return None, message
            else:
                self._credentials = credentials
                self._persist_credentials(credentials)
                self._notify_credentials_listeners()

        if not credentials.valid:
            message = "Stored credentials are invalid. Please re-authorize."
            if log_status:
                self.set_status(message)
            return None, message

        if not self._credentials_have_required_scopes(credentials):
            message = (
                "Stored credentials are missing Google Docs access. Please re-authorize."
            )
            if log_status:
                self.set_status(message)
            return None, message

        return credentials, None

    def render(self, row: int) -> None:
        ttk.Label(self.parent, text="Credentials file:").grid(
            row=row, column=0, sticky="w", pady=(8, 0)
        )

        path_display = ttk.Entry(
            self.parent,
            textvariable=self.credentials_path_var,
            state="readonly",
        )
        path_display.grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        select_button = ttk.Button(
            self.parent,
            text="Select credentials.json",
            command=self.select_credentials_file,
        )
        select_button.grid(row=row + 1, column=0, sticky="w", pady=8)

        authorize_button = ttk.Button(
            self.parent,
            text="Authorize",
            command=self.authorize,
        )
        authorize_button.grid(row=row + 1, column=1, sticky="e", pady=8)

        ttk.Label(self.parent, text="Currently Logged-in User:").grid(
            row=row + 2, column=0, sticky="w"
        )

        user_display = ttk.Entry(
            self.parent,
            textvariable=self.logged_in_user_var,
            state="readonly",
        )
        user_display.grid(row=row + 2, column=1, sticky="ew", padx=(8, 0), pady=(4, 0))

        self.set_status("No credentials loaded.", log=False)
        self._load_persisted_credentials()
        self._schedule_user_poll(0)

    def add_credentials_listener(
        self, callback: Callable[[Optional[Credentials]], None]
    ) -> None:
        self._credential_listeners.append(callback)
        try:
            callback(self._credentials)
        except Exception:
            pass

    def _notify_credentials_listeners(self) -> None:
        for listener in list(self._credential_listeners):
            try:
                listener(self._credentials)
            except Exception:
                continue

    def _credentials_have_required_scopes(self, credentials: Credentials) -> bool:
        try:
            has_scopes = credentials.has_scopes(self.SCOPES)  # type: ignore[attr-defined]
        except Exception:
            has_scopes = None

        if has_scopes is not None:
            return bool(has_scopes)

        scopes = getattr(credentials, "scopes", None)
        if isinstance(scopes, Sequence):
            scope_set = {str(scope) for scope in scopes}
            return all(scope in scope_set for scope in self.SCOPES)

        return False

    def select_credentials_file(self) -> None:
        filename = filedialog.askopenfilename(
            parent=self.parent,
            title="Select Google Drive credentials.json",
            filetypes=[("JSON files", "*.json"), ("All files", "*")],
        )
        if filename:
            self.credentials_path_var.set(filename)
            self.set_status("Credentials file selected. Click Authorize to continue.")

    def authorize(self) -> None:
        credentials_path = self.credentials_path_var.get().strip()
        if not credentials_path:
            self.set_status("Please select a credentials.json file first.")
            return

        if InstalledAppFlow is None:
            self.set_status(
                "google-auth-oauthlib is not installed. Install it with 'pip install google-auth-oauthlib'."
            )
            return

        self.set_status("Starting OAuth flow...")

        def _run_flow() -> None:
            try:
                flow = InstalledAppFlow.from_client_secrets_file(
                    credentials_path, self.SCOPES
                )
                credentials = flow.run_local_server(port=0, open_browser=True)
            except Exception as exc:  # pragma: no cover - user driven
                self.parent.after(
                    0,
                    lambda: self.set_status(f"Authorization failed: {exc}"[:500]),
                )
                return

            def _on_success() -> None:
                self._credentials = credentials
                message = self._persist_credentials(credentials)
                if message is None:
                    message = "Authorization successful. Credentials are loaded in memory."
                self.set_status(message)
                self._schedule_user_poll(0)
                self._notify_credentials_listeners()

            self.parent.after(0, _on_success)

        threading.Thread(target=_run_flow, daemon=True).start()

    def set_status(self, message: str, *, log: bool = True) -> None:
        if log:
            self.console.log(f"[Drive Credentials] {message}")

    def _load_persisted_credentials(self) -> None:
        if Credentials is None:
            return

        storage_path = self.STORAGE_PATH
        if not storage_path.exists():
            return

        try:
            credentials = Credentials.from_authorized_user_file(
                str(storage_path), self.SCOPES
            )
        except Exception as exc:  # pragma: no cover - file parsing
            self.set_status(f"Failed to load saved credentials: {exc}"[:500])
            return

        if not credentials.valid and getattr(credentials, "refresh_token", None):
            if Request is None:
                self.set_status(
                    "Saved credentials are expired but google-auth is missing the Requests transport."
                )
                return
            try:
                credentials.refresh(Request())
            except Exception as exc:  # pragma: no cover - network interaction
                self.set_status(f"Failed to refresh saved credentials: {exc}"[:500])
                return

        if credentials.valid:
            if not self._credentials_have_required_scopes(credentials):
                self.set_status(
                    "Saved credentials are missing Google Docs access. Please re-authorize."
                )
                return
            self._credentials = credentials
            self.credentials_path_var.set(str(storage_path))
            self.set_status("Loaded saved credentials. You're ready to go.")
            self._persist_credentials(credentials)
            self._schedule_user_poll(0)
            self._notify_credentials_listeners()
        else:
            self.set_status("Saved credentials are invalid. Please re-authorize.")

    def _persist_credentials(self, credentials: Credentials) -> Optional[str]:
        if Credentials is None:
            return None

        storage_path = self.STORAGE_PATH
        try:
            storage_path.parent.mkdir(parents=True, exist_ok=True)
            storage_path.write_text(credentials.to_json())
        except Exception as exc:  # pragma: no cover - filesystem issues
            return (
                "Authorization succeeded, but saving credentials failed: "
                f"{exc}"[:500]
            )
        self.credentials_path_var.set(str(storage_path))
        return f"Authorization successful. Credentials saved to {storage_path}."

    def _schedule_user_poll(self, delay_ms: int) -> None:
        if self._user_poll_after_id is not None:
            try:
                self.parent.after_cancel(self._user_poll_after_id)
            except Exception:
                pass
        self._user_poll_after_id = self.parent.after(delay_ms, self._poll_logged_in_user)

    def _poll_logged_in_user(self) -> None:
        credentials, error = self.get_valid_credentials(log_status=False)
        if credentials is None:
            if error:
                self._set_logged_in_user("Unavailable")
                self._log_user_lookup_message(error)
            else:
                self._set_logged_in_user("No credentials loaded")
                self._last_user_message = None
            self._schedule_user_poll(self.USER_POLL_INTERVAL_MS)
            return

        if build is None:
            message = (
                "google-api-python-client is not installed. Install it to fetch the logged-in user."
            )
            self._set_logged_in_user("Library missing")
            self._log_user_lookup_message(message)
            self._schedule_user_poll(self.USER_POLL_INTERVAL_MS)
            return

        threading.Thread(
            target=self._fetch_logged_in_user, args=(credentials,), daemon=True
        ).start()

    def _fetch_logged_in_user(self, credentials: Credentials) -> None:
        try:
            service = build("drive", "v3", credentials=credentials, cache_discovery=False)
            about = service.about().get(fields="user(displayName,emailAddress)").execute()
        except Exception as exc:  # pragma: no cover - network interaction
            message = f"Failed to fetch logged-in user: {exc}"[:500]
            self.parent.after(0, lambda: self._handle_user_lookup_failure(message))
            return

        user_info = about.get("user", {}) if isinstance(about, dict) else {}
        display_name = user_info.get("displayName")
        email = user_info.get("emailAddress")
        self.parent.after(
            0,
            lambda: self._handle_user_lookup_success(
                str(display_name) if display_name else "",
                str(email) if email else "",
            ),
        )

    def _handle_user_lookup_success(self, display_name: str, email: str) -> None:
        parts = [part for part in (display_name.strip(), email.strip()) if part]
        display_text = " • ".join(parts) if parts else "Unknown user"
        self._set_logged_in_user(display_text)
        self._log_user_lookup_message(f"Logged in as {display_text}.")
        self._schedule_user_poll(self.USER_POLL_INTERVAL_MS)

    def _handle_user_lookup_failure(self, message: str) -> None:
        self._set_logged_in_user("Lookup failed")
        self._log_user_lookup_message(message)
        self._schedule_user_poll(self.USER_POLL_INTERVAL_MS)

    def _set_logged_in_user(self, value: str) -> None:
        self.logged_in_user_var.set(value)

    def _log_user_lookup_message(self, message: str) -> None:
        if not message:
            return
        if message == self._last_user_message:
            return
        self.console.log(f"[Drive Credentials] {message}")
        self._last_user_message = message


class ROSDocumentLoaderUI:
    """Allow the user to select and persist the ROS spreadsheet URL."""

    STORAGE_PATH = Path.home() / ".fgc_team_filler" / "ros_document_url.txt"

    def __init__(
        self,
        parent: ttk.Frame,
        console: ApplicationConsole,
        *,
        credentials_manager: Optional["GoogleDriveCredentialsManager"] = None,
    ) -> None:
        self.parent = parent
        self.console = console
        self._credentials_manager = credentials_manager
        self.sheet_url_var = tk.StringVar(value=self._load_saved_url())
        self.document_name_var = tk.StringVar()
        self._status_var = tk.StringVar()
        self._listeners: List[Callable[[str], None]] = []
        self._name_listeners: List[Callable[[str], None]] = []
        self._resolved_names: Dict[str, str] = {}
        self._lookup_in_progress: Set[str] = set()

        self.sheet_url_var.trace_add("write", self._on_url_var_changed)
        self._update_document_name(self.sheet_url_var.get())

        if self._credentials_manager is not None:
            self._credentials_manager.add_credentials_listener(
                self._on_credentials_changed
            )

    def render(self, row: int) -> None:
        ttk.Label(self.parent, text="Google Sheets URL:").grid(
            row=row, column=0, sticky="w", pady=(8, 0)
        )

        entry = ttk.Entry(self.parent, textvariable=self.sheet_url_var)
        entry.grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        ttk.Label(self.parent, text="File name:").grid(
            row=row + 1, column=0, sticky="w", pady=(4, 0)
        )

        name_display = ttk.Entry(
            self.parent,
            textvariable=self.document_name_var,
            state="readonly",
        )
        name_display.grid(row=row + 1, column=1, sticky="ew", padx=(8, 0), pady=(4, 0))

        save_button = ttk.Button(
            self.parent,
            text="Save Document URL",
            command=self.save_document_url,
        )
        save_button.grid(row=row + 2, column=1, sticky="e", pady=8)

        status_label = ttk.Label(
            self.parent,
            textvariable=self._status_var,
            wraplength=520,
            justify="left",
        )
        status_label.grid(row=row + 3, column=0, columnspan=2, sticky="w")

        initial_url = self.sheet_url_var.get().strip()
        if initial_url:
            self._set_status("Loaded saved ROS document URL.", log=False)
        else:
            self._set_status("Enter a Google Sheets link to use with ROS tools.", log=False)

    def add_listener(self, callback: Callable[[str], None]) -> None:
        self._listeners.append(callback)

    def add_name_listener(self, callback: Callable[[str], None]) -> None:
        self._name_listeners.append(callback)

    def get_document_url(self) -> str:
        return self.sheet_url_var.get().strip()

    def get_document_name(self) -> str:
        return self.document_name_var.get().strip()

    def save_document_url(self) -> None:
        url = self.sheet_url_var.get().strip()
        if not url:
            self._set_status("Please enter a Google Sheets link before saving.")
            return

        spreadsheet_id = extract_spreadsheet_id(url)
        if not spreadsheet_id:
            self._set_status("Unable to determine spreadsheet ID from the provided link.")
            return

        storage_path = self.STORAGE_PATH
        try:
            storage_path.parent.mkdir(parents=True, exist_ok=True)
            storage_path.write_text(url)
        except Exception as exc:  # pragma: no cover - filesystem issues
            self._set_status(f"Failed to save document URL: {exc}"[:500])
            return

        self._set_status("ROS document URL saved.")
        self._update_document_name(url)
        for listener in list(self._listeners):
            try:
                listener(url)
            except Exception:
                continue

    def _load_saved_url(self) -> str:
        storage_path = self.STORAGE_PATH
        if not storage_path.exists():
            return ""
        try:
            return storage_path.read_text().strip()
        except Exception:
            return ""

    def _set_status(self, message: str, *, log: bool = True) -> None:
        self._status_var.set(message)
        if log:
            self.console.log(f"[ROS Document] {message}")

    def _on_url_var_changed(self, *_: Any) -> None:
        self._update_document_name(self.sheet_url_var.get())

    def set_document_name(self, name: str) -> None:
        text = (name or "").strip()
        if text:
            spreadsheet_id = extract_spreadsheet_id(self.sheet_url_var.get().strip())
            if spreadsheet_id:
                self._resolved_names[spreadsheet_id] = text
        display = text if text else "Unknown"
        self._set_document_name_display(display)

    def _update_document_name(self, url: str) -> None:
        url = (url or "").strip()
        if not url:
            self._set_document_name_display("Not set")
            return

        spreadsheet_id = extract_spreadsheet_id(url)
        if spreadsheet_id:
            cached_name = self._resolved_names.get(spreadsheet_id)
            if cached_name:
                self._set_document_name_display(cached_name)
                return

        name = derive_document_name(url)
        if not name:
            if spreadsheet_id:
                name = f"Spreadsheet {spreadsheet_id}"
        display = name or "Unknown"

        self._set_document_name_display(display)
        if spreadsheet_id:
            self._maybe_lookup_remote_title(spreadsheet_id)

    def _set_document_name_display(self, display: str) -> None:
        self.document_name_var.set(display)
        for listener in list(self._name_listeners):
            try:
                listener(display)
            except Exception:
                continue

    def _on_credentials_changed(self, credentials: Optional[Credentials]) -> None:
        if credentials is None:
            return
        url = self.sheet_url_var.get().strip()
        spreadsheet_id = extract_spreadsheet_id(url)
        if spreadsheet_id:
            self._maybe_lookup_remote_title(spreadsheet_id, credentials=credentials)

    def _maybe_lookup_remote_title(
        self,
        spreadsheet_id: str,
        *,
        credentials: Optional[Credentials] = None,
    ) -> None:
        if not spreadsheet_id:
            return
        if spreadsheet_id in self._resolved_names:
            return
        if spreadsheet_id in self._lookup_in_progress:
            return
        if build is None:
            return

        manager = self._credentials_manager
        if credentials is None:
            if manager is None:
                return
            credentials, error = manager.get_valid_credentials(log_status=False)
            if credentials is None:
                if error:
                    self.console.log(f"[ROS Document] {error}")
                return
        self._lookup_in_progress.add(spreadsheet_id)

        def _worker() -> None:
            try:
                spreadsheet, _theme_supported, _service = _fetch_spreadsheet(
                    credentials, spreadsheet_id
                )
            except Exception as exc:  # pragma: no cover - network interaction
                message = f"Failed to fetch spreadsheet details: {exc}"[:500]
                self.parent.after(
                    0,
                    lambda: self._finalize_remote_lookup(
                        spreadsheet_id, None, error_message=message
                    ),
                )
                return

            title = str(spreadsheet.get("properties", {}).get("title", "")).strip()
            self.parent.after(
                0,
                lambda: self._finalize_remote_lookup(spreadsheet_id, title or None),
            )

        threading.Thread(target=_worker, daemon=True).start()

    def _finalize_remote_lookup(
        self,
        spreadsheet_id: str,
        title: Optional[str],
        *,
        error_message: Optional[str] = None,
    ) -> None:
        self._lookup_in_progress.discard(spreadsheet_id)
        if error_message:
            self.console.log(f"[ROS Document] {error_message}")
        if not title:
            return

        self._resolved_names[spreadsheet_id] = title
        current_id = extract_spreadsheet_id(self.sheet_url_var.get().strip())
        if current_id == spreadsheet_id:
            self._set_document_name_display(title)


class ROSPlaceholderGeneratorUI:
    """User interface wrapper for the ROS placeholder generator tool."""

    def __init__(
        self,
        parent: ttk.Frame,
        credentials_manager: GoogleDriveCredentialsManager,
        document_loader: ROSDocumentLoaderUI,
        console: ApplicationConsole,
    ) -> None:
        self.parent = parent
        self.credentials_manager = credentials_manager
        self.document_loader = document_loader
        self.console = console
        self.current_document_var = tk.StringVar()
        self._status_var = tk.StringVar()
        self._default_status = "Load a ROS document and Google credentials to begin."

        self.document_loader.add_listener(self._on_document_url_changed)
        self.document_loader.add_name_listener(self._on_document_name_changed)
        self._on_document_url_changed(self.document_loader.get_document_url())
        self._on_document_name_changed(self.document_loader.get_document_name())

    def render(self, row: int) -> None:
        ttk.Label(self.parent, text="Current ROS Document:").grid(
            row=row, column=0, sticky="w", pady=(8, 0)
        )

        document_display = ttk.Entry(
            self.parent,
            textvariable=self.current_document_var,
            state="readonly",
        )
        document_display.grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        generate_button = ttk.Button(
            self.parent,
            text="Generate Placeholders",
            command=self.generate_placeholders,
        )
        generate_button.grid(row=row + 1, column=1, sticky="e", pady=8)

        status_label = ttk.Label(
            self.parent,
            textvariable=self._status_var,
            wraplength=520,
            justify="left",
        )
        status_label.grid(row=row + 2, column=0, columnspan=2, sticky="w")

        self.set_status(self._default_status, log=False)

    def _on_document_url_changed(self, url: str) -> None:
        if not url:
            self.set_status("Save a ROS document URL before generating placeholders.", log=False)

    def _on_document_name_changed(self, name: str) -> None:
        display_value = name if name else "Not set"
        self.current_document_var.set(display_value)

    def generate_placeholders(self) -> None:
        document_url = self.document_loader.get_document_url()
        if not document_url:
            self.set_status("Save a ROS document URL before generating placeholders.")
            return

        credentials, error = self.credentials_manager.get_valid_credentials()
        if credentials is None:
            if error:
                self.set_status(error)
            else:
                self.set_status("Load Google Drive credentials before running this tool.")
            return

        if build is None:
            self.set_status(
                "google-api-python-client is not installed. Install it with 'pip install google-api-python-client'."
            )
            return

        spreadsheet_id = extract_spreadsheet_id(document_url)
        if not spreadsheet_id:
            self.set_status("Unable to determine spreadsheet ID from the saved document URL.")
            return

        self.set_status("Contacting Google Sheets API...")

        def _worker() -> None:
            try:
                report, diagnostics, spreadsheet_title = generate_placeholders_for_sheet(
                    credentials, spreadsheet_id
                )
            except Exception as exc:  # pragma: no cover - network interaction
                message = f"Failed to update spreadsheet: {exc}"
                self.parent.after(0, lambda: self.set_status(message))
                return

            self.parent.after(
                0,
                lambda: self._handle_placeholder_success(
                    report, diagnostics, spreadsheet_title
                ),
            )

        threading.Thread(target=_worker, daemon=True).start()

    def set_status(self, message: str, *, log: bool = True) -> None:
        self._status_var.set(message)
        if log:
            self.console.log(f"[ROS Placeholder Generator] {message}")

    def _handle_placeholder_success(
        self,
        report: Dict[str, List[str]],
        diagnostics: Sequence[str],
        spreadsheet_title: str,
    ) -> None:
        if spreadsheet_title:
            self.document_loader.set_document_name(spreadsheet_title)

        console_message = format_report(report, diagnostics)

        if report:
            total_updates = sum(len(entries) for entries in report.values())
            sheet_count = len(report)
            cell_word = "cell" if total_updates == 1 else "cells"
            sheet_word = "sheet" if sheet_count == 1 else "sheets"
            status_message = (
                "Placeholder updates applied to "
                f"{total_updates} {cell_word} across {sheet_count} {sheet_word}."
            )
            breakdowns = []
            for sheet, entries in report.items():
                entry_count = len(entries)
                entry_cell_word = "cell" if entry_count == 1 else "cells"
                breakdowns.append(
                    f"{entry_count} {entry_cell_word} updated on sheet ({sheet})"
                )
            if breakdowns:
                status_message += " " + ", ".join(breakdowns)
        else:
            status_message = "No matching placeholders were found."

        self.console.log_info(f"[ROS Placeholder Generator] {console_message}")
        self.set_status(status_message, log=False)


class MatchNumberGeneratorUI:
    """User interface for numbering RANKING MATCH entries in ROS documents."""

    def __init__(
        self,
        parent: ttk.Frame,
        credentials_manager: GoogleDriveCredentialsManager,
        document_loader: ROSDocumentLoaderUI,
        match_schedule_importer: MatchScheduleImporterUI,
        console: ApplicationConsole,
    ) -> None:
        self.parent = parent
        self.credentials_manager = credentials_manager
        self.document_loader = document_loader
        self.match_schedule_importer = match_schedule_importer
        self.console = console
        self.current_document_var = tk.StringVar()
        self._status_var = tk.StringVar()
        self._default_status = "Load a ROS document and Google credentials to begin."

        self.document_loader.add_listener(self._on_document_url_changed)
        self.document_loader.add_name_listener(self._on_document_name_changed)
        self._on_document_url_changed(self.document_loader.get_document_url())
        self._on_document_name_changed(self.document_loader.get_document_name())

    def render(self, row: int) -> None:
        ttk.Label(self.parent, text="Current ROS Document:").grid(
            row=row, column=0, sticky="w", pady=(8, 0)
        )

        document_display = ttk.Entry(
            self.parent,
            textvariable=self.current_document_var,
            state="readonly",
        )
        document_display.grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        generate_button = ttk.Button(
            self.parent,
            text="Generate Match Numbers",
            command=self.generate_match_numbers,
        )
        generate_button.grid(row=row + 1, column=1, sticky="e", pady=8)

        status_label = ttk.Label(
            self.parent,
            textvariable=self._status_var,
            wraplength=520,
            justify="left",
        )
        status_label.grid(row=row + 2, column=0, columnspan=2, sticky="w")

        self.set_status(self._default_status, log=False)

    def _on_document_url_changed(self, url: str) -> None:
        if not url:
            self.set_status("Save a ROS document URL before generating match numbers.", log=False)

    def _on_document_name_changed(self, name: str) -> None:
        display_value = name if name else "Not set"
        self.current_document_var.set(display_value)

    def generate_match_numbers(self) -> None:
        document_url = self.document_loader.get_document_url()
        if not document_url:
            self.set_status("Save a ROS document URL before generating match numbers.")
            return

        if not self.match_schedule_importer.has_loaded_schedule():
            messagebox.showerror(
                "Match Schedule Required",
                "Import a match schedule JSON before generating match numbers.",
                parent=self.parent.winfo_toplevel(),
            )
            self.set_status("Import a match schedule JSON before generating match numbers.")
            return

        schedule_by_date = self.match_schedule_importer.get_matches_by_date_for_selected_field()
        if not schedule_by_date:
            field_number = self.match_schedule_importer.get_imported_field_number()
            if field_number is None:
                schedule_message = "The imported schedule does not include matches for the selected field."
            else:
                schedule_message = (
                    f"The imported schedule does not include matches for Field {field_number}."
                )
            messagebox.showerror(
                "Match Schedule Missing Matches",
                schedule_message,
                parent=self.parent.winfo_toplevel(),
            )
            self.set_status(schedule_message)
            return

        credentials, error = self.credentials_manager.get_valid_credentials()
        if credentials is None:
            if error:
                self.set_status(error)
            else:
                self.set_status("Load Google Drive credentials before running this tool.")
            return

        if build is None:
            self.set_status(
                "google-api-python-client is not installed. Install it with 'pip install google-api-python-client'."
            )
            return

        spreadsheet_id = extract_spreadsheet_id(document_url)
        if not spreadsheet_id:
            self.set_status("Unable to determine spreadsheet ID from the saved document URL.")
            return

        self.set_status("Contacting Google Sheets API...")

        def _worker() -> None:
            try:
                spreadsheet, _theme_supported, _service = _fetch_spreadsheet(
                    credentials, spreadsheet_id
                )
            except Exception as exc:  # pragma: no cover - network interaction
                message = f"Failed to update spreadsheet: {exc}"
                self.parent.after(0, lambda: self.set_status(message))
                return

            sheet_entries, existing_numbers_found = inspect_ranking_match_numbers(
                spreadsheet
            )
            spreadsheet_title = str(
                spreadsheet.get("properties", {}).get("title", "")
            )
            self.parent.after(
                0,
                lambda: self._handle_match_analysis(
                    credentials,
                    spreadsheet_id,
                    sheet_entries,
                    existing_numbers_found,
                    spreadsheet_title,
                    schedule_by_date,
                ),
            )

        threading.Thread(target=_worker, daemon=True).start()

    def set_status(self, message: str, *, log: bool = True) -> None:
        self._status_var.set(message)
        if log:
            self.console.log(f"[Match Number Generator] {message}")

    def _handle_match_analysis(
        self,
        credentials: Credentials,
        spreadsheet_id: str,
        sheet_entries: Sequence[Dict[str, Any]],
        existing_numbers_found: bool,
        spreadsheet_title: str,
        schedule_by_date: Sequence[Tuple[str, Sequence[Dict[str, Any]]]],
    ) -> None:
        if spreadsheet_title:
            self.document_loader.set_document_name(spreadsheet_title)

        total_matches = sum(len(entry.get("matches", [])) for entry in sheet_entries)
        diagnostics = self._gather_analysis_diagnostics(sheet_entries)

        if total_matches == 0:
            if diagnostics:
                self._log_diagnostics(diagnostics)
            self.set_status("No RANKING MATCH cells were updated.")
            return

        match_numbers_by_sheet, mismatch_error, alignment_notes = (
            self._derive_match_number_assignments(sheet_entries, schedule_by_date)
        )
        if mismatch_error:
            if diagnostics:
                self._log_diagnostics(diagnostics)
            messagebox.showerror(
                "Match Schedule Mismatch",
                mismatch_error,
                parent=self.parent.winfo_toplevel(),
            )
            self.console.log(f"[Match Number Generator] {mismatch_error}")
            self.set_status("Match schedule does not match the ROS document.")
            return

        for note in alignment_notes:
            self.console.log(f"[Match Number Generator] {note}")

        if existing_numbers_found:
            response = messagebox.askyesno(
                "Existing Match Numbers Found",
                "Existing Match Numbers found, do you want to renumber?",
                parent=self.parent.winfo_toplevel(),
            )
            if not response:
                if diagnostics:
                    self._log_diagnostics(diagnostics)
                self.console.log(
                    "[Match Number Generator] Existing match numbers left unchanged by user request."
                )
                self.set_status("Existing match numbers left unchanged.")
                return
            renumber_all = True
        else:
            renumber_all = False

        self._apply_match_numbers(
            credentials,
            spreadsheet_id,
            sheet_entries,
            renumber_all=renumber_all,
            initial_diagnostics=diagnostics,
            match_numbers_by_sheet=match_numbers_by_sheet,
        )

    def _derive_match_number_assignments(
        self,
        sheet_entries: Sequence[Dict[str, Any]],
        schedule_by_date: Sequence[Tuple[str, Sequence[Dict[str, Any]]]],
    ) -> Tuple[Dict[str, List[int]], Optional[str], List[str]]:
        sheet_match_entries = [
            entry for entry in sheet_entries if entry.get("matches")
        ]
        assignments: Dict[str, List[int]] = {}
        alignment_notes: List[str] = []

        schedule_count = len(schedule_by_date)
        sheet_count = len(sheet_match_entries)

        if schedule_count > sheet_count:
            extra_date, extra_matches = schedule_by_date[sheet_count]
            message = (
                f"The number of matches ({len(extra_matches)}) in the schedule for {extra_date} "
                "doesn't match the number of RANKING MATCH slots in "
                "N/A (no sheet available)."
            )
            return {}, message, []

        if sheet_count > schedule_count:
            extra_sheet = sheet_match_entries[schedule_count]
            sheet_name = str(extra_sheet.get("title", "Untitled"))
            message = (
                "The number of matches (0) in the schedule for N/A doesn't match the number "
                f"of RANKING MATCH slots in {sheet_name}."
            )
            return {}, message, []

        for (date, matches_for_date), sheet_entry in zip(
            schedule_by_date, sheet_match_entries
        ):
            sheet_name = str(sheet_entry.get("title", "Untitled"))
            sheet_matches: Sequence[Dict[str, Any]] = sheet_entry.get("matches", [])
            schedule_total = len(matches_for_date)
            slot_total = len(sheet_matches)
            alignment_notes.append(
                f"Verified {schedule_total} matches for {date} align with sheet {sheet_name}."
            )

            if schedule_total != slot_total:
                message = (
                    f"The number of matches ({schedule_total}) in the schedule for {date} "
                    f"doesn't match the number of RANKING MATCH slots in {sheet_name}."
                )
                return {}, message, []

            numbers: List[int] = []
            for match in matches_for_date:
                match_number = self.match_schedule_importer.extract_match_number(match)
                if match_number is None:
                    match_details = self.match_schedule_importer.describe_match(match)
                    message = (
                        "The imported schedule entry is missing a match number. "
                        f"Date: {date}. Match details: {match_details}. "
                        "Ensure this entry includes an 'id' or 'matchNumber' value."
                    )
                    return {}, message, []
                numbers.append(match_number)

            assignments[sheet_name] = numbers

        return assignments, None, alignment_notes

    def _apply_match_numbers(
        self,
        credentials: Credentials,
        spreadsheet_id: str,
        sheet_entries: Sequence[Dict[str, Any]],
        *,
        renumber_all: bool,
        initial_diagnostics: Sequence[str],
        match_numbers_by_sheet: Dict[str, Sequence[int]],
    ) -> None:
        self.set_status("Applying ranking match numbers...")

        def _worker() -> None:
            try:
                report, diagnostics = apply_ranking_match_number_updates(
                    credentials,
                    spreadsheet_id,
                    sheet_entries,
                    renumber_all=renumber_all,
                    initial_diagnostics=initial_diagnostics,
                    match_numbers_by_sheet=match_numbers_by_sheet,
                )
            except Exception as exc:  # pragma: no cover - network interaction
                message = f"Failed to update spreadsheet: {exc}"
                self.parent.after(0, lambda: self.set_status(message))
                return

            self.parent.after(
                0,
                lambda: self._handle_match_success(report, diagnostics),
            )

        threading.Thread(target=_worker, daemon=True).start()

    def _handle_match_success(
        self, report: Dict[str, List[str]], diagnostics: Sequence[str]
    ) -> None:
        if diagnostics:
            self._log_diagnostics(diagnostics)

        console_message = format_report(
            report,
            (),
            success_header="Ranking match numbers applied:",
            empty_message="No RANKING MATCH cells were updated.",
        )
        if report:
            total_updates = sum(len(entries) for entries in report.values())
            sheet_count = len(report)
            sheet_word = "sheet" if sheet_count == 1 else "sheets"
            status_message = (
                "Ranking match numbers applied to "
                f"{total_updates} cell{'s' if total_updates != 1 else ''} across "
                f"{sheet_count} {sheet_word}."
            )
            breakdowns = []
            for sheet, entries in report.items():
                entry_count = len(entries)
                cell_word = "cell" if entry_count == 1 else "cells"
                breakdowns.append(
                    f"{entry_count} {cell_word} applied on sheet ({sheet})"
                )
            if breakdowns:
                status_message += " " + ", ".join(breakdowns)
        else:
            status_message = "No RANKING MATCH cells were updated."

        self.console.log_info(f"[Match Number Generator] {console_message}")
        self.set_status(status_message, log=False)

    def _gather_analysis_diagnostics(
        self, sheet_entries: Sequence[Dict[str, Any]]
    ) -> List[str]:
        diagnostics: List[str] = []
        for entry in sheet_entries:
            diagnostics.extend(entry.get("diagnostics", []))
        return diagnostics

    def _log_diagnostics(self, diagnostics: Sequence[str]) -> None:
        for message in diagnostics:
            self.console.log(f"[Match Number Generator] {message}")


class FillScriptUI:
    """User interface for scanning Google Docs and ROS sheets for block markers."""

    def __init__(
        self,
        parent: ttk.Frame,
        credentials_manager: "GoogleDriveCredentialsManager",
        ros_document_loader: "ROSDocumentLoaderUI",
        console: ApplicationConsole,
    ) -> None:
        self.parent = parent
        self.credentials_manager = credentials_manager
        self.ros_document_loader = ros_document_loader
        self.console = console
        self.document_url_var = tk.StringVar()
        self._status_var = tk.StringVar()
        self._block_selection_var = tk.StringVar()
        self._ros_sheet_var = tk.StringVar()
        self._ros_block_selection_var = tk.StringVar()
        self._default_status = (
            "Paste a Google Docs link and press Read to scan for block markers."
        )
        self._available_blocks: List[int] = []
        self._ros_available_blocks: List[int] = []
        self.block_selector: Optional[ttk.Combobox] = None
        self.ros_sheet_selector: Optional[ttk.Combobox] = None
        self.ros_block_selector: Optional[ttk.Combobox] = None
        self._ros_sheet_values: Tuple[str, ...] = ()
        self._ros_spreadsheet_id: Optional[str] = None
        self._ros_loaded_spreadsheet_id: Optional[str] = None
        self._ros_sheet_map: Dict[str, Mapping[str, Any]] = {}
        self._ros_loading = False
        self._ros_spreadsheet: Optional[Mapping[str, Any]] = None
        self._document_id: Optional[str] = None
        self._pending_document_id: Optional[str] = None

        self.ros_document_loader.add_listener(self._on_ros_document_url_changed)
        self.credentials_manager.add_credentials_listener(self._on_credentials_changed)
        self._on_ros_document_url_changed(self.ros_document_loader.get_document_url())

    def render(self, row: int) -> None:
        self.parent.columnconfigure(1, weight=1)

        ttk.Label(self.parent, text="Google Docs link:").grid(
            row=row, column=0, sticky="w", pady=(8, 0)
        )

        link_entry = ttk.Entry(
            self.parent,
            textvariable=self.document_url_var,
        )
        link_entry.grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        read_button = ttk.Button(
            self.parent,
            text="Read",
            command=self.read_document,
        )
        read_button.grid(row=row + 1, column=1, sticky="e", pady=8)

        ttk.Label(self.parent, text="Available Blocks:").grid(
            row=row + 2, column=0, sticky="w"
        )

        self.block_selector = ttk.Combobox(
            self.parent,
            textvariable=self._block_selection_var,
            state="disabled",
            values=(),
        )
        self.block_selector.grid(
            row=row + 2, column=1, sticky="ew", padx=(8, 0), pady=(0, 8)
        )

        ttk.Label(self.parent, text="ROS Tab:").grid(
            row=row + 3, column=0, sticky="w"
        )

        self.ros_sheet_selector = ttk.Combobox(
            self.parent,
            textvariable=self._ros_sheet_var,
            state="disabled",
            values=self._ros_sheet_values,
        )
        self.ros_sheet_selector.grid(
            row=row + 3, column=1, sticky="ew", padx=(8, 0), pady=(0, 8)
        )
        self.ros_sheet_selector.bind("<<ComboboxSelected>>", self._on_ros_sheet_selected)

        ttk.Label(self.parent, text="ROS Available Blocks:").grid(
            row=row + 4, column=0, sticky="w"
        )

        self.ros_block_selector = ttk.Combobox(
            self.parent,
            textvariable=self._ros_block_selection_var,
            state="disabled",
            values=(),
        )
        self.ros_block_selector.grid(
            row=row + 4, column=1, sticky="ew", padx=(8, 0), pady=(0, 8)
        )

        generate_button = ttk.Button(
            self.parent,
            text="Generate Text for Block",
            command=self.generate_block_text,
        )
        generate_button.grid(row=row + 5, column=1, sticky="e", pady=(0, 8))

        status_label = ttk.Label(
            self.parent,
            textvariable=self._status_var,
            wraplength=520,
            justify="left",
        )
        status_label.grid(row=row + 6, column=0, columnspan=2, sticky="w")

        self.set_status(self._default_status, log=False)

    def read_document(self) -> None:
        url = self.document_url_var.get().strip()
        if not url:
            self.set_status("Enter a Google Docs link before reading.")
            return

        document_id = extract_document_id(url)
        if not document_id:
            self.set_status(
                "Unable to determine document ID from the provided link."
            )
            return

        self._pending_document_id = document_id
        self._document_id = None

        credentials, error = self.credentials_manager.get_valid_credentials()
        if credentials is None:
            if error:
                self.set_status(error)
            else:
                self.set_status(
                    "Load Google Drive credentials before reading the script document."
                )
            return

        if build is None:
            self.set_status(
                "google-api-python-client is not installed. Install it with 'pip install google-api-python-client'."
            )
            return

        self.set_status("Fetching Google Doc...")

        threading.Thread(
            target=self._read_document_worker,
            args=(credentials, document_id),
            daemon=True,
        ).start()

    def _read_document_worker(self, credentials: Credentials, document_id: str) -> None:
        try:
            service = build("drive", "v3", credentials=credentials, cache_discovery=False)
            data = (
                service.files()
                .export(fileId=document_id, mimeType="text/plain")
                .execute()
            )
        except Exception as exc:  # pragma: no cover - network interaction
            message = f"Failed to fetch Google Doc: {exc}"[:500]
            self.parent.after(0, lambda: self.set_status(message))
            return

        if isinstance(data, bytes):
            text = data.decode("utf-8", errors="replace")
        else:
            text = str(data)

        self.parent.after(0, lambda: self._handle_document_text(text, document_id))

    def _handle_document_text(self, text: str, document_id: str) -> None:
        if self._pending_document_id and document_id != self._pending_document_id:
            return

        self._document_id = document_id

        blocks = find_fill_script_blocks(text)
        self._available_blocks = blocks

        if blocks:
            self._update_block_selector(blocks)
            block_summary = ", ".join(str(number) for number in blocks)
            self.console.log(f"[Fill Script] Found block markers: {block_summary}")
            count = len(blocks)
            self.set_status(
                f"Found {count} block{'s' if count != 1 else ''} with matching start and end tags."
            )
        else:
            self._update_block_selector([])
            if text.strip():
                self.set_status("No matching block tags were found in the document.")
            else:
                self.set_status("The document was fetched but appears to be empty.")

    def _update_block_selector(self, blocks: Sequence[int]) -> None:
        if not self.block_selector:
            return

        if blocks:
            values = tuple(str(number) for number in blocks)
            self.block_selector.configure(values=values, state="readonly")
            self._block_selection_var.set(values[0])
        else:
            self.block_selector.configure(values=(), state="disabled")
            self._block_selection_var.set("")

    def set_status(self, message: str, *, log: bool = True) -> None:
        self._status_var.set(message)
        if log:
            self.console.log(f"[Fill Script] {message}")

    def _on_ros_document_url_changed(self, url: str) -> None:
        spreadsheet_id = extract_spreadsheet_id(url)
        if spreadsheet_id != self._ros_spreadsheet_id:
            self._ros_spreadsheet_id = spreadsheet_id
            self._ros_loaded_spreadsheet_id = None
            self._ros_sheet_map = {}
            self._ros_spreadsheet = None
            self._update_ros_sheet_selector(())
            self._update_ros_block_selector([])
        if spreadsheet_id:
            self._maybe_refresh_ros_tabs()

    def _on_credentials_changed(self, credentials: Optional[Credentials]) -> None:
        if credentials is None:
            return
        if (
            self._ros_spreadsheet_id
            and self._ros_loaded_spreadsheet_id != self._ros_spreadsheet_id
        ):
            self._maybe_refresh_ros_tabs()

    def _maybe_refresh_ros_tabs(self) -> None:
        if self._ros_loading:
            return
        spreadsheet_id = self._ros_spreadsheet_id
        if not spreadsheet_id:
            return
        if build is None:
            self.set_status(
                "google-api-python-client is not installed. Install it with 'pip install google-api-python-client'."
            )
            return
        credentials, error = self.credentials_manager.get_valid_credentials()
        if credentials is None:
            if error:
                self.set_status(error)
            else:
                self.set_status(
                    "Load Google Drive credentials before reading the ROS spreadsheet."
                )
            return

        self._ros_loading = True
        self._ros_spreadsheet = None
        self.set_status("Loading ROS tabs...")

        def _worker() -> None:
            try:
                spreadsheet, _theme_supported, _service = _fetch_spreadsheet(
                    credentials, spreadsheet_id
                )
            except Exception as exc:  # pragma: no cover - network interaction
                message = f"Failed to load ROS spreadsheet: {exc}"[:500]
                self.parent.after(
                    0,
                    lambda: self._handle_ros_spreadsheet_error(message, spreadsheet_id),
                )
                return

            self.parent.after(
                0,
                lambda: self._handle_ros_spreadsheet_success(
                    spreadsheet, spreadsheet_id
                ),
            )

        threading.Thread(target=_worker, daemon=True).start()

    def _handle_ros_spreadsheet_error(
        self, message: str, spreadsheet_id: str
    ) -> None:
        self._ros_loading = False
        if spreadsheet_id != self._ros_spreadsheet_id:
            return
        self._ros_spreadsheet = None
        self.set_status(message)

    def _handle_ros_spreadsheet_success(
        self, spreadsheet: Mapping[str, Any], spreadsheet_id: str
    ) -> None:
        self._ros_loading = False
        if spreadsheet_id != self._ros_spreadsheet_id:
            return

        self._ros_spreadsheet = spreadsheet
        sheets = []
        for sheet in spreadsheet.get("sheets", []):
            properties = sheet.get("properties", {}) if isinstance(sheet, Mapping) else {}
            title = str(properties.get("title", "")).strip()
            index = properties.get("index", 0)
            if title:
                sheets.append((index, title, sheet))

        sheets.sort(key=lambda item: item[0])
        titles = tuple(title for _index, title, _sheet in sheets)
        self._ros_sheet_map = {title: sheet for _index, title, sheet in sheets}
        self._ros_loaded_spreadsheet_id = spreadsheet_id
        self._update_ros_sheet_selector(titles)

        if titles:
            current = self._ros_sheet_var.get()
            if current not in self._ros_sheet_map:
                self._ros_sheet_var.set(titles[0])
            self._update_ros_blocks()
            count = len(titles)
            self.set_status(
                f"Loaded {count} ROS tab{'s' if count != 1 else ''}. Select a tab to view block markers.",
                log=False,
            )
        else:
            self.set_status("No tabs were found in the ROS spreadsheet.")

    def _update_ros_sheet_selector(self, titles: Sequence[str]) -> None:
        values = tuple(titles)
        self._ros_sheet_values = values
        if not self.ros_sheet_selector:
            return

        if values:
            self.ros_sheet_selector.configure(values=values, state="readonly")
            current = self._ros_sheet_var.get()
            if current not in values:
                self._ros_sheet_var.set(values[0])
        else:
            self.ros_sheet_selector.configure(values=(), state="disabled")
            self._ros_sheet_var.set("")

    def _update_ros_block_selector(self, blocks: Sequence[int]) -> None:
        self._ros_available_blocks = list(blocks)
        if not self.ros_block_selector:
            return

        if blocks:
            values = tuple(str(number) for number in blocks)
            self.ros_block_selector.configure(values=values, state="readonly")
            current = self._ros_block_selection_var.get()
            if current not in values:
                self._ros_block_selection_var.set(values[0])
        else:
            self.ros_block_selector.configure(values=(), state="disabled")
            self._ros_block_selection_var.set("")

    def _on_ros_sheet_selected(self, _event: Any) -> None:
        self._update_ros_blocks()

    def _update_ros_blocks(self) -> None:
        sheet_title = self._ros_sheet_var.get().strip()
        if not sheet_title:
            self._update_ros_block_selector([])
            return

        sheet = self._ros_sheet_map.get(sheet_title)
        if not sheet:
            self._update_ros_block_selector([])
            return

        sheet_data = sheet.get("data", []) if isinstance(sheet, Mapping) else []
        if not sheet_data:
            self._update_ros_block_selector([])
            self.set_status(
                f"{sheet_title}: Sheet data is empty; no ROS block markers found.",
                log=False,
            )
            return

        task_column = find_task_column(sheet_data)
        if task_column is None:
            self._update_ros_block_selector([])
            self.set_status(
                f"{sheet_title}: No TASK column found when scanning for ROS blocks.",
                log=False,
            )
            return

        starts: Dict[int, List[int]] = {}
        ends: Dict[int, List[int]] = {}
        for row_index, cell in _iter_column_cells(sheet_data, task_column):
            text = _extract_cell_text(cell)
            if not text:
                continue
            for match in BLOCK_START_PATTERN.finditer(text):
                number = int(match.group(1))
                starts.setdefault(number, []).append(row_index)
            for match in BLOCK_END_PATTERN.finditer(text):
                number = int(match.group(1))
                ends.setdefault(number, []).append(row_index)

        matching_numbers: List[int] = []
        for number in sorted(starts.keys() & ends.keys()):
            start_rows = starts[number]
            end_rows = ends[number]
            if any(start <= end for start in start_rows for end in end_rows):
                matching_numbers.append(number)

        self._update_ros_block_selector(matching_numbers)

        if matching_numbers:
            block_summary = ", ".join(str(number) for number in matching_numbers)
            self.console.log(
                f"[Fill Script] {sheet_title}: Found ROS block markers: {block_summary}"
            )
            count = len(matching_numbers)
            self.set_status(
                f"{sheet_title}: Found {count} ROS block{'s' if count != 1 else ''} with matching start and end tags.",
                log=False,
            )
        else:
            self.set_status(
                f"{sheet_title}: No matching ROS block tags were found in the TASK column.",
                log=False,
            )

    def generate_block_text(self) -> None:
        if self._ros_loading:
            self.set_status("Wait for the ROS spreadsheet to finish loading before generating text.")
            return

        doc_block_value = self._block_selection_var.get().strip()
        if not doc_block_value:
            self.set_status("Select a Google Doc block before generating text.")
            return

        try:
            doc_block_number = int(doc_block_value)
        except ValueError:
            self.set_status("Select a valid Google Doc block before generating text.")
            return

        sheet_title = self._ros_sheet_var.get().strip()
        if not sheet_title:
            self.set_status("Select a ROS tab before generating text.")
            return

        ros_block_value = self._ros_block_selection_var.get().strip()
        if not ros_block_value:
            self.set_status("Select a ROS block before generating text.")
            return

        try:
            ros_block_number = int(ros_block_value)
        except ValueError:
            self.set_status("Select a valid ROS block before generating text.")
            return

        if ros_block_number != doc_block_number:
            self.set_status(
                "Select matching block numbers in the Google Doc and ROS dropdowns before generating text."
            )
            return

        if doc_block_number not in self._available_blocks:
            self.set_status(
                f"Block {doc_block_number} is not available in the scanned Google Doc."
            )
            return

        if ros_block_number not in self._ros_available_blocks:
            self.set_status(
                f"Block {ros_block_number} is not available in ROS tab '{sheet_title}'."
            )
            return

        spreadsheet = self._ros_spreadsheet
        if spreadsheet is None:
            self.set_status("Load the ROS spreadsheet before generating text.")
            return

        script_lines, diagnostics = build_script_lines_for_block(
            spreadsheet, sheet_title, doc_block_number
        )

        if diagnostics:
            for message in diagnostics:
                self.console.log(f"[Fill Script] {message}")

        if not script_lines:
            self.set_status(
                f"No eligible rows were found for block {doc_block_number} on '{sheet_title}'."
            )
            return

        formatted_lines = [self._format_script_line(line) for line in script_lines]
        indented_output = "\n".join(
            f"  {line}" if line else "" for line in formatted_lines
        )
        log_message = (
            f"[Fill Script] Block {doc_block_number} text ({sheet_title}):"
        )
        if indented_output:
            log_message = f"{log_message}\n{indented_output}"
        self.console.log(log_message)

        self._apply_block_text_to_document(
            doc_block_number, script_lines, sheet_title
        )

    def _apply_block_text_to_document(
        self, block_number: int, script_lines: Sequence[ScriptLine], sheet_title: str
    ) -> None:
        document_id = self._document_id
        if not document_id:
            self.set_status(
                "Read the Google Doc before generating text so it can be updated automatically. "
                "The generated script text has been logged to the console.",
            )
            return

        credentials, error = self.credentials_manager.get_valid_credentials()
        if credentials is None:
            message = (
                error
                if error
                else "Load Google Drive credentials before updating the Google Doc."
            )
            self.set_status(
                f"{message} The generated script text has been logged to the console."
            )
            return

        if build is None:
            self.set_status(
                "google-api-python-client is not installed. Install it with 'pip install google-api-python-client'. "
                "The generated script text has been logged to the console.",
            )
            return

        normalized_lines = [
            line.text.replace("\r\n", "\n").replace("\r", "\n")
            for line in script_lines
        ]
        block_text = "\n".join(normalized_lines)
        line_ranges: List[Tuple[int, int, ScriptLine]] = []
        position = 0
        for index, line in enumerate(script_lines):
            text = normalized_lines[index]
            start = position
            end = start + len(text)
            line_ranges.append((start, end, line))
            position = end
            if index < len(normalized_lines) - 1:
                position += 1
        newline_added = False
        if block_text and not block_text.endswith("\n"):
            block_text += "\n"
            newline_added = True

        status_message = (
            f"Inserting generated script text for block {block_number} into the Google Doc..."
        )
        self.set_status(status_message, log=False)
        self.console.log(
            f"[Fill Script] Updating Google Doc block {block_number} with generated text from '{sheet_title}'."
        )

        def _worker() -> None:
            try:
                docs_service = build(
                    "docs", "v1", credentials=credentials, cache_discovery=False
                )
                document = (
                    docs_service.documents().get(documentId=document_id).execute()
                )
            except Exception as exc:  # pragma: no cover - network interaction
                message = f"Failed to load Google Doc for update: {exc}"[:500]

                def _handle_failure() -> None:
                    self.set_status(
                        f"{message} The generated script text has been logged to the console."
                    )

                self.parent.after(0, _handle_failure)
                return

            block_range = _locate_document_block_content_range(document, block_number)
            if block_range is None:

                def _handle_missing_block() -> None:
                    self.set_status(
                        f"Block {block_number} markers were not found or are mismatched in the Google Doc. "
                        "The generated script text has been logged to the console."
                    )

                self.parent.after(0, _handle_missing_block)
                return

            start_index, end_index = block_range
            requests: List[Dict[str, Any]] = []
            if end_index > start_index:
                requests.append(
                    {
                        "deleteContentRange": {
                            "range": {
                                "startIndex": start_index,
                                "endIndex": end_index,
                            }
                        }
                    }
                )
            style_requests: List[Dict[str, Any]] = []
            if block_text:
                requests.append(
                    {
                        "insertText": {
                            "location": {"index": start_index},
                            "text": block_text,
                        }
                    }
                )
                if line_ranges:
                    for index, (relative_start, relative_end, line) in enumerate(
                        line_ranges
                    ):
                        abs_start = start_index + relative_start
                        abs_end = start_index + relative_end
                        paragraph_end = abs_end
                        if index < len(line_ranges) - 1:
                            paragraph_end += 1
                        elif newline_added:
                            paragraph_end += 1

                        alignment = line.alignment.strip().upper()
                        if (
                            alignment
                            and alignment != "LEFT"
                            and paragraph_end > abs_start
                        ):
                            style_requests.append(
                                {
                                    "updateParagraphStyle": {
                                        "range": {
                                            "startIndex": abs_start,
                                            "endIndex": paragraph_end,
                                        },
                                        "paragraphStyle": {"alignment": alignment},
                                        "fields": "alignment",
                                    }
                                }
                            )

                        if abs_end > abs_start:
                            style_requests.append(
                                {
                                    "updateTextStyle": {
                                        "range": {
                                            "startIndex": abs_start,
                                            "endIndex": abs_end,
                                        },
                                        "textStyle": {"bold": bool(line.bold)},
                                        "fields": "bold",
                                    }
                                }
                            )
            if style_requests:
                requests.extend(style_requests)

            if not requests:

                def _handle_noop() -> None:
                    self.set_status(
                        f"Block {block_number} markers were found but no content needed to be inserted."
                    )

                self.parent.after(0, _handle_noop)
                return

            try:
                docs_service.documents().batchUpdate(
                    documentId=document_id, body={"requests": requests}
                ).execute()
            except Exception as exc:  # pragma: no cover - network interaction
                base_message = f"Failed to update Google Doc: {exc}"
                hint = ""
                error_text = str(exc).lower()
                if "insufficient authentication scopes" in error_text:
                    hint = (
                        " Your Google credentials do not grant access to Google Docs. "
                        "Open the Drive Credentials tab and click Authorize to sign in again."
                    )
                elif HttpError is not None and isinstance(exc, HttpError):
                    try:
                        details = getattr(exc, "error_details", None)
                        if isinstance(details, list):
                            for detail in details:
                                if (
                                    isinstance(detail, Mapping)
                                    and detail.get("reason") == "ACCESS_TOKEN_SCOPE_INSUFFICIENT"
                                ):
                                    hint = (
                                        " Your Google credentials do not grant access to Google Docs. "
                                        "Open the Drive Credentials tab and click Authorize to sign in again."
                                    )
                                    break
                    except Exception:
                        pass
                message = base_message[:500]
                if hint:
                    message = f"{message}{hint}"

                def _handle_update_failure() -> None:
                    self.set_status(
                        f"{message} The generated script text has been logged to the console."
                    )

                self.parent.after(0, _handle_update_failure)
                return

            success_message = (
                f"Inserted generated script text for block {block_number} into the Google Doc. "
                "The text is also logged in the console."
            )

            def _handle_success() -> None:
                self.console.log(
                    f"[Fill Script] Block {block_number} text inserted into the Google Doc from '{sheet_title}'."
                )
                self.set_status(success_message, log=False)

            self.parent.after(0, _handle_success)

        threading.Thread(target=_worker, daemon=True).start()

    @staticmethod
    def _format_script_line(line: ScriptLine) -> str:
        markers: List[str] = []
        alignment = line.alignment.strip().lower()
        if alignment and alignment != "left":
            markers.append(alignment.upper())
        if line.bold:
            markers.append("BOLD")

        prefix = f"[{', '.join(markers)}] " if markers else ""
        text = line.text
        if not text and prefix:
            return prefix.rstrip()
        return f"{prefix}{text}" if text else ""


class OptimizeTeamVideosUI:
    """User interface wrapper for the team video optimization tool."""

    def __init__(
        self,
        parent: ttk.Frame,
        credentials_manager: "GoogleDriveCredentialsManager",
        document_loader: "ROSDocumentLoaderUI",
        match_schedule_importer: MatchScheduleImporterUI,
        console: ApplicationConsole,
    ) -> None:
        self.parent = parent
        self.credentials_manager = credentials_manager
        self.document_loader = document_loader
        self.match_schedule_importer = match_schedule_importer
        self.console = console
        self.current_document_var = tk.StringVar()
        self._status_var = tk.StringVar()
        self._default_status = (
            "Load a ROS document, match schedule, and Google credentials to begin."
        )

        self.document_loader.add_listener(self._on_document_url_changed)
        self.document_loader.add_name_listener(self._on_document_name_changed)
        self._on_document_url_changed(self.document_loader.get_document_url())
        self._on_document_name_changed(self.document_loader.get_document_name())

    def render(self, row: int) -> None:
        ttk.Label(self.parent, text="Current ROS Document:").grid(
            row=row, column=0, sticky="w", pady=(8, 0)
        )

        document_display = ttk.Entry(
            self.parent,
            textvariable=self.current_document_var,
            state="readonly",
        )
        document_display.grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        generate_button = ttk.Button(
            self.parent,
            text="Optimize Team Videos",
            command=self.optimize_team_videos,
        )
        generate_button.grid(row=row + 1, column=1, sticky="e", pady=8)

        status_label = ttk.Label(
            self.parent,
            textvariable=self._status_var,
            wraplength=520,
            justify="left",
        )
        status_label.grid(row=row + 2, column=0, columnspan=2, sticky="w")

        self.set_status(self._default_status, log=False)

    def _on_document_url_changed(self, url: str) -> None:
        if not url:
            self.set_status(
                "Save a ROS document URL before optimizing team videos.", log=False
            )

    def _on_document_name_changed(self, name: str) -> None:
        display_value = name if name else "Not set"
        self.current_document_var.set(display_value)

    def optimize_team_videos(self) -> None:
        document_url = self.document_loader.get_document_url()
        if not document_url:
            self.set_status("Save a ROS document URL before optimizing team videos.")
            return

        if not self.match_schedule_importer.has_loaded_schedule():
            messagebox.showerror(
                "Match Schedule Required",
                "Import a match schedule JSON before optimizing team videos.",
                parent=self.parent.winfo_toplevel(),
            )
            self.set_status("Import a match schedule JSON before optimizing team videos.")
            return

        matches = self.match_schedule_importer.get_matches_for_selected_field()
        if not matches:
            messagebox.showerror(
                "Match Schedule Missing Matches",
                "The imported schedule does not include matches for the selected field.",
                parent=self.parent.winfo_toplevel(),
            )
            self.set_status(
                "The imported schedule does not include matches for the selected field."
            )
            return

        credentials, error = self.credentials_manager.get_valid_credentials()
        if credentials is None:
            if error:
                self.set_status(error)
            else:
                self.set_status(
                    "Load Google Drive credentials before running this tool."
                )
            return

        if build is None:
            self.set_status(
                "google-api-python-client is not installed. Install it with 'pip install google-api-python-client'."
            )
            return

        spreadsheet_id = extract_spreadsheet_id(document_url)
        if not spreadsheet_id:
            self.set_status(
                "Unable to determine spreadsheet ID from the saved document URL."
            )
            return

        self.set_status("Contacting Google Sheets API...")

        def _worker() -> None:
            try:
                report, diagnostics, spreadsheet_title = optimize_team_videos_for_sheet(
                    credentials,
                    spreadsheet_id,
                    matches,
                    self.match_schedule_importer,
                    console=self.console,
                )
            except Exception as exc:  # pragma: no cover - network interaction
                message = f"Failed to update spreadsheet: {exc}"
                self.parent.after(0, lambda: self.set_status(message))
                return

            self.parent.after(
                0,
                lambda: self._handle_optimize_success(
                    report, diagnostics, spreadsheet_title
                ),
            )

        threading.Thread(target=_worker, daemon=True).start()

    def set_status(self, message: str, *, log: bool = True) -> None:
        self._status_var.set(message)
        if log:
            self.console.log(f"[Optimize Team Videos] {message}")

    def _handle_optimize_success(
        self,
        report: Dict[str, List[str]],
        diagnostics: Sequence[str],
        spreadsheet_title: str,
    ) -> None:
        if spreadsheet_title:
            self.document_loader.set_document_name(spreadsheet_title)

        if diagnostics:
            self._log_diagnostics(diagnostics)

        console_message = format_report(
            report,
            (),
            success_header="Team video assignments applied:",
            empty_message="No TEAM VIDEO PLACEHOLDER entries were updated.",
        )

        if report:
            total_updates = sum(len(entries) for entries in report.values())
            sheet_count = len(report)
            sheet_word = "sheet" if sheet_count == 1 else "sheets"
            status_message = (
                "Updated "
                f"{total_updates} placeholder{'s' if total_updates != 1 else ''} across "
                f"{sheet_count} {sheet_word}."
            )
        else:
            status_message = "No TEAM VIDEO PLACEHOLDER entries were updated."

        self.console.log_info(f"[Optimize Team Videos] {console_message}")
        self.set_status(status_message, log=False)

    def _log_diagnostics(self, diagnostics: Sequence[str]) -> None:
        for message in diagnostics:
            self.console.log(f"[Optimize Team Videos] {message}")


def extract_document_id(url: str) -> str:
    """Extract the Google Docs document ID from a URL."""

    match = GOOGLE_DOCUMENT_ID_PATTERN.search(url)
    if match:
        return match.group(1)
    return ""


def find_fill_script_blocks(text: str) -> List[int]:
    """Return block numbers that include matching start and end markers."""

    starts: Dict[int, List[int]] = {}
    for match in BLOCK_START_PATTERN.finditer(text):
        number = int(match.group(1))
        starts.setdefault(number, []).append(match.start())

    ends: Dict[int, List[int]] = {}
    for match in BLOCK_END_PATTERN.finditer(text):
        number = int(match.group(1))
        ends.setdefault(number, []).append(match.start())

    matching_numbers: List[int] = []
    for number in sorted(starts.keys() & ends.keys()):
        start_positions = starts[number]
        end_positions = ends[number]
        if any(start < end for start in start_positions for end in end_positions):
            matching_numbers.append(number)

    return matching_numbers


def _iter_document_text_runs(
    elements: Sequence[Mapping[str, Any]]
) -> Iterable[Tuple[int, str]]:
    """Yield ``(start_index, text)`` pairs for text runs in a Docs response."""

    for element in elements:
        if not isinstance(element, Mapping):
            continue

        paragraph = element.get("paragraph")
        if isinstance(paragraph, Mapping):
            for run in paragraph.get("elements", []):
                if not isinstance(run, Mapping):
                    continue
                text_run = run.get("textRun")
                if not isinstance(text_run, Mapping):
                    continue
                text = text_run.get("content")
                if not isinstance(text, str) or not text:
                    continue
                start_index = run.get("startIndex")
                if not isinstance(start_index, int):
                    start_index = element.get("startIndex")
                if not isinstance(start_index, int):
                    continue
                yield start_index, text

        table = element.get("table")
        if isinstance(table, Mapping):
            for row in table.get("tableRows", []):
                if not isinstance(row, Mapping):
                    continue
                for cell in row.get("tableCells", []):
                    if not isinstance(cell, Mapping):
                        continue
                    cell_content = cell.get("content", [])
                    if isinstance(cell_content, list):
                        yield from _iter_document_text_runs(cell_content)

        table_of_contents = element.get("tableOfContents")
        if isinstance(table_of_contents, Mapping):
            toc_content = table_of_contents.get("content", [])
            if isinstance(toc_content, list):
                yield from _iter_document_text_runs(toc_content)


def _extract_document_text_and_index_map(
    document: Mapping[str, Any]
) -> Tuple[str, List[int]]:
    """Return document text and a map of character positions to Docs indexes."""

    body = document.get("body")
    if not isinstance(body, Mapping):
        return "", []

    content = body.get("content", [])
    if not isinstance(content, list):
        return "", []

    text_parts: List[str] = []
    index_map: List[int] = []

    for start_index, text in _iter_document_text_runs(content):
        text_parts.append(text)
        for offset, _character in enumerate(text):
            index_map.append(start_index + offset)

    return "".join(text_parts), index_map


def _char_index_to_doc_position(position: int, index_map: Sequence[int]) -> int:
    """Convert a character offset within text to a Google Docs index."""

    if not index_map:
        return 1
    if position <= 0:
        return index_map[0]
    if position >= len(index_map):
        return index_map[-1] + 1
    return index_map[position]


def _locate_document_block_content_range(
    document: Mapping[str, Any], block_number: int
) -> Optional[Tuple[int, int]]:
    """Return the Docs range that sits between the block markers."""

    text, index_map = _extract_document_text_and_index_map(document)
    if not text:
        return None

    start_pattern = re.compile(rf"\[Block\s+{block_number}\s+start\]", re.IGNORECASE)
    end_pattern = re.compile(rf"\[Block\s+{block_number}\s+end\]", re.IGNORECASE)

    start_match = start_pattern.search(text)
    if not start_match:
        return None

    end_match = end_pattern.search(text, start_match.end())
    if not end_match:
        return None

    content_start_pos = start_match.end()
    while content_start_pos < len(text) and text[content_start_pos] in ("\n", "\r"):
        content_start_pos += 1

    content_end_pos = end_match.start()
    if content_end_pos < content_start_pos:
        content_end_pos = content_start_pos

    start_index = _char_index_to_doc_position(content_start_pos, index_map)
    end_index = _char_index_to_doc_position(content_end_pos, index_map)

    return start_index, end_index


def extract_spreadsheet_id(url: str) -> str:
    """Extract the spreadsheet ID from a Google Sheets URL."""

    parts = url.split("/")
    if "spreadsheets" in parts:
        try:
            index = parts.index("d")
        except ValueError:
            pass
        else:
            if index + 1 < len(parts):
                return parts[index + 1]

    if "spreadsheets" in url and "#gid" in url:
        # Fall back to query parsing
        import urllib.parse

        parsed = urllib.parse.urlparse(url)
        path_parts = parsed.path.split("/")
        if "d" in path_parts:
            d_index = path_parts.index("d")
            if d_index + 1 < len(path_parts):
                return path_parts[d_index + 1]

    return ""


def derive_document_name(url: str) -> str:
    """Best-effort extraction of a human-friendly name from a Sheets URL."""

    if not url:
        return ""

    import urllib.parse

    parsed = urllib.parse.urlparse(url)
    path_parts = [part for part in parsed.path.split("/") if part]

    for part in reversed(path_parts):
        lowered = part.lower()
        if lowered in {"edit", "view", "copy"}:
            continue
        if lowered == "d":
            continue
        if lowered in {"spreadsheets", "file"}:
            continue
        return urllib.parse.unquote(part)

    query = urllib.parse.parse_qs(parsed.query)
    for key in ("name", "title", "resourcekey"):
        values = query.get(key)
        if values:
            return urllib.parse.unquote(values[0])

    fragment_query = urllib.parse.parse_qs(parsed.fragment)
    for key in ("name", "title"):
        values = fragment_query.get(key)
        if values:
            return urllib.parse.unquote(values[0])

    return ""


def _fetch_spreadsheet(
    credentials: Credentials, spreadsheet_id: str
) -> Tuple[Dict[str, Any], bool, Any]:
    """Retrieve spreadsheet metadata and return it with theme support info."""

    service = build("sheets", "v4", credentials=credentials)
    spreadsheets_resource = service.spreadsheets()
    base_fields = (
        "sheets("
        "properties(title,sheetId,index),"
        "data(rowData(values("
        "userEnteredValue,"
        "formattedValue,"
        "effectiveFormat(backgroundColor,backgroundColorStyle),"
        "userEnteredFormat(backgroundColor,backgroundColorStyle)"
        "))))"
    )
    fields_with_theme = base_fields + ",spreadsheetTheme"
    get_kwargs: Dict[str, Any] = {
        "spreadsheetId": spreadsheet_id,
        "includeGridData": True,
        "fields": fields_with_theme,
    }
    theme_supported = True
    if _method_supports_parameter(spreadsheets_resource.get, "supportsAllDrives"):
        get_kwargs["supportsAllDrives"] = True

    try:
        spreadsheet = spreadsheets_resource.get(**get_kwargs).execute()
    except Exception as exc:
        if HttpError is not None and isinstance(exc, HttpError):
            status = getattr(exc, "status_code", None) or getattr(exc, "resp", None)
            if hasattr(status, "status"):
                status = getattr(status, "status")
            message = str(exc)
            if status == 400 and "not supported for this document" in message.lower():
                raise ValueError(
                    "The selected file is stored in Office Compatibility mode. "
                    "Open it in Google Sheets and use File → Save as Google Sheets, "
                    "then try again."
                ) from exc
            if status == 400 and "spreadsheettheme" in message.lower():
                fallback_kwargs = dict(get_kwargs, fields=base_fields)
                spreadsheet = spreadsheets_resource.get(**fallback_kwargs).execute()
                theme_supported = False
            else:
                raise
        else:
            raise

    return spreadsheet, theme_supported, service


def _collect_sheet_rows(sheet_data: Sequence[dict]) -> List[Tuple[int, Dict[int, dict]]]:
    rows: Dict[int, Dict[int, dict]] = {}
    for grid_data in sheet_data:
        if not isinstance(grid_data, dict):
            continue
        start_row = grid_data.get("startRow", 0)
        start_column = grid_data.get("startColumn", 0)
        row_data = grid_data.get("rowData", [])
        for row_offset, row in enumerate(row_data or []):
            if not isinstance(row, dict):
                continue
            row_index = start_row + row_offset
            row_cells = rows.setdefault(row_index, {})
            values = row.get("values", [])
            for column_offset, cell in enumerate(values or []):
                column_index = start_column + column_offset
                if isinstance(cell, dict):
                    row_cells[column_index] = cell
    return sorted(rows.items())


def extract_script_resources_from_videos_tab(
    spreadsheet: Mapping[str, Any]
) -> Tuple[Dict[str, VideoScriptRow], Dict[str, VideoScriptRow], List[BoothInterviewRow], List[str]]:
    diagnostics: List[str] = []
    videos_sheet: Optional[Mapping[str, Any]] = None
    for sheet in spreadsheet.get("sheets", []):
        properties = sheet.get("properties", {}) if isinstance(sheet, Mapping) else {}
        title = str(properties.get("title", ""))
        if title.strip().lower() == "videos":
            videos_sheet = sheet
            break

    if videos_sheet is None:
        diagnostics.append("Videos tab not found in the ROS spreadsheet.")
        return {}, {}, [], diagnostics

    sheet_data = videos_sheet.get("data", []) if isinstance(videos_sheet, Mapping) else []
    if not sheet_data:
        diagnostics.append("Videos tab does not contain any data.")
        return {}, {}, [], diagnostics

    rows = _collect_sheet_rows(sheet_data)
    team_by_number: Dict[str, VideoScriptRow] = {}
    feature_by_number: Dict[str, VideoScriptRow] = {}
    booth_rows: List[BoothInterviewRow] = []

    for _row_index, columns in rows:
        team_number_text = _extract_cell_text(columns.get(0, {}))
        team_label = _extract_cell_text(columns.get(1, {}))
        team_script = _extract_cell_text(columns.get(5, {}))

        team_number = _normalize_video_number(team_number_text) if team_number_text else None
        if team_number and team_script:
            entry = VideoScriptRow(number=team_number, label=team_label or team_number_text, script=team_script)
            team_by_number.setdefault(team_number, entry)

        feature_number_text = _extract_cell_text(columns.get(8, {}))
        feature_label = _extract_cell_text(columns.get(9, {}))
        feature_script = _extract_cell_text(columns.get(14, {}))

        feature_number = _normalize_video_number(feature_number_text) if feature_number_text else None
        if feature_number and feature_script:
            entry = VideoScriptRow(
                number=feature_number,
                label=feature_label or feature_number_text,
                script=feature_script,
            )
            feature_by_number.setdefault(feature_number, entry)

        booth_key = _extract_cell_text(columns.get(16, {}))
        booth_script = _extract_cell_text(columns.get(19, {}))
        if booth_key and booth_script:
            normalized_key = _normalize_booth_key(booth_key)
            if normalized_key:
                booth_rows.append(
                    BoothInterviewRow(key=booth_key, script=booth_script, normalized_key=normalized_key)
                )

    if not team_by_number:
        diagnostics.append("Videos tab did not include any team video script entries (columns A/B/F).")
    if not feature_by_number:
        diagnostics.append("Videos tab did not include any feature video script entries (columns I/J/O).")
    if not booth_rows:
        diagnostics.append("Videos tab did not include any booth interview script entries (columns Q/T).")

    return team_by_number, feature_by_number, booth_rows, diagnostics


def extract_video_dataset_from_spreadsheet(
    spreadsheet: Mapping[str, Any]
) -> Tuple[VideoDataset, List[str]]:
    dataset = VideoDataset(entries=[], by_code={}, by_normalized_name={})
    diagnostics: List[str] = []

    videos_sheet: Optional[Mapping[str, Any]] = None
    for sheet in spreadsheet.get("sheets", []):
        properties = sheet.get("properties", {}) if isinstance(sheet, Mapping) else {}
        title = str(properties.get("title", ""))
        if title.lower() == "videos":
            videos_sheet = sheet
            break

    if videos_sheet is None:
        diagnostics.append("No sheet named 'Videos' was found in the spreadsheet.")
        return dataset, diagnostics

    sheet_data = videos_sheet.get("data", []) if isinstance(videos_sheet, Mapping) else []
    if not sheet_data:
        diagnostics.append("Videos sheet does not contain any data.")
        return dataset, diagnostics

    rows = _collect_sheet_rows(sheet_data)
    header_row_index: Optional[int] = None
    header_map: Dict[str, int] = {}
    for row_index, columns in rows:
        detected: Dict[str, int] = {}
        for column_index, cell in columns.items():
            text = _extract_cell_text(cell)
            if not text:
                continue
            normalized = _normalize_header_alias(text, uppercase=False)
            for key, aliases in VIDEO_SHEET_HEADER_ALIASES.items():
                if normalized in aliases:
                    detected[key] = column_index
        if "team" in detected and "value" in detected:
            header_map = detected
            header_row_index = row_index
            break

    if header_row_index is None:
        diagnostics.append("Videos sheet is missing 'Team' and 'Value' headers.")
        return dataset, diagnostics

    team_column = header_map["team"]
    value_column = header_map["value"]
    video_id_column = header_map.get("video_id")
    time_column = header_map.get("time")
    dataset.sheet_title = title

    # The videos sheet consistently stores the source video number immediately
    # to the left of the team column and the duration two columns to the right
    # of the team. Use those fixed relative positions rather than attempting to
    # detect header labels, which may repeat elsewhere in the sheet.
    dataset.video_number_column = team_column - 1 if team_column > 0 else None
    duration_candidate = team_column + 2
    dataset.duration_column = duration_candidate if duration_candidate >= 0 else None
    dataset.match_column = header_map.get("match")
    dataset.team_column = team_column

    processed_rows = 0

    for row_index, columns in rows:
        if row_index <= header_row_index:
            continue
        team_cell = columns.get(team_column)
        if not team_cell:
            continue
        team_text = _extract_cell_text(team_cell).strip()
        if not team_text:
            continue

        cleaned_team = strip_leading_flag_emoji(team_text)
        display_team = _strip_country_name_noise(cleaned_team)
        normalized_name = normalize_country_lookup_value(team_text)

        value_cell = columns.get(value_column)
        value = _extract_numeric_cell_value(value_cell or {}) if value_cell else None
        if value is None and value_cell is not None:
            fallback_text = _extract_cell_text(value_cell).strip()
            if fallback_text:
                try:
                    value = float(fallback_text.replace(",", ""))
                except ValueError:
                    pass

        video_id: Optional[str] = None
        if video_id_column is not None:
            video_cell = columns.get(video_id_column)
            if video_cell is not None:
                candidate = _extract_cell_text(video_cell).strip()
                if candidate:
                    video_id = candidate

        time_value: Optional[str] = None
        if time_column is not None:
            time_cell = columns.get(time_column)
            if time_cell is not None:
                candidate = _extract_cell_text(time_cell).strip()
                if candidate:
                    time_value = candidate

        video_number_value: Optional[str] = None
        video_number_text: Optional[str] = None
        if dataset.video_number_column is not None:
            number_cell = columns.get(dataset.video_number_column)
            if number_cell is not None:
                candidate = _extract_cell_text(number_cell).strip()
                if candidate:
                    video_number_text = candidate
                    digits = re.findall(r"\d+", candidate)
                    if digits:
                        video_number_value = digits[-1].lstrip("0") or "0"
                    else:
                        video_number_value = candidate

        duration_value: Optional[str] = None
        duration_text: Optional[str] = None
        if dataset.duration_column is not None:
            duration_cell = columns.get(dataset.duration_column)
            if duration_cell is not None:
                candidate = _extract_cell_text(duration_cell).strip()
                if candidate:
                    duration_text = candidate
                    duration_value = candidate

        entry = VideoEntry(
            team_name=display_team if display_team else (cleaned_team if cleaned_team else team_text),
            normalized_name=normalized_name,
            value=value,
            video_id=video_id,
            time=time_value,
            row_index=row_index,
            video_number=video_number_value,
            duration=duration_value,
            team_cell_text=team_text,
            video_number_cell_text=video_number_text,
            duration_cell_text=duration_text,
        )
        dataset.entries.append(entry)
        processed_rows += 1

        if normalized_name and normalized_name not in dataset.by_normalized_name:
            dataset.by_normalized_name[normalized_name] = entry

        secondary_normalized = _normalize_country_name(cleaned_team)
        if (
            secondary_normalized
            and secondary_normalized not in dataset.by_normalized_name
        ):
            dataset.by_normalized_name[secondary_normalized] = entry

        iso_candidate: Optional[str] = None
        paren_match = re.search(r"\(([A-Z]{3})\)", team_text)
        if paren_match:
            candidate = paren_match.group(1).upper()
            if candidate in COUNTRY_CODE_TO_INFO:
                iso_candidate = candidate
        if iso_candidate is None:
            for match in ISO_CODE_IN_TEXT_PATTERN.findall(team_text.upper()):
                candidate = match.upper()
                if candidate in COUNTRY_CODE_TO_INFO:
                    iso_candidate = candidate
                    break
        if iso_candidate is None:
            iso_candidate = lookup_country_code(team_text)
        if iso_candidate and iso_candidate not in dataset.by_code:
            dataset.by_code[iso_candidate] = entry

    diagnostics.append(
        f"Videos sheet: processed {processed_rows} row{'s' if processed_rows != 1 else ''}."
    )
    return dataset, diagnostics


def collect_team_video_slots(
    spreadsheet: Mapping[str, Any]
) -> Tuple[List[PlaceholderSlot], List[str]]:
    slots: List[PlaceholderSlot] = []
    diagnostics: List[str] = []

    for sheet in spreadsheet.get("sheets", []):
        if not isinstance(sheet, Mapping):
            continue
        properties = sheet.get("properties", {})
        title = str(properties.get("title", "Untitled"))
        sheet_data = sheet.get("data", [])
        if not sheet_data:
            continue

        task_column = find_task_column(sheet_data)
        if task_column is None:
            continue

        rows = _collect_sheet_rows(sheet_data)

        header_video_column: Optional[int] = None
        header_duration_column: Optional[int] = None

        for _row_index, columns in rows:
            for column_index, cell in columns.items():
                text = _extract_cell_text(cell)
                if not text:
                    continue
                normalized = _normalize_header_alias(text, uppercase=True)
                if (
                    header_video_column is None
                    and normalized in VIDEO_NUMBER_HEADER_ALIASES
                ):
                    header_video_column = column_index
                if (
                    header_duration_column is None
                    and normalized in DURATION_HEADER_ALIASES
                ):
                    header_duration_column = column_index
            if header_video_column is not None and header_duration_column is not None:
                break

        if header_video_column is None and task_column > 0:
            header_video_column = task_column - 1
        if header_duration_column is None and task_column > 1:
            header_duration_column = task_column - 2

        pending: List[Tuple[int, str]] = []
        for row_index, columns in rows:
            cell = columns.get(task_column)
            if not cell:
                continue
            text = _extract_cell_text(cell)
            if not text:
                continue
            normalized = text.strip()
            upper = normalized.upper()
            if normalized.upper().startswith("TEAM VIDEO PLACEHOLDER"):
                pending.append((row_index, normalized))
                continue

            if upper.startswith("RANKING MATCH"):
                match_number = _extract_match_number_from_text(normalized)
                if match_number is None:
                    diagnostics.append(
                        f"{title}: Unable to extract match number from cell value {normalized!r}."
                    )
                    pending.clear()
                    continue

                if not pending:
                    continue

                for placeholder_index, (placeholder_row, _placeholder_text) in enumerate(
                    pending
                ):
                    slots.append(
                        PlaceholderSlot(
                            sheet_title=title,
                            row_index=placeholder_row,
                            task_column=task_column,
                            match_number=match_number,
                            placeholder_index=placeholder_index,
                            video_number_column=header_video_column,
                            duration_column=header_duration_column,
                        )
                    )
                pending.clear()

        if pending:
            diagnostics.append(
                f"{title}: {len(pending)} TEAM VIDEO PLACEHOLDER row(s) without a following RANKING MATCH were ignored."
            )

    slots.sort(key=lambda slot: (slot.sheet_title, slot.row_index, slot.placeholder_index))
    return slots, diagnostics


def _extract_countries_from_match(match: Mapping[str, Any]) -> Tuple[List[str], List[str]]:
    raw_values: List[str] = []
    raw_seen: Set[str] = set()
    codes_seen: Set[str] = set()
    codes_ordered: List[str] = []

    def visit(value: Any, depth: int = 0) -> None:
        if depth > 6:
            return
        if isinstance(value, str):
            iso3_candidate = normalize_country_code(value)
            if iso3_candidate:
                stripped = value.strip()
                if stripped and stripped not in raw_seen:
                    raw_seen.add(stripped)
                    raw_values.append(stripped)
                if iso3_candidate not in codes_seen:
                    codes_seen.add(iso3_candidate)
                    codes_ordered.append(iso3_candidate)
            return
        if isinstance(value, Mapping):
            for key in (
                "country",
                "countryCode",
                "country_code",
                "countrycode",
                "teamCountry",
            ):
                if key in value:
                    visit(value[key], depth + 1)
            for sub_value in value.values():
                if isinstance(sub_value, (list, tuple, set, dict)):
                    visit(sub_value, depth + 1)
            return
        if isinstance(value, (list, tuple, set)):
            for item in value:
                visit(item, depth + 1)

    for key in (
        "countries",
        "countryCodes",
        "teams",
        "participants",
        "alliances",
        "blueAllianceTeams",
        "redAllianceTeams",
    ):
        if key in match:
            visit(match[key])

    if not codes_ordered:
        for value in match.values():
            if isinstance(value, (list, tuple, set)):
                visit(value)
            elif isinstance(value, Mapping):
                visit(value)
    return raw_values, codes_ordered


def build_match_country_map(
    matches: Sequence[Mapping[str, Any]],
    importer: MatchScheduleImporterUI,
) -> Tuple[Dict[int, List[str]], List[str]]:
    mapping: Dict[int, List[str]] = {}
    diagnostics: List[str] = []

    for match in matches:
        if not isinstance(match, Mapping):
            continue
        match_number = importer.extract_match_number(dict(match))
        if match_number is None:
            details = importer.describe_match(dict(match))
            diagnostics.append(
                f"Schedule entry missing match number. Details: {details}."
            )
            continue

        raw_values, countries = _extract_countries_from_match(match)
        console = getattr(importer, "console", None)
        if console is not None:
            prefix = "[Match Schedule Importer]"
            raw_formatted = (
                "[]" if not raw_values else f"[{', '.join(raw_values)}]"
            )
            normalized_formatted = (
                "[]" if not countries else f"[{', '.join(countries)}]"
            )
            display_names: List[str] = []
            for code in countries:
                display = get_country_display_name(code)
                display_names.append(display if display else f"(unrecognized {code})")
            display_formatted = (
                "[]" if not display_names else f"[{', '.join(display_names)}]"
            )
            console.log_info(
                f"{prefix} Match #{match_number}: raw country values -> {raw_formatted}"
            )
            console.log_info(
                f"{prefix} Match #{match_number}: normalized country codes -> {normalized_formatted}"
            )
            console.log_info(
                f"{prefix} Match #{match_number}: display country names -> {display_formatted}"
            )
        if not countries:
            details = importer.describe_match(dict(match))
            diagnostics.append(
                f"No country entries found for match #{match_number}. Details: {details}."
            )
            continue
        mapping[match_number] = countries

    return mapping, diagnostics


def find_video_entry_for_code(code: str, dataset: VideoDataset) -> Optional[VideoEntry]:
    code = code.upper()
    entry = dataset.by_code.get(code)
    if entry is not None:
        return entry

    candidate_names: List[str] = []
    display_name = get_country_display_name(code)
    if display_name:
        candidate_names.append(display_name)
    candidate_names.extend(ADDITIONAL_COUNTRY_ALIASES.get(code, ()))

    for name in candidate_names:
        normalized_variants = [
            normalize_country_lookup_value(name),
            _normalize_country_name(strip_leading_flag_emoji(name)),
        ]
        for normalized in normalized_variants:
            if not normalized:
                continue
            entry = dataset.by_normalized_name.get(normalized)
            if entry is not None:
                return entry

    for entry in dataset.entries:
        team_upper = entry.team_name.upper()
        if re.search(rf"\b{code}\b", team_upper):
            return entry

    return None


def _hungarian_algorithm(cost_matrix: Sequence[Sequence[float]]) -> List[int]:
    rows = len(cost_matrix)
    if rows == 0:
        return []
    cols = len(cost_matrix[0]) if cost_matrix else 0
    if cols < rows:
        raise ValueError("Cost matrix must have at least as many columns as rows.")

    u = [0.0] * (rows + 1)
    v = [0.0] * (cols + 1)
    p = [0] * (cols + 1)
    way = [0] * (cols + 1)

    for i in range(1, rows + 1):
        p[0] = i
        minv = [math.inf] * (cols + 1)
        used = [False] * (cols + 1)
        j0 = 0
        while True:
            used[j0] = True
            i0 = p[j0]
            delta = math.inf
            j1 = 0
            for j in range(1, cols + 1):
                if used[j]:
                    continue
                cur = cost_matrix[i0 - 1][j - 1] - u[i0] - v[j]
                if cur < minv[j]:
                    minv[j] = cur
                    way[j] = j0
                if minv[j] < delta:
                    delta = minv[j]
                    j1 = j
            if delta is math.inf:
                break
            for j in range(cols + 1):
                if used[j]:
                    u[p[j]] += delta
                    v[j] -= delta
                else:
                    minv[j] -= delta
            j0 = j1
            if p[j0] == 0:
                break
        while True:
            j1 = way[j0]
            p[j0] = p[j1]
            j0 = j1
            if j0 == 0:
                break

    assignment = [-1] * rows
    for j in range(1, cols + 1):
        if p[j] != 0:
            assignment[p[j] - 1] = j - 1
    return assignment


def compute_team_video_assignments(
    slots: Sequence[PlaceholderSlot],
    match_countries: Mapping[int, Sequence[str]],
    dataset: VideoDataset,
    *,
    console: Optional[ApplicationConsole] = None,
) -> Tuple[List[PlaceholderAssignment], List[str]]:
    diagnostics: List[str] = []
    slot_candidates: List[Tuple[PlaceholderSlot, List[str]]] = []

    logged_matches: Set[Tuple[str, int]] = set()
    unresolved_logged: Set[Tuple[str, int]] = set()
    missing_logged: Set[Tuple[str, int]] = set()

    for slot in slots:
        countries = match_countries.get(slot.match_number, [])
        if not countries:
            diagnostics.append(
                f"Match #{slot.match_number} on sheet {slot.sheet_title} is missing country data in the imported schedule."
            )
            continue

        normalized_codes: List[str] = []
        unresolved_codes: List[str] = []
        for code in countries:
            iso3_code = normalize_country_code(code)
            if iso3_code:
                normalized_codes.append(iso3_code)
            else:
                unresolved_codes.append(code)

        match_key = (slot.sheet_title, slot.match_number)

        if unresolved_codes and match_key not in unresolved_logged:
            formatted_unresolved = ", ".join(
                sorted({str(code).strip() or "(blank)" for code in unresolved_codes})
            )
            diagnostics.append(
                "Unrecognized country codes in schedule for match "
                f"#{slot.match_number}: {formatted_unresolved}."
            )
            unresolved_logged.add(match_key)

        normalized_codes = list(dict.fromkeys(normalized_codes))

        if not normalized_codes:
            diagnostics.append(
                f"Match #{slot.match_number} on sheet {slot.sheet_title} has country entries, but none could be normalized."
            )
            continue

        if match_key not in logged_matches:
            if console is not None:
                console.log_info(
                    "[Optimize Team Videos] Match #"
                    f"{slot.match_number} schedule countries: {_format_country_codes_for_log(normalized_codes)}"
                )
            logged_matches.add(match_key)

        valid_codes: List[str] = []
        missing_entries: List[str] = []
        missing_value_codes: List[str] = []
        for code in normalized_codes:
            entry = find_video_entry_for_code(code, dataset)
            if entry is None:
                missing_entries.append(code)
            elif entry.value is None:
                missing_value_codes.append(code)
            else:
                valid_codes.append(code)

        if missing_entries:
            formatted_missing = _format_country_codes_for_log(sorted(set(missing_entries)))
            if match_key not in missing_logged:
                diagnostics.append(
                    f"No videos found for countries {formatted_missing} in match #{slot.match_number}."
                )
                missing_logged.add(match_key)
            if console is not None:
                console.log_info(
                    "[Optimize Team Videos] Missing videos for match #"
                    f"{slot.match_number}: {formatted_missing}"
                )

        if not valid_codes:
            diagnostics.append(
                f"Match #{slot.match_number} on sheet {slot.sheet_title} has no assignable countries with videos."
            )
            continue

        if missing_value_codes:
            formatted_missing_value = _format_country_codes_for_log(
                sorted(set(missing_value_codes))
            )
            diagnostics.append(
                "Countries "
                f"{formatted_missing_value} in match #{slot.match_number} lack value scores and were ignored."
            )
            if console is not None:
                console.log_info(
                    "[Optimize Team Videos] Ignored countries without value scores for match #"
                    f"{slot.match_number}: {formatted_missing_value}"
                )

        slot_candidates.append((slot, valid_codes))

    if not slot_candidates:
        return [], diagnostics

    if console is not None:
        console.log_info(
            "[Optimize Team Videos] Slot candidates and available country videos:"
        )
        for slot, codes in slot_candidates:
            console.log_info(
                "  - Sheet="
                f"{slot.sheet_title!r}, match #{slot.match_number}, placeholder index {slot.placeholder_index}: "
                f"{', '.join(sorted(set(codes)))}"
            )

    unique_codes = sorted({code for _, codes in slot_candidates for code in codes})
    if console is not None:
        console.log_info(
            "[Optimize Team Videos] Unique country videos identified: "
            f"{', '.join(unique_codes) if unique_codes else '(none)'}"
        )
    if len(unique_codes) < len(slot_candidates):
        diagnostics.append(
            f"{len(slot_candidates)} placeholders reached assignment, but only {len(unique_codes)} unique country videos were available."
        )
        diagnostics.append(
            "Fewer unique country videos are available than placeholders; duplicates may be required."
        )

    inf_cost = 1e12
    duplicate_trigger_cost = 1e11
    cost_matrix: List[List[float]] = []

    column_codes: List[Optional[str]] = list(unique_codes)
    if len(column_codes) < len(slot_candidates):
        column_codes.extend([None] * (len(slot_candidates) - len(column_codes)))

    for slot, codes in slot_candidates:
        row: List[float] = []
        for column_code in column_codes:
            if column_code is None:
                row.append(duplicate_trigger_cost + slot.placeholder_index * 1e-6)
            elif column_code in codes:
                entry = find_video_entry_for_code(column_code, dataset)
                if entry is None or entry.value is None:
                    row.append(inf_cost)
                    continue
                value = float(entry.value)
                cost = value + slot.placeholder_index * 1e-6
                row.append(cost)
            else:
                row.append(inf_cost)
        cost_matrix.append(row)

    if not cost_matrix or not cost_matrix[0]:
        return [], diagnostics

    try:
        assignments = _hungarian_algorithm(cost_matrix)
    except ValueError:
        diagnostics.append(
            "Unable to compute assignments because there are fewer available countries than placeholders."
        )
        return [], diagnostics

    results_by_index: Dict[int, PlaceholderAssignment] = {}
    unmatched_indices: List[int] = []
    assigned_counts: Dict[str, int] = {}

    for row_index, (slot, _codes) in enumerate(slot_candidates):
        column_index = assignments[row_index] if row_index < len(assignments) else -1
        if column_index < 0 or column_index >= len(column_codes):
            diagnostics.append(
                f"No available country could be assigned to placeholder before match #{slot.match_number} on sheet {slot.sheet_title}."
            )
            unmatched_indices.append(row_index)
            continue
        cost = cost_matrix[row_index][column_index]
        if cost >= inf_cost:
            diagnostics.append(
                f"No valid assignment found for placeholder before match #{slot.match_number} on sheet {slot.sheet_title}."
            )
            unmatched_indices.append(row_index)
            continue
        code = column_codes[column_index]
        if code is None:
            unmatched_indices.append(row_index)
            continue
        entry = find_video_entry_for_code(code, dataset)
        if entry is None or entry.value is None:
            unmatched_indices.append(row_index)
            continue
        if assigned_counts.get(code, 0) > 0:
            unmatched_indices.append(row_index)
            continue
        assigned_counts[code] = assigned_counts.get(code, 0) + 1
        results_by_index[row_index] = PlaceholderAssignment(
            slot=slot,
            country_code=code,
            video=entry,
            dataset=dataset,
        )

    duplicate_messages: List[str] = []
    for row_index in unmatched_indices:
        slot, codes = slot_candidates[row_index]
        candidate_entries: List[Tuple[int, float, str, VideoEntry]] = []
        for code in codes:
            entry = find_video_entry_for_code(code, dataset)
            if entry is None or entry.value is None:
                continue
            assigned_count = assigned_counts.get(code, 0)
            candidate_entries.append((assigned_count, float(entry.value), code, entry))
        if not candidate_entries:
            diagnostics.append(
                f"Unable to assign any video to placeholder before match #{slot.match_number} on sheet {slot.sheet_title}, even after allowing duplicates."
            )
            continue
        candidate_entries.sort(key=lambda item: (item[0], item[1], item[2]))
        assigned_count, _value, code, entry = candidate_entries[0]
        assigned_counts[code] = assigned_count + 1
        if assigned_count > 0:
            duplicate_messages.append(
                f"Duplicate assignment: country {code} reused for match #{slot.match_number}."
            )
        results_by_index[row_index] = PlaceholderAssignment(
            slot=slot,
            country_code=code,
            video=entry,
            dataset=dataset,
        )

    duplicates_by_code: Dict[str, List[int]] = {}
    ordered_results: List[PlaceholderAssignment] = []
    for index, (slot, _codes) in enumerate(slot_candidates):
        assignment = results_by_index.get(index)
        if assignment is None:
            continue
        ordered_results.append(assignment)
        duplicates_by_code.setdefault(assignment.country_code, []).append(
            assignment.slot.match_number
        )

    duplicate_codes = {
        code: sorted(set(matches))
        for code, matches in duplicates_by_code.items()
        if len(matches) > 1
    }
    if duplicate_codes and console is not None:
        for code, matches in sorted(duplicate_codes.items()):
            match_text = ", ".join(str(match) for match in matches)
            console.log_warn(
                "[Optimize Team Videos] Duplicate video required: "
                f"{code} assigned to matches {match_text}."
            )
    if duplicate_messages:
        diagnostics.extend(sorted(duplicate_messages))

    return ordered_results, diagnostics


def apply_team_video_updates(
    credentials: Credentials,
    spreadsheet_id: str,
    assignments: Sequence[PlaceholderAssignment],
    *,
    console: Optional[ApplicationConsole] = None,
) -> Tuple[Dict[str, List[str]], List[str]]:
    updates: Dict[str, List[str]] = {}
    data_updates: List[Dict[str, Any]] = []
    match_numbers_by_entry: Dict[Tuple[str, int], List[int]] = {}
    entry_context_by_key: Dict[Tuple[str, int], Tuple[VideoEntry, VideoDataset]] = {}

    for assignment in assignments:
        slot = assignment.slot
        entry = assignment.video
        code = assignment.country_code
        dataset = assignment.dataset
        display_name = get_country_display_name(code)
        if not display_name:
            display_name = strip_leading_flag_emoji(entry.team_name)
        if not display_name:
            display_name = code
        new_text = display_name

        row_number = slot.row_index + 1
        task_cell = column_index_to_letter(slot.task_column) + str(row_number)
        sheet_title = slot.sheet_title
        sheet_updates = updates.setdefault(sheet_title, [])
        sheet_updates.append(f"{task_cell}: {new_text}")
        data_updates.append(
            {
                "range": single_cell_range(sheet_title, task_cell),
                "values": [[new_text]],
            }
        )

        source_sheet_title = dataset.sheet_title
        source_row_number = entry.row_index + 1
        source_team_column = dataset.team_column
        source_team_cell = (
            column_index_to_letter(source_team_column) + str(source_row_number)
            if source_team_column is not None
            else None
        )
        source_team_value = entry.team_cell_text or entry.team_name
        source_video_number_column = dataset.video_number_column
        source_video_number_cell = (
            column_index_to_letter(source_video_number_column) + str(source_row_number)
            if source_video_number_column is not None
            else None
        )
        source_duration_column = dataset.duration_column
        source_duration_cell = (
            column_index_to_letter(source_duration_column) + str(source_row_number)
            if source_duration_column is not None
            else None
        )
        source_video_number_value = (
            entry.video_number_cell_text
            if entry.video_number_cell_text is not None
            else entry.video_number or ""
        )
        source_duration_value = (
            entry.duration_cell_text
            if entry.duration_cell_text is not None
            else entry.duration or ""
        )

        entry_key = (dataset.sheet_title, entry.row_index)
        matches = match_numbers_by_entry.setdefault(entry_key, [])
        if slot.match_number not in matches:
            matches.append(slot.match_number)
        entry_context_by_key[entry_key] = (entry, dataset)

        video_number_value = (entry.video_number or "").strip()
        video_number_column = slot.video_number_column
        if video_number_column is None and slot.task_column > 0:
            video_number_column = slot.task_column - 1
        video_number_cell = (
            column_index_to_letter(video_number_column) + str(row_number)
            if video_number_column is not None
            else None
        )
        if console is not None:
            console.log_info(
                "[Apply Team Video Updates] Video number inputs: "
                f"sheet={sheet_title!r}, column={video_number_column}, cell={video_number_cell}, "
                f"value={video_number_value!r}; source_sheet={source_sheet_title!r}, source_row={source_row_number}, "
                f"source_team_cell={source_team_cell}, source_team_value={source_team_value!r}, "
                f"source_video_number_cell={source_video_number_cell}, source_video_number_value={source_video_number_value!r}, "
                f"normalized_source_video_number={entry.video_number!r}"
            )
        if video_number_column is not None and video_number_value:
            assert video_number_cell is not None
            sheet_updates.append(f"{video_number_cell}: Video Nº {video_number_value}")
            data_updates.append(
                {
                    "range": single_cell_range(sheet_title, video_number_cell),
                    "values": [[video_number_value]],
                }
            )

        duration_value = (entry.duration or "").strip()
        duration_column = slot.duration_column
        if duration_column is None and slot.task_column > 1:
            duration_column = slot.task_column - 2
        duration_cell = (
            column_index_to_letter(duration_column) + str(row_number)
            if duration_column is not None
            else None
        )
        if console is not None:
            console.log_info(
                "[Apply Team Video Updates] Duration inputs: "
                f"sheet={sheet_title!r}, column={duration_column}, cell={duration_cell}, "
                f"value={duration_value!r}; source_sheet={source_sheet_title!r}, source_row={source_row_number}, "
                f"source_team_cell={source_team_cell}, source_team_value={source_team_value!r}, "
                f"source_duration_cell={source_duration_cell}, source_duration_value={source_duration_value!r}, "
                f"normalized_source_duration={entry.duration!r}"
            )
        if duration_column is not None and duration_value:
            assert duration_cell is not None
            sheet_updates.append(f"{duration_cell}: Duration {duration_value}")
            data_updates.append(
                {
                    "range": single_cell_range(sheet_title, duration_cell),
                    "values": [[duration_value]],
                }
            )

    for entry_key, match_numbers in match_numbers_by_entry.items():
        entry, dataset = entry_context_by_key[entry_key]
        match_column = dataset.match_column
        if match_column is None:
            continue
        match_cell = column_index_to_letter(match_column) + str(entry.row_index + 1)
        unique_matches = sorted(set(match_numbers))
        match_value = ", ".join(str(number) for number in unique_matches)
        video_sheet_title = dataset.sheet_title
        updates.setdefault(video_sheet_title, []).append(
            f"{match_cell}: Match {match_value}"
        )
        data_updates.append(
            {
                "range": single_cell_range(video_sheet_title, match_cell),
                "values": [[match_value]],
            }
        )

    if data_updates:
        if build is None:
            raise ModuleNotFoundError(
                "google-api-python-client is not installed. Install it with 'pip install google-api-python-client'."
            )
        service = build("sheets", "v4", credentials=credentials)
        values_resource = service.spreadsheets().values()
        batch_update_kwargs = {
            "spreadsheetId": spreadsheet_id,
            "body": {
                "valueInputOption": "USER_ENTERED",
                "data": data_updates,
            },
        }
        if _method_supports_parameter(values_resource.batchUpdate, "supportsAllDrives"):
            batch_update_kwargs["supportsAllDrives"] = True
        values_resource.batchUpdate(**batch_update_kwargs).execute()

    formatted = {
        sheet: list(entries)
        for sheet, entries in sorted(updates.items())
    }

    return formatted, []


def optimize_team_videos_for_sheet(
    credentials: Credentials,
    spreadsheet_id: str,
    matches: Sequence[Mapping[str, Any]],
    importer: MatchScheduleImporterUI,
    *,
    console: Optional[ApplicationConsole] = None,
) -> Tuple[Dict[str, List[str]], List[str], str]:
    spreadsheet, _theme_supported, _service = _fetch_spreadsheet(
        credentials, spreadsheet_id
    )
    spreadsheet_title = str(
        spreadsheet.get("properties", {}).get("title", "")
    )

    diagnostics: List[str] = []
    active_console = console or getattr(importer, "console", None)

    video_dataset, video_diagnostics = extract_video_dataset_from_spreadsheet(spreadsheet)
    diagnostics.extend(video_diagnostics)

    if not video_dataset.entries:
        diagnostics.append("No video entries were found on the Videos sheet.")
        return {}, diagnostics, spreadsheet_title

    slots, slot_diagnostics = collect_team_video_slots(spreadsheet)
    diagnostics.extend(slot_diagnostics)

    if not slots:
        diagnostics.append("No TEAM VIDEO PLACEHOLDER entries were found.")
        return {}, diagnostics, spreadsheet_title

    match_countries, schedule_diagnostics = build_match_country_map(matches, importer)
    diagnostics.extend(schedule_diagnostics)

    if not match_countries:
        diagnostics.append("The imported match schedule did not include country assignments.")
        return {}, diagnostics, spreadsheet_title

    assignments, assignment_diagnostics = compute_team_video_assignments(
        slots, match_countries, video_dataset, console=active_console
    )
    diagnostics.extend(assignment_diagnostics)

    if not assignments:
        diagnostics.append(
            "Unable to assign videos to any placeholders with the available data."
        )
        return {}, diagnostics, spreadsheet_title

    report, update_diagnostics = apply_team_video_updates(
        credentials, spreadsheet_id, assignments, console=active_console
    )
    diagnostics.extend(update_diagnostics)

    return report, diagnostics, spreadsheet_title


def generate_placeholders_for_sheet(
    credentials: Credentials, spreadsheet_id: str
) -> Tuple[Dict[str, List[str]], List[str], str]:
    """Scan the spreadsheet and replace placeholder cells.

    Returns the applied updates together with diagnostics describing how the
    spreadsheet was inspected (for example, which column contained the ``TASK``
    header and sample values that began with a flag emoji), and the spreadsheet
    title.
    """

    spreadsheet, _theme_supported, service = _fetch_spreadsheet(
        credentials, spreadsheet_id
    )
    spreadsheet_title = str(
        spreadsheet.get("properties", {}).get("title", "")
    )

    updates: Dict[str, List[Tuple[str, str]]] = {}
    diagnostics: List[str] = [
        "Placeholder detection looks for rows whose TASK value begins with an emoji flag.",
    ]
    data_updates: List[Dict[str, Sequence[Sequence[str]]]] = []

    for sheet in spreadsheet.get("sheets", []):
        properties = sheet.get("properties", {})
        title = properties.get("title", "Untitled")

        if "OC" in title:
            diagnostics.append(f"{title}: Skipped because the sheet name contains 'OC'.")
            continue
        sheet_index = properties.get("index", 0)
        letter_prefix = chr(ord("A") + sheet_index)

        sheet_data = sheet.get("data", [])
        if not sheet_data:
            continue

        task_column = find_task_column(sheet_data)
        if task_column is None:
            diagnostics.append(f"{title}: No TASK column found within the first 10 rows.")
            continue

        column_letter = column_index_to_letter(task_column)
        diagnostics.append(
            f"{title}: TASK column located at index {task_column} (column {column_letter})."
        )

        matching_cells, existing_codes = find_placeholder_cells(
            sheet_data, task_column, diagnostics, title
        )
        if not matching_cells:
            diagnostics.append(
                f"{title}: No cells beginning with a flag emoji were found in column {column_letter}."
            )
            continue

        updates[title] = []
        code_iter = placeholder_code_iter(letter_prefix)

        highest_existing_index = _highest_placeholder_index(existing_codes, letter_prefix)
        if highest_existing_index is not None:
            code_iter = itertools.islice(
                code_iter,
                highest_existing_index + 1,
                None,
            )
        for row_index in matching_cells:
            code = next(code_iter)
            cell_a1 = column_index_to_letter(task_column) + str(row_index + 1)
            text = f"TEAM VIDEO PLACEHOLDER {code}"
            updates[title].append((cell_a1, text))
            data_updates.append(
                {
                    "range": single_cell_range(title, cell_a1),
                    "values": [[text]],
                }
            )

    if data_updates:
        values_resource = service.spreadsheets().values()
        batch_update_kwargs = {
            "spreadsheetId": spreadsheet_id,
            "body": {
                "valueInputOption": "USER_ENTERED",
                "data": data_updates,
            },
        }
        if _method_supports_parameter(values_resource.batchUpdate, "supportsAllDrives"):
            batch_update_kwargs["supportsAllDrives"] = True

        values_resource.batchUpdate(**batch_update_kwargs).execute()

    return (
        {
            sheet: [f"{cell}: {text}" for cell, text in entries]
            for sheet, entries in updates.items()
        },
        diagnostics,
        spreadsheet_title,
    )


def inspect_ranking_match_numbers(
    spreadsheet: Dict[str, Any]
) -> Tuple[List[Dict[str, Any]], bool]:
    """Return per-sheet match metadata and whether numbering already exists."""

    sheets: List[Dict[str, Any]] = []
    existing_numbers_found = False

    for sheet in spreadsheet.get("sheets", []):
        if not isinstance(sheet, dict):
            continue

        properties = sheet.get("properties", {})
        title = properties.get("title", "Untitled")
        index = int(properties.get("index", 0))
        sheet_data = sheet.get("data", [])

        entry: Dict[str, Any] = {
            "title": title,
            "index": index,
            "matches": [],
            "diagnostics": [],
        }

        if not sheet_data:
            entry["diagnostics"].append(f"{title}: No data available to inspect.")
            sheets.append(entry)
            continue

        matches: List[Dict[str, Any]] = []
        for grid_data in sheet_data:
            if not isinstance(grid_data, dict):
                continue
            start_row = grid_data.get("startRow", 0)
            start_column = grid_data.get("startColumn", 0)
            for row_offset, row in enumerate(grid_data.get("rowData", [])):
                if not isinstance(row, dict):
                    continue
                row_index = start_row + row_offset
                values = row.get("values", [])
                for column_offset, cell in enumerate(values):
                    if not isinstance(cell, dict):
                        continue
                    column_index = start_column + column_offset
                    cell_text = _extract_cell_text(cell)
                    if not cell_text:
                        continue
                    normalized = cell_text.strip()
                    upper_normalized = normalized.upper()
                    if upper_normalized == "RANKING MATCH" or upper_normalized.startswith(
                        "RANKING MATCH #"
                    ):
                        matches.append(
                            {
                                "row_index": row_index,
                                "column_index": column_index,
                                "original_text": normalized,
                            }
                        )
                        if upper_normalized.startswith("RANKING MATCH #"):
                            existing_numbers_found = True

        if matches:
            entry["matches"] = matches
        else:
            entry["diagnostics"].append(f"{title}: No RANKING MATCH cells found.")

        sheets.append(entry)

    return sheets, existing_numbers_found


def apply_ranking_match_number_updates(
    credentials: Credentials,
    spreadsheet_id: str,
    sheet_entries: Sequence[Dict[str, Any]],
    *,
    renumber_all: bool,
    initial_diagnostics: Optional[Sequence[str]] = None,
    match_numbers_by_sheet: Optional[Mapping[str, Sequence[int]]] = None,
) -> Tuple[Dict[str, List[str]], List[str]]:
    """Apply numbering updates to the provided match metadata."""

    updates: Dict[str, List[Tuple[str, str]]] = {}
    diagnostics: List[str] = list(initial_diagnostics or [])
    data_updates: List[Dict[str, Any]] = []

    counter = 1
    provided_numbers: Dict[str, List[int]] = (
        {
            sheet: [int(number) for number in numbers]
            for sheet, numbers in (match_numbers_by_sheet or {}).items()
        }
    )
    for entry in sorted(
        sheet_entries, key=lambda item: (int(item.get("index", 0)), item.get("title", ""))
    ):
        for message in entry.get("diagnostics", []):
            diagnostics.append(message)

        matches: Sequence[Dict[str, Any]] = entry.get("matches", [])
        if not matches:
            continue

        title = str(entry.get("title", "Untitled"))
        sheet_updates: List[Tuple[str, str]] = []
        assigned_numbers = provided_numbers.get(title)
        use_schedule_numbers = bool(assigned_numbers)

        if use_schedule_numbers and len(assigned_numbers) != len(matches):
            diagnostics.append(
                f"{title}: Provided match numbers ({len(assigned_numbers)}) do not match the number of slots ({len(matches)})."
            )
            continue

        sorted_matches = sorted(
            matches, key=lambda item: (item["row_index"], item["column_index"])
        )
        for idx, match in enumerate(sorted_matches):
            row_index = int(match["row_index"])
            column_index = int(match["column_index"])
            cell_a1 = column_index_to_letter(column_index) + str(row_index + 1)
            if use_schedule_numbers:
                match_number = assigned_numbers[idx]
            else:
                match_number = counter
                counter += 1
            new_text = f"RANKING MATCH #{match_number}"

            if renumber_all or match.get("original_text", "") != new_text:
                sheet_updates.append((cell_a1, new_text))
                data_updates.append(
                    {
                        "range": single_cell_range(title, cell_a1),
                        "values": [[new_text]],
                    }
                )

        diagnostics.append(
            f"{title}: Numbered {len(matches)} RANKING MATCH cell(s); "
            f"updated {len(sheet_updates)} of them."
        )

        if sheet_updates:
            updates[title] = sheet_updates

    if data_updates:
        if build is None:
            raise ModuleNotFoundError(
                "google-api-python-client is not installed. Install it with 'pip install google-api-python-client'."
            )

        service = build("sheets", "v4", credentials=credentials)
        values_resource = service.spreadsheets().values()
        batch_update_kwargs = {
            "spreadsheetId": spreadsheet_id,
            "body": {
                "valueInputOption": "USER_ENTERED",
                "data": data_updates,
            },
        }
        if _method_supports_parameter(values_resource.batchUpdate, "supportsAllDrives"):
            batch_update_kwargs["supportsAllDrives"] = True

        values_resource.batchUpdate(**batch_update_kwargs).execute()

    return (
        {
            sheet: [f"{cell}: {text}" for cell, text in entries]
            for sheet, entries in updates.items()
        },
        diagnostics,
    )


def generate_ranking_match_numbers(
    credentials: Credentials, spreadsheet_id: str, *, renumber_all: bool = False
) -> Tuple[Dict[str, List[str]], List[str]]:
    """Number every RANKING MATCH cell found in the spreadsheet."""

    spreadsheet, _theme_supported, _service = _fetch_spreadsheet(
        credentials, spreadsheet_id
    )
    sheet_entries, _existing_numbers = inspect_ranking_match_numbers(spreadsheet)
    initial_diagnostics: List[str] = []
    for entry in sheet_entries:
        initial_diagnostics.extend(entry.get("diagnostics", []))
    return apply_ranking_match_number_updates(
        credentials,
        spreadsheet_id,
        sheet_entries,
        renumber_all=renumber_all,
        initial_diagnostics=initial_diagnostics,
    )


def _extract_cell_text(cell: Dict[str, Any]) -> str:
    """Return the string contents of a cell, if available."""

    formatted = cell.get("formattedValue")
    if isinstance(formatted, str) and formatted.strip():
        return formatted.strip()

    user_entered = cell.get("userEnteredValue")
    if isinstance(user_entered, dict):
        string_value = user_entered.get("stringValue")
        if isinstance(string_value, str) and string_value.strip():
            return string_value.strip()

    effective = cell.get("effectiveValue")
    if isinstance(effective, dict):
        string_value = effective.get("stringValue")
        if isinstance(string_value, str) and string_value.strip():
            return string_value.strip()

    return ""

def find_task_column(sheet_data: Sequence[dict]) -> Optional[int]:
    """Locate the zero-based column index containing the TASK header."""

    for global_row_index in range(10):
        for grid_data in sheet_data:
            start_row = grid_data.get("startRow", 0)
            start_column = grid_data.get("startColumn", 0)
            relative_row_index = global_row_index - start_row
            row_data = grid_data.get("rowData", [])
            if relative_row_index < 0 or relative_row_index >= len(row_data):
                continue

            row = row_data[relative_row_index]
            for offset, cell in enumerate(row.get("values", [])):
                value = cell.get("formattedValue") or cell.get("userEnteredValue", {}).get("stringValue")
                if isinstance(value, str) and value.strip().upper() == "TASK":
                    return start_column + offset
    return None


def build_script_lines_for_block(
    spreadsheet: Mapping[str, Any], sheet_title: str, block_number: int
) -> Tuple[List[ScriptLine], List[str]]:
    diagnostics: List[str] = []
    script_lines: List[ScriptLine] = []

    target_sheet: Optional[Mapping[str, Any]] = None
    for sheet in spreadsheet.get("sheets", []):
        properties = sheet.get("properties", {}) if isinstance(sheet, Mapping) else {}
        title = str(properties.get("title", ""))
        if title.strip().lower() == sheet_title.strip().lower():
            target_sheet = sheet
            break

    if target_sheet is None:
        diagnostics.append(f"Sheet '{sheet_title}' was not found in the ROS spreadsheet.")
        return script_lines, diagnostics

    sheet_data = target_sheet.get("data", []) if isinstance(target_sheet, Mapping) else []
    if not sheet_data:
        diagnostics.append(f"Sheet '{sheet_title}' does not contain any data.")
        return script_lines, diagnostics

    task_column = find_task_column(sheet_data)
    if task_column is None:
        diagnostics.append(f"Sheet '{sheet_title}' is missing a TASK column.")
        return script_lines, diagnostics

    host_column = find_host_column(sheet_data, task_column)

    start_row: Optional[int] = None
    end_row: Optional[int] = None
    for row_index, cell in _iter_column_cells(sheet_data, task_column):
        text = _extract_cell_text(cell)
        if not text:
            continue
        for match in BLOCK_START_PATTERN.finditer(text):
            if int(match.group(1)) == block_number:
                start_row = row_index
        if start_row is not None:
            for match in BLOCK_END_PATTERN.finditer(text):
                if int(match.group(1)) == block_number and row_index >= start_row:
                    end_row = row_index
                    break
        if start_row is not None and end_row is not None:
            break

    if start_row is None or end_row is None or end_row <= start_row:
        diagnostics.append(
            f"Block {block_number} markers were not found or are mismatched in sheet '{sheet_title}'."
        )
        return script_lines, diagnostics

    rows = _collect_sheet_rows(sheet_data)
    team_lookup, feature_lookup, booth_rows, video_diagnostics = extract_script_resources_from_videos_tab(
        spreadsheet
    )
    diagnostics.extend(video_diagnostics)

    eligible_found = False
    fallback_host_numbers = itertools.cycle(("1", "2", "3"))

    for row_index, columns in rows:
        if row_index <= start_row or row_index >= end_row:
            continue

        task_cell = columns.get(task_column)
        task_text = _extract_cell_text(task_cell) if task_cell else ""
        task_upper = task_text.strip().upper()

        video_text = ""
        if task_column > 0:
            video_cell = columns.get(task_column - 1)
            video_text = _extract_cell_text(video_cell) if video_cell else ""

        host_text = ""
        if host_column is not None:
            host_cell = columns.get(host_column)
            host_text = _extract_cell_text(host_cell) if host_cell else ""
        else:
            fallback_cell = columns.get(task_column + 1)
            host_text = _extract_cell_text(fallback_cell) if fallback_cell else ""

        host_number = _extract_host_number(host_text)
        if not host_number:
            host_number = next(fallback_host_numbers)
        host_line_text = f"HOST {host_number}"

        if task_upper.startswith("RANKING MATCH"):
            match_number = _extract_match_number_from_text(task_text)
            if match_number is not None:
                header = f"<Ranking Match {match_number}>"
            else:
                diagnostics.append(
                    f"Row {row_index + 1}: Unable to parse ranking match number from '{task_text}'."
                )
                header = "<Ranking Match>"
            script_lines.append(ScriptLine(text=header))
            script_lines.append(ScriptLine(text=""))
            script_lines.append(ScriptLine(text="[Match Commentary]"))
            script_lines.append(ScriptLine(text=""))
            eligible_found = True
            continue

        if task_upper.startswith("FIELD INTERVIEW"):
            script_lines.append(ScriptLine(text="[Field Interview]"))
            script_lines.append(ScriptLine(text=""))
            eligible_found = True
            continue

        if task_upper.startswith("PIT INTERVIEW"):
            script_lines.append(ScriptLine(text="[Pit Interview]"))
            script_lines.append(ScriptLine(text=""))
            eligible_found = True
            continue

        if task_upper.startswith("BOOTH INTERVIEW"):
            match = re.match(r"BOOTH\s+INTERVIEW\s*[:\-]?\s*(.*)", task_text, flags=re.IGNORECASE)
            booth_label = match.group(1).strip() if match and match.group(1) else task_text.strip()
            display_label = booth_label if booth_label else "Booth Interview"
            script_lines.append(ScriptLine(text=f"[{display_label} - See Below]"))
            script_lines.append(ScriptLine(text=""))
            script_lines.append(ScriptLine(text=host_line_text, bold=True, alignment="center"))
            booth_entry = _match_booth_interview(booth_label, booth_rows)
            if booth_entry is None:
                diagnostics.append(
                    f"Row {row_index + 1}: Booth interview '{display_label}' not found in Videos tab."
                )
            else:
                script_lines.append(
                    ScriptLine(text=booth_entry.script.strip(), alignment="center")
                )
                script_lines.append(ScriptLine(text=""))
            eligible_found = True
            continue

        video_number = _normalize_video_number(video_text) if video_text else None
        if not video_number:
            continue

        entry = team_lookup.get(video_number)
        if entry is not None:
            script_lines.append(ScriptLine(text=host_line_text, bold=True, alignment="center"))
            script_lines.append(ScriptLine(text=entry.script.strip(), alignment="center"))
            script_lines.append(ScriptLine(text=""))
            script_lines.append(ScriptLine(text=f"<Team {entry.label} Video>"))
            script_lines.append(ScriptLine(text=""))
            eligible_found = True
            continue

        entry = feature_lookup.get(video_number)
        if entry is not None:
            script_lines.append(ScriptLine(text=host_line_text, bold=True, alignment="center"))
            script_lines.append(ScriptLine(text=entry.script.strip(), alignment="center"))
            script_lines.append(ScriptLine(text=""))
            script_lines.append(ScriptLine(text=f"<Feature Video {entry.label}>"))
            script_lines.append(ScriptLine(text=""))
            eligible_found = True
            continue

        diagnostics.append(
            f"Row {row_index + 1}: Video number '{video_text}' not found in Videos tab columns A or I."
        )

    if not eligible_found:
        diagnostics.append(
            f"Block {block_number} in sheet '{sheet_title}' did not contain any eligible rows."
        )

    return script_lines, diagnostics


def find_placeholder_cells(
    sheet_data: Sequence[dict],
    column_index: int,
    diagnostics: Optional[List[str]] = None,
    sheet_title: str = "",
) -> Tuple[List[int], Set[str]]:
    """Return placeholder rows and any existing placeholder codes in the column."""

    matches: List[int] = []
    existing_codes: Set[str] = set()
    seen_rows: set[int] = set()
    sample_count = 0

    for row_index, cell in _iter_column_cells(sheet_data, column_index):
        text = _extract_cell_text(cell)
        if not text:
            continue

        code = _extract_placeholder_code(text)
        if code:
            existing_codes.add(code)

        if not _starts_with_flag_emoji(text):
            continue

        if row_index in seen_rows:
            continue

        matches.append(row_index)
        seen_rows.add(row_index)

        if diagnostics is not None and sample_count < 5:
            preview = text.splitlines()[0][:40]
            diagnostics.append(
                f"{sheet_title}: Row {row_index + 1} flagged for placeholder replacement with value {preview!r}."
            )
            sample_count += 1

    return matches, existing_codes


def _starts_with_flag_emoji(text: str) -> bool:
    """Return True if *text* begins with an emoji flag after optional whitespace."""

    return bool(FLAG_EMOJI_PATTERN.match(text))


def _extract_placeholder_code(text: str) -> Optional[str]:
    """Extract an existing placeholder code from *text*, if present."""

    match = PLACEHOLDER_CODE_PATTERN.match(text.strip())
    if not match:
        return None
    return match.group(1).upper()


def _highest_placeholder_index(existing_codes: Set[str], prefix: str) -> Optional[int]:
    """Return the highest placeholder index encountered for *prefix*, if any."""

    highest: Optional[int] = None
    for code in existing_codes:
        if len(code) != 3 or not code.startswith(prefix):
            continue

        suffix = code[1:]
        if not suffix.isalpha():
            continue

        index = (ord(suffix[0]) - 65) * 26 + (ord(suffix[1]) - 65)
        if highest is None or index > highest:
            highest = index

    return highest


def _iter_column_cells(
    sheet_data: Sequence[dict], column_index: int
) -> Iterable[Tuple[int, dict]]:
    """Yield ``(row_index, cell)`` pairs for the specified column."""

    for grid_data in sheet_data:
        start_row = grid_data.get("startRow", 0)
        start_column = grid_data.get("startColumn", 0)
        cell_index = column_index - start_column
        if cell_index < 0:
            continue

        for offset, row in enumerate(grid_data.get("rowData", [])):
            row_index = start_row + offset
            values = row.get("values", [])
            if cell_index >= len(values):
                continue
            yield row_index, values[cell_index]


def placeholder_code_iter(prefix: str) -> Iterable[str]:
    """Yield placeholder codes with the given prefix."""

    def letters() -> Iterable[str]:
        for first in range(26):
            for second in range(26):
                yield chr(ord("A") + first) + chr(ord("A") + second)

    for suffix in letters():
        yield prefix + suffix


def column_index_to_letter(index: int) -> str:
    """Convert a zero-based column index to its Excel-style letter."""

    index += 1
    result = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def single_cell_range(sheet_title: str, cell_a1: str) -> str:
    """Return a properly quoted A1 range for a single cell."""

    escaped_title = sheet_title.replace("'", "''")
    return f"'{escaped_title}'!{cell_a1}:{cell_a1}"


def format_report(
    report: Dict[str, List[str]],
    diagnostics: Sequence[str] = (),
    *,
    success_header: str = "Placeholder updates applied:",
    empty_message: str = "No matching placeholders were found.",
) -> str:
    """Create a human-readable report of the updates performed."""

    lines: List[str] = []
    if report:
        if success_header:
            lines.append(success_header)
        for sheet, entries in report.items():
            lines.append(f"• {sheet}:")
            for entry in entries:
                lines.append(f"  - {entry}")
    else:
        lines.append(empty_message)

    if diagnostics:
        if lines:
            lines.append("")
        lines.append("Diagnostics:")
        for entry in diagnostics:
            lines.append(f"• {entry}")

    return "\n".join(lines)


def _method_supports_parameter(method: object, parameter: str) -> bool:
    """Return True if a Google API client method accepts the given parameter."""

    method_desc = getattr(method, "_methodDesc", {})
    if isinstance(method_desc, dict):
        parameters = method_desc.get("parameters", {})
        return isinstance(parameters, dict) and parameter in parameters
    return False


def main() -> None:
    """Run the application."""
    root = create_main_window()
    root.mainloop()


if __name__ == "__main__":
    main()
