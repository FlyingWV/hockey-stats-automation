import re
import time
import requests
import pandas as pd
import numpy as np
from bs4 import BeautifulSoup
from typing import Optional, Iterable, Tuple, List, Set, Dict

try:
    from lxml import html as LH  # fast HTML text extraction
except Exception:
    LH = None
import asyncio, os
import multiprocessing as mp

try:
    import aiohttp
except Exception:
    aiohttp = None
from collections import defaultdict, Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from functools import lru_cache
from pathlib import Path

SCRIPT_START_TIME = time.perf_counter()

# ---------------- CONFIG ----------------
# Create Results directory in current working directory
results_dir = Path.cwd() / "Results"
results_dir.mkdir(parents=True, exist_ok=True)


SEASON = "100"
START_GAME_NO = 1
END_GAME_NO = 47

input_file_path = results_dir / "NameList_Regular.CSV"  # (unused now)
output_excel_path = results_dir / f"Results_VHLM_{SEASON}_Playoffs.xlsx"

PARALLEL_WORKERS = 24
BATCH_SIZE = 120
MAX_RETRIES = 5
REQUEST_TIMEOUT = 20
STOP_AFTER_CONSEC_FAILS = 10

# Async + CPU pool tuning
ASYNC_MAX_CONNECTIONS = 32
CPU_WORKERS = max(1, (os.cpu_count() - 1 or 4))

SUCCESS_HITS_MODE = "half_substring"
STRICT_PBP_ONLY = True
NEAR_DUP_WINDOW = 3

DEBUG_GTG = False
DEBUG_GTG_CSV_PATH = results_dir / "GTG_Debug.csv"

TEAM_PAGE_MAX_ID = 60
ATTR_COLUMNS = [
    "CK",
    "FG",
    "DI",
    "SK",
    "ST",
    "EN",
    "DU",
    "PH",
    "FO",
    "PA",
    "SC",
    "DF",
    "PS",
    "EX",
    "LD",
    "PO",
    "MO",
    "OV",
]

# Goalie roster attributes
GOALIE_ATTR_COLUMNS = [
    "SK",
    "DU",
    "EN",
    "SZ",
    "AG",
    "RB",
    "SC",
    "HS",
    "RT",
    "PH",
    "PS",
    "EX",
    "LD",
    "PO",
    "MO",
    "OV",
]

# Trained ratings (from a player's FINAL page)
GOALIE_TRAINED_SKILL_COLUMNS = [
    "SK",
    "SZ",
    "AG",
    "RB",
    "SC",
    "HS",
    "RT",
    "PS",
    "EX",
    "LD",
]
GTR_PREFIX = "GTR_"
GTR_COLUMNS = [f"{GTR_PREFIX}{c}" for c in GOALIE_TRAINED_SKILL_COLUMNS]

TRAINED_SKILL_COLUMNS = [
    "DK",
    "SH",
    "PA",
    "BC",
    "GR",
    "FO",
    "PC",
    "DC",
    "OV",
    "SP",
    "SS",
    "WS",
    "LD",
    "FG",
    "PO",
    "EX",
]
TR_PREFIX = "TR_"
TR_COLUMNS = [f"{TR_PREFIX}{c}" for c in TRAINED_SKILL_COLUMNS]

# ---------------- DEBUG (PSA / PS%) ----------------
DEBUG_GOALIE_NAME = "Draw Mac"  # set "" to disable
DEBUG_PS_TRACE = True  # True -> attempt-level traces in parsers
DEBUG_PS_CSV_PATH = results_dir / "PS_Debug_DrawMac.csv"

# Results directory is already created above, no need for additional directory creation

# storage for per-minute rows
DEBUG_PS_ROWS: list[dict] = []


def _debug_ps_emit(
    game_no: int,
    url: str,
    label: str,
    att_map: Optional[Dict[str, int]],
    sv_map: Optional[Dict[str, int]],
):
    """
    Print + record PSA/PSSV for DEBUG_GOALIE_NAME from given maps under a phase label.
    """
    if not DEBUG_GOALIE_NAME:
        return
    a = int((att_map or {}).get(DEBUG_GOALIE_NAME, 0))
    s = int((sv_map or {}).get(DEBUG_GOALIE_NAME, 0))
    if a or s:
        print(
            f"[DEBUG PS] game {game_no:>3} {label:<12} :: {DEBUG_GOALIE_NAME} :: PSA={a}, PSSV={s}"
        )
    DEBUG_PS_ROWS.append(
        {
            "game_no": game_no,
            "url": url,
            "phase": label,
            "goalie": DEBUG_GOALIE_NAME,
            "psa": a,
            "pssv": s,
        }
    )


def _dbg_ps_trace(tag: str, goalie: str, saved: bool, line: str):
    """Unified, non-duplicative PS trace output."""
    if DEBUG_PS_TRACE and DEBUG_GOALIE_NAME and goalie == DEBUG_GOALIE_NAME:
        what = "+ATT, +SV" if saved else "+ATT"
        print(f"[TRACE {tag}] {goalie}: {what} :: {line}")


def game_no_from_url(u: str, season: str = SEASON) -> int:
    m = re.search(rf"/VHLM{season}-(\d+)\.html", u)
    return int(m.group(1)) if m else -1


# ---------------- HELPERS ----------------
WS_RE = re.compile(r"\s+")


def normalize_ws(s: str) -> str:
    return WS_RE.sub(" ", s).strip()


def make_name_pattern(player_name: str) -> str:
    esc = re.escape(player_name.strip())
    esc = esc.replace(r"\ ", r"\s+").replace(r"\.", r"\.?")
    return rf"\b{esc}\b"


@lru_cache(maxsize=2048)
def mmss_to_seconds(mmss: str) -> int:
    try:
        mm, ss = mmss.strip().split(":")
        return int(mm) * 60 + int(ss)
    except Exception:
        return 0


def seconds_to_mmss(total_seconds: int) -> str:
    if total_seconds <= 0:
        return "0:00"
    m, s = divmod(total_seconds, 60)
    return f"{m}:{s:02d}"


def extract_pos_categories(pos: str) -> Set[str]:
    if not isinstance(pos, str):
        return set()
    s = pos.lower()
    s = re.sub(r"[/,|]+", " ", s)
    tokens = re.findall(r"[a-z]+", s)

    cats: Set[str] = set()
    for t in tokens:
        if t in {"c", "cen", "cent", "center", "centre"}:
            cats.add("C")
        elif t in {"lw", "lwing", "leftwing", "leftwinger", "left"}:
            cats.add("LW")
        elif t in {"rw", "rwing", "rightwing", "rightwinger", "right"}:
            cats.add("RW")
        elif t in {
            "d",
            "def",
            "defn",
            "defense",
            "defence",
            "defender",
            "ld",
            "rd",
            "dl",
            "dr",
        }:
            cats.add("D")

    if not cats:
        if re.search(r"\bcenter|centre|\bc\b", s):
            cats.add("C")
        if "left wing" in s or "lw" in s:
            cats.add("LW")
        if "right wing" in s or "rw" in s:
            cats.add("RW")
        if "defen" in s or re.search(r"\b(d|ld|rd)\b", s):
            cats.add("D")
    return cats


@lru_cache(maxsize=50000)
def _canon(s: str) -> str:
    return canonical_name(s)


@lru_cache(maxsize=200000)
def canonical_name(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = re.sub(r"\([^)]*\)", "", s)
    s = normalize_ws(s).lower()
    return s


def fast_get(count_map: Dict[str, int], name: str) -> int:
    v = count_map.get(name)
    if v:
        return int(v)
    # rare fallback: canonical match only if exact failed
    cn = _canon(name)
    for k, val in count_map.items():
        if _canon(k) == cn:
            return int(val)
    return 0


def _looks_abbr(s: str) -> bool:
    return bool(re.fullmatch(r"[A-Z]{2,4}", (s or "").strip()))


# ---- Extract main PBP blocks (timestamped) ----
_TIME_PERIOD_CORE = r"(?P<time>\d{1,2}:\d{2})\s+of\s+(?P<period>(?:1st|2nd|3rd|OT\d*|OT|Overtime))\s+period\s*-\s*"
_NEXT_EVENT_AHEAD = (
    r"(?=\d{1,2}:\d{2}\s+of\s+(?:1st|2nd|3rd|OT\d*|OT|Overtime)\s+period\s*-\s*|$)"
)
PBP_BLOCK_RE = re.compile(
    _TIME_PERIOD_CORE + r"(?P<block>.*?)" + _NEXT_EVENT_AHEAD, re.I | re.DOTALL
)
_SENT_RE = re.compile(r"[^.]*\.")


def extract_pbp_blocks(text: str) -> Iterable[Tuple[str, str, str]]:
    for m in PBP_BLOCK_RE.finditer(text):
        yield m.group("time"), m.group("period"), m.group("block")


def iter_sentences(block: str) -> Iterable[str]:
    for m in _SENT_RE.finditer(block):
        sent = normalize_ws(m.group(0))
        if sent:
            yield sent


def extract_full_pbp_tail(full_text: str) -> str:
    m = re.search(r"\bFull\s+Play-?by-?Play\b", full_text, flags=re.IGNORECASE)
    return full_text[m.start() :] if m else full_text


FPP_TIMEHDR_RE = re.compile(_TIME_PERIOD_CORE, re.IGNORECASE)

# ---------------- GOALS & ASSISTS ----------------
GOAL_SUMMARY_RE = re.compile(
    r"""
    \b\d+\.\s*
    ([^,]+)\s*,\s*        # team
    ([^(]*?\S)\s+         # scorer
    (\d+)\s*              # number
    \(\s*([^)]*?)\s*\)\s* # helpers
    at\s+\d{1,2}:\d{2}
    (?:\s*\([^)]*\))*
    """,
    re.IGNORECASE | re.VERBOSE,
)
EXPLICIT_PP_PAT = re.compile(r"\(PP\)", re.IGNORECASE)
EXPLICIT_SH_PAT = re.compile(r"\(SH\)", re.IGNORECASE)


def _strip_trailing_number(name: str) -> str:
    return re.sub(r"\s+\d+\s*$", "", name).strip()


def parse_goals_assists_from_summary(
    full_text: str,
) -> Tuple[Dict[str, int], Dict[str, int]]:
    goals: Dict[str, int] = defaultdict(int)
    assists: Dict[str, int] = defaultdict(int)
    text = normalize_ws(full_text)
    for m in GOAL_SUMMARY_RE.finditer(text):
        scorer = normalize_ws(_strip_trailing_number(m.group(2)))
        helpers_raw = m.group(4)
        if scorer:
            goals[scorer] += 1
        if not helpers_raw or re.search(r"\bunassisted\b", helpers_raw, re.I):
            continue
        for tok in helpers_raw.split(","):
            nm = normalize_ws(_strip_trailing_number(tok))
            if nm:
                assists[nm] += 1
    return goals, assists


def parse_pp_from_scoring_summary(
    full_text: str,
) -> Tuple[Dict[str, int], Dict[str, int]]:
    ppg: Dict[str, int] = defaultdict(int)
    ppa: Dict[str, int] = defaultdict(int)
    text = normalize_ws(full_text)
    for m in GOAL_SUMMARY_RE.finditer(text):
        line = m.group(0)
        if not EXPLICIT_PP_PAT.search(line):
            continue
        scorer = normalize_ws(_strip_trailing_number(m.group(2)))
        helpers_raw = m.group(4)
        if scorer:
            ppg[scorer] += 1
        if helpers_raw and not re.search(r"\bunassisted\b", helpers_raw, re.I):
            for tok in helpers_raw.split(","):
                nm = normalize_ws(_strip_trailing_number(tok))
                if nm:
                    ppa[nm] += 1
    return ppg, ppa


def parse_sh_from_scoring_summary(
    full_text: str,
) -> Tuple[Dict[str, int], Dict[str, int]]:
    pkg: Dict[str, int] = defaultdict(int)
    pka: Dict[str, int] = defaultdict(int)
    text = normalize_ws(full_text)
    for m in GOAL_SUMMARY_RE.finditer(text):
        line = m.group(0)
        if not EXPLICIT_SH_PAT.search(line):
            continue
        scorer = normalize_ws(_strip_trailing_number(m.group(2)))
        helpers_raw = m.group(4)
        if scorer:
            pkg[scorer] += 1
        if helpers_raw and not re.search(r"\bunassisted\b", helpers_raw, re.I):
            for tok in helpers_raw.split(","):
                nm = normalize_ws(_strip_trailing_number(tok))
                if nm:
                    pka[nm] += 1
    return pkg, pka


# ---------------- 3 Stars ----------------
STAR_HDR_RE = re.compile(r"\b(?:3\s*Stars|Three\s*Stars)\b", re.IGNORECASE)
STAR_LINE_RE = re.compile(r"^\s*([123])\s*[-–—\.:]?\s*(.+?)\s*$")


def _strip_trailing_team_paren(s: str) -> str:
    return re.sub(r"\s*\([^)]+\)\s*$", "", s).strip()


def parse_three_stars(
    full_text_lines: str,
) -> Tuple[Dict[str, int], Dict[str, int], Dict[str, int]]:
    s1: Dict[str, int] = defaultdict(int)
    s2: Dict[str, int] = defaultdict(int)
    s3: Dict[str, int] = defaultdict(int)
    lines = full_text_lines.splitlines()
    in_block = False
    got_ranks: Set[int] = set()
    for raw in lines:
        line = raw.strip()
        if not in_block:
            if STAR_HDR_RE.search(line):
                in_block = True
            continue
        if len(got_ranks) >= 3:
            break
        m = STAR_LINE_RE.match(line)
        if not m:
            if re.search(
                r"(Players\s+Stats|Goalies\s+Stats|Team\s+Stats|Scoring\s+Summary|Penalties|Shots|Goals)\b",
                line,
                re.I,
            ):
                break
            continue
        try:
            rank = int(m.group(1))
        except Exception:
            continue
        name_rest = _strip_trailing_team_paren(m.group(2))
        name = normalize_ws(_strip_trailing_number(name_rest))
        if not name:
            continue
        if rank == 1 and 1 not in got_ranks:
            s1[name] += 1
            got_ranks.add(1)
        elif rank == 2 and 2 not in got_ranks:
            s2[name] += 1
            got_ranks.add(2)
        elif rank == 3 and 3 not in got_ranks:
            s3[name] += 1
            got_ranks.add(3)
    return s1, s2, s3


# ---------------- Period header spans ----------------
_PERIOD_HDR_RE = re.compile(
    r"""
    ^(?P<label>
        (?:1st|2nd|3rd)\s+period
        | OT(?:\d+)?(?:\s+period)?
        | Overtime(?:\s*\#\d+)?(?:\s+period)?
        | Shootout(?:\s*\#\d+)?(?:\s+round)?
        | SO(?:\s*\#\d+)?
    )\s*:?\s*$
    """,
    re.IGNORECASE | re.MULTILINE | re.VERBOSE,
)

# --- Goalie PIM from "Penalties :" sections (served by someone else still counts) ---
PENALTIES_HDR_RE = re.compile(r"^\s*Penalties\s*:\s*", re.I | re.M)
PEN_ENTRY_RE = re.compile(
    r"""
    (?P<name>[A-Za-z][A-Za-z .'\-]+?)        # penalized player
    \s*(?:\([A-Z]{2,4}\))?                   # optional (TEAM)
    \s+for\s+[^()]*\(\s*(?P<class>[^)]+)\s*\) # (Minor|Double Minor|Major|Misconduct|Game Misconduct|Match|...)
    \s+at\s*\d{1,2}:\d{2}                    # time
    """,
    re.I | re.X,
)


def _pen_minutes_from_class(cls: str) -> int:
    c = (cls or "").strip().lower()
    # order matters (double before single)
    if "double major" in c:
        return 10
    if "double minor" in c:
        return 4
    if "game misconduct" in c:
        return 10
    if "misconduct" in c:
        return 10
    if "match" in c:
        return 5
    if "major" in c:
        return 5
    if "minor" in c:
        return 2
    # safe default
    return 0


def _extract_penalty_sections(full_text_lines: str) -> list[tuple[int, int]]:
    """
    Return [(start, end), ...] spans for each 'Penalties :' block.
    Ends at the next period header or end of text.
    """
    spans = []
    for m in PENALTIES_HDR_RE.finditer(full_text_lines):
        start = m.end()
        # find the next period header after 'start'
        next_hdr = _PERIOD_HDR_RE.search(full_text_lines, pos=start)
        end = next_hdr.start() if next_hdr else len(full_text_lines)
        spans.append((start, end))
    return spans


def parse_goalie_pim_from_penalties(
    full_text_lines: str, goalie_canon_to_display: Dict[str, str]
) -> Dict[str, int]:
    """
    Scan 'Penalties :' blocks and attribute minutes to goalies present in this game.
    Keys in the returned dict are DISPLAY NAMES (matching the Goalie table).
    """
    out: Dict[str, int] = defaultdict(int)
    spans = _extract_penalty_sections(full_text_lines)
    if not spans:
        return {}

    goalie_canon_set = set(goalie_canon_to_display.keys())
    for start, end in spans:
        chunk = full_text_lines[start:end]
        for m in PEN_ENTRY_RE.finditer(chunk):
            raw_name = normalize_ws(m.group("name"))
            cls = m.group("class") or ""
            mins = _pen_minutes_from_class(cls)
            if mins <= 0:
                continue
            c = canonical_name(re.sub(r"\([^)]*\)", "", raw_name))
            if c in goalie_canon_set:
                disp = goalie_canon_to_display[c]
                out[disp] += mins
    return dict(out)


def __period_spans_for_summary(full_text: str) -> List[Tuple[int, int, str]]:
    spans = []
    matches = list(_PERIOD_HDR_RE.finditer(full_text))
    if not matches:
        return [(0, len(full_text), "ALL")]
    for i, m in enumerate(matches):
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(full_text)
        label = m.group("label").strip()
        spans.append((start, end, label))
    return spans


# --- Shootout winner inference (debug) ---
_ABBR_FROM_NAME_LINE_RE = re.compile(
    r"(?P<name>[A-Za-z][A-Za-z .'\-]+?)\s*\((?P<abbr>[A-Z]{2,4})\)"
)
_GOALIE_WIN_LINE_RE = re.compile(
    r"^\s*(?P<name>[^()\n]+?)\s*\((?P<abbr>[A-Z]{2,4})\)[^\n]*?\bW\b", re.MULTILINE
)


def _build_abbrev_to_team(
    full_text_lines: str, team_map_hint: Optional[Dict[str, str]]
) -> Dict[str, str]:
    out: Dict[str, str] = {}
    if not team_map_hint:
        return out
    for line in full_text_lines.splitlines():
        for mm in _ABBR_FROM_NAME_LINE_RE.finditer(line):
            nm = normalize_ws(mm.group("name"))
            ab = mm.group("abbr").upper()
            team = team_map_hint.get(nm)
            if team and ab not in out:
                out[ab] = team
    return out


def _infer_shootout_winner_team(
    full_text_lines: str, team_map_hint: Optional[Dict[str, str]]
) -> Optional[str]:
    m = _GOALIE_WIN_LINE_RE.search(full_text_lines or "")
    if not m:
        return None
    winner_abbr = m.group("abbr").upper()
    abbr_to_team = _build_abbrev_to_team(full_text_lines or "", team_map_hint or {})
    return abbr_to_team.get(winner_abbr)


# ---------------- GWG + GTG (debuggable) ----------------
def compute_gwg_gtg_and_debug(
    full_text: str,
    *,
    full_text_lines: Optional[str] = None,
    team_map_hint: Optional[Dict[str, str]] = None,
) -> Tuple[Dict[str, int], Dict[str, int], dict]:
    debug = {
        "went_to_ot": False,
        "went_to_so": False,
        "gwg_scorer": "",
        "gtg_scorer": "",
        "gtg_scorers": [],
        "gwg_idx": -1,
        "last_tie_idx": -1,
        "events_seq": [],
        "so_winner_team": "",
    }
    events: List[Tuple[str, str, int]] = []
    for m in GOAL_SUMMARY_RE.finditer(full_text):
        team = normalize_ws(m.group(1))
        scorer = normalize_ws(_strip_trailing_number(m.group(2)))
        if team and scorer:
            events.append((team, scorer, m.start()))
    if not events:
        return {}, {}, debug

    spans = __period_spans_for_summary(full_text)
    first_event_pos = min(pos for _, _, pos in events)
    regulation_end_pos = len(full_text)
    for s, _, lbl in spans:
        if s >= first_event_pos and re.search(
            r"\bOT\d*\b|\bOT\b|\bOvertime\b|\bShootout\b|\bSO\b", lbl, re.I
        ):
            regulation_end_pos = s
            break

    went_to_ot = any(pos >= regulation_end_pos for _, _, pos in events)
    went_to_so = bool(re.search(r"\bShootout\b|\bSO\b", full_text, re.I))
    debug["went_to_ot"] = went_to_ot
    debug["went_to_so"] = went_to_so

    seq: List[Tuple[str, str]] = [(t, s) for (t, s, _) in events]
    debug["events_seq"] = [f"{t} :: {s}" for (t, s) in seq]

    totals: Dict[str, int] = defaultdict(int)
    teams: List[str] = []
    for t, _ in seq:
        totals[t] += 1
        if t not in teams:
            teams.append(t)

    winner = max(totals.items(), key=lambda kv: kv[1])[0]
    loser_total = sum(v for k, v in totals.items() if k != winner)

    if went_to_so and full_text_lines:
        so_team = _infer_shootout_winner_team(full_text_lines, team_map_hint or {})
        if so_team:
            debug["so_winner_team"] = so_team

    gwg: Dict[str, int] = {}
    gwg_idx: Optional[int] = None
    if went_to_ot and not went_to_so:
        gwg_idx = len(seq) - 1
        gwg_scorer = seq[gwg_idx][1]
        gwg[gwg_scorer] = 1
        debug["gwg_idx"] = gwg_idx
        debug["gwg_scorer"] = gwg_scorer
    elif not went_to_ot and not went_to_so:
        if totals[winner] > loser_total:
            need = loser_total + 1
            cnt = 0
            for i, (t, s) in enumerate(seq):
                if t == winner:
                    cnt += 1
                    if cnt == need:
                        gwg_idx = i
                        gwg[s] = 1
                        debug["gwg_idx"] = i
                        debug["gwg_scorer"] = s
                        break

    score_before: List[Dict[str, int]] = []
    running: Dict[str, int] = defaultdict(int)
    for i, (t, _s) in enumerate(seq):
        score_before.append(dict(running))
        running[t] += 1

    def created_tie(i: int) -> bool:
        if len(teams) < 2:
            return False
        t, _s = seq[i]
        sb = score_before[i]
        a = sb.get(teams[0], 0)
        b = sb.get(teams[1], 0)
        if a == b:
            return False
        trailing_team = teams[0] if a < b else teams[1]
        return t == trailing_team and abs(a - b) == 1

    gtg: Dict[str, int] = {}
    if len(teams) >= 2:
        tie_all: List[int] = []
        tie_reg: List[int] = []
        score = {teams[0]: 0, teams[1]: 0}
        for i, (t, _s) in enumerate(seq):
            score[t] += 1
            if score[teams[0]] == score[teams[1]]:
                tie_all.append(i)
                if events[i][2] < regulation_end_pos:
                    tie_reg.append(i)

        if went_to_so:
            if tie_reg:
                j = tie_reg[-1]
                if created_tie(j):
                    gtg_scorer = seq[j][1]
                    gtg[gtg_scorer] = 1
                    debug["last_tie_idx"] = j
                    debug["gtg_scorer"] = gtg_scorer
                    debug["gtg_scorers"] = [gtg_scorer]
        elif went_to_ot:
            eventual_winner_team = winner
            if tie_reg:
                j = tie_reg[-1]
                if created_tie(j) and seq[j][0] == eventual_winner_team:
                    gtg_scorer = seq[j][1]
                    gtg[gtg_scorer] = 1
                    debug["last_tie_idx"] = j
                    debug["gtg_scorer"] = gtg_scorer
                    debug["gtg_scorers"] = [gtg_scorer]
        else:
            if gwg_idx is not None and tie_all:
                j = None
                for i in tie_all:
                    if i < gwg_idx:
                        j = i
                if j is not None and created_tie(j) and seq[j][0] == winner:
                    gtg_scorer = seq[j][1]
                    gtg[gtg_scorer] = 1
                    debug["last_tie_idx"] = j
                    debug["gtg_scorer"] = gtg_scorer
                    debug["gtg_scorers"] = [gtg_scorer]
    return gwg, gtg, debug


# ---------------- PLAYER STATS TABLE ----------------
TEAM_HEADER_RE = re.compile(r"Players\s+Stats\s+for\s+(?P<team>.+)", re.I)
PLAYER_ROW_RE = re.compile(
    r"""
    ^(?P<name>.*?)\s+
    (?P<G>-?\d+)\s+
    (?P<A>-?\d+)\s+
    (?P<P>-?\d+)\s+
    (?P<PM>-?\d+)\s+
    (?P<PIM>-?\d+)\s+
    (?P<S>\d+)\s+
    (?P<H>\d+)\s+
    (?P<SB>\d+)\s+
    (?P<GA>\d+)\s+
    (?P<TA>\d+)\s+
    (?P<FO>\d+/\d+)\s+
    (?P<MP>\d+:\d{2})\s+
    (?P<PP>\d+:\d{2})\s+
    (?P<PK>\d+:\d{2})\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)
HEADER_HINT_RE = re.compile(
    r"(Player\s+Name|G\s+A\s+P|Shots|H\s+SB|FO\s+MP|PP\s+MP|PK\s+MP)", re.I
)


def parse_player_stats_table(full_text_lines: str) -> Tuple[
    Dict[str, int],
    Dict[str, int],
    Dict[str, int],
    Dict[str, int],
    Dict[str, int],
    Dict[str, int],
    Dict[str, str],
    Dict[str, int],
    Dict[str, int],
]:
    plus_map: Dict[str, int] = {}
    hits_map: Dict[str, int] = {}
    pim_map: Dict[str, int] = {}
    shots_map: Dict[str, int] = {}
    sb_map: Dict[str, int] = {}
    mp_map: Dict[str, int] = {}
    team_map: Dict[str, str] = {}
    fow_map: Dict[str, int] = {}
    fot_map: Dict[str, int] = {}

    lines = full_text_lines.splitlines()
    in_section = False
    after_header = False
    current_team: Optional[str] = None

    for raw in lines:
        line = raw.rstrip()
        if not line.strip():
            continue

        mteam = TEAM_HEADER_RE.match(line)
        if mteam:
            in_section = True
            after_header = False
            current_team = normalize_ws(mteam.group("team"))
            continue

        if in_section and line.strip().startswith("-"):
            after_header = True
            continue

        if in_section and re.match(
            r"(Goalies\s+Stats|Team\s+Stats|Scoring\s+Summary|Players\s+Stats\s+for\s+.+)",
            line,
            re.I,
        ):
            m2 = TEAM_HEADER_RE.match(line)
            if m2:
                in_section = True
                after_header = False
                current_team = normalize_ws(m2.group("team"))
            else:
                in_section = False
                after_header = False
                current_team = None
            continue

        if not in_section or not after_header:
            continue

        if HEADER_HINT_RE.search(line):
            continue

        name = ""
        pm_val = h_val = pim_v = s_val = sb_v = 0
        mp_sec = 0
        fow = fot = 0

        mm = PLAYER_ROW_RE.match(line.strip())
        if mm:
            name = mm.group("name").strip()
            try:
                pm_val = int(mm.group("PM"))
                h_val = int(mm.group("H"))
                pim_v = int(mm.group("PIM"))
                s_val = int(mm.group("S"))
                sb_v = int(mm.group("SB"))
                mp_sec = mmss_to_seconds(mm.group("MP"))
                fo_str = mm.group("FO")
                if fo_str and re.fullmatch(r"\d+/\d+", fo_str):
                    fow, fot = map(int, fo_str.split("/"))
            except Exception:
                continue
        else:
            tokens = line.strip().split()
            if len(tokens) < 15:
                continue
            if any(
                tok.lower() in {"g", "a", "p", "shots", "sb", "fo", "mp", "pp", "pk"}
                for tok in tokens
            ):
                continue
            stats = tokens[-14:]
            name = " ".join(tokens[:-14])
            try:
                pm_val = int(stats[3])
                pim_v = int(stats[4])
                s_val = int(stats[5])
                h_val = int(stats[6])
                sb_v = int(stats[7])
                fo_tok = stats[10]
                if re.fullmatch(r"\d+/\d+", fo_tok):
                    fow, fot = map(int, fo_tok.split("/"))
                mp_sec = mmss_to_seconds(stats[11])
            except Exception:
                continue

        if not name:
            continue

        plus_map[name] = pm_val
        hits_map[name] = h_val
        pim_map[name] = pim_v
        shots_map[name] = s_val
        sb_map[name] = sb_v
        mp_map[name] = mp_sec
        fow_map[name] = fow
        fot_map[name] = fot
        if current_team:
            team_map[name] = current_team

    return (
        plus_map,
        hits_map,
        pim_map,
        shots_map,
        sb_map,
        mp_map,
        team_map,
        fow_map,
        fot_map,
    )


# ---------------- GOALIES STATS ----------------
GOALIES_GENERIC_HEADER_RE = re.compile(r"\bGoalie[s]?\s+Stats(?:istics)?\b", re.I)
SAVES_FROM_SHOTS_RE = re.compile(
    r"\b(?P<saves>\d+)\s+saves?\s+from\s+(?P<shots>\d+)\s+shots\b", re.IGNORECASE
)
GOALIES_TEAM_HEADER_RE = re.compile(r"Goalies\s+Stats\s+for\s+(?P<team>.+)", re.I)
GOALIE_LINE_RE = re.compile(
    r"^\s*(?P<name>[^()\n]+?)\s*\((?P<abbr>[A-Z]{2,4})\)\s*,?\s*(?P<rest>.*)$"
)
GOALIE_LINE_FALLBACK_RE = re.compile(
    r"^\s*(?P<name>[A-Za-z][A-Za-z .'\-]+?)(?:\s*\([A-Z]{2,4}\))?\s{2,}(?P<rest>.*)$"
)
_TIME_TOKEN_RE = re.compile(r"\b(?P<mp>\d{1,3}:\d{2})\b")
_INT_FIELD = lambda lab: re.compile(rf"\b{lab}\s*[:]\s*(-?\d+)\b", re.I)
_RATIO_FIELD = lambda lab: re.compile(rf"\b{lab}\s*[:]\s*(\d+)\s*/\s*(\d+)\b", re.I)
_FLOAT_FIELD = lambda lab: re.compile(rf"\b{lab}\s*[:]\s*([0-9]*\.?[0-9]+)\b", re.I)
GA_RE = _INT_FIELD("GA")
SA_RE = _INT_FIELD("SA")
SAVES_RE = _INT_FIELD("Saves?")
PIM_RE = _INT_FIELD("PIM")
A_RE = _INT_FIELD(r"A(?!T)")
EG_RE = _INT_FIELD(r"(?:EG|EN)")
SAR_RE = _INT_FIELD(r"SAR|Rebound(?:ed)?\s*Shots\s*Against")
PS_RATIO_RE = _RATIO_FIELD("PS")
PSA_RE = _INT_FIELD("PSA")
PSSV_RE = _INT_FIELD("PSSV|PSsv")
SVPCT_RE = _FLOAT_FIELD(r"(?:SV%|SVPCT|PCT)")
RESULT_TOKEN_RE = re.compile(r"\b(OTL|SOL|W|L)\b")


def parse_goalies_stats_table(full_text_lines: str):
    team_map: Dict[str, str] = {}
    mp_map: Dict[str, int] = {}
    ga_map: Dict[str, int] = {}
    sa_map: Dict[str, int] = {}
    pim_map: Dict[str, int] = {}
    w_map: Dict[str, int] = {}
    l_map: Dict[str, int] = {}
    otl_map: Dict[str, int] = {}
    so_map: Dict[str, int] = {}
    a_map: Dict[str, int] = {}
    eg_map: Dict[str, int] = {}
    ps_sv_map: Dict[str, int] = {}
    ps_att_map: Dict[str, int] = {}
    sar_map: Dict[str, int] = {}

    lines = full_text_lines.splitlines()
    in_section = False
    current_team: Optional[str] = None

    def flush_goalie(name: str, team: Optional[str], rest: str):
        name = normalize_ws(_strip_trailing_team_paren(name))
        if not name:
            return
        mp_sec = 0
        ga = None
        sa = None
        pim = 0
        a = 0
        eg = 0
        sar = 0
        ps_sv = None
        ps_att = None
        so_flag = 0
        result = None

        m = _TIME_TOKEN_RE.search(rest)
        if m:
            mp_sec = mmss_to_seconds(m.group("mp"))

        m = GA_RE.search(rest)
        ga = int(m.group(1)) if m else ga
        m = SA_RE.search(rest)
        sa = int(m.group(1)) if m else sa
        m = SAVES_RE.search(rest)
        _ = int(m.group(1)) if m else None
        m = PIM_RE.search(rest)
        pim = int(m.group(1)) if m else 0
        m = A_RE.search(rest)
        a = int(m.group(1)) if m else 0
        m = EG_RE.search(rest)
        eg = int(m.group(1)) if m else 0
        m = SAR_RE.search(rest)
        sar = int(m.group(1)) if m else 0

        m = SAVES_FROM_SHOTS_RE.search(rest)
        if m:
            saves = int(m.group("saves"))
            shots = int(m.group("shots"))
            sa = shots
            if ga is None:
                ga = shots - saves

        m = PS_RATIO_RE.search(rest)
        if m:
            ps_sv = int(m.group(1))
            ps_att = int(m.group(2))
        else:
            mA = PSA_RE.search(rest)
            mS = PSSV_RE.search(rest)
            if mA:
                ps_att = int(mA.group(1))
            if mS:
                ps_sv = int(mS.group(1))

        m = RESULT_TOKEN_RE.search(rest)
        if m:
            result = m.group(1).upper()

        if re.search(r"\bSO\b", rest) or (ga is not None and ga == 0):
            so_flag = 1

        if team:
            team_map[name] = normalize_ws(team)

        if mp_sec:
            mp_map[name] = mp_map.get(name, 0) + int(mp_sec)
        if ga is not None:
            ga_map[name] = ga_map.get(name, 0) + int(ga)
        if sa is not None:
            sa_map[name] = sa_map.get(name, 0) + int(sa)
        if pim:
            pim_map[name] = pim_map.get(name, 0) + int(pim)
        if a:
            a_map[name] = a_map.get(name, 0) + int(a)
        if eg:
            eg_map[name] = eg_map.get(name, 0) + int(eg)
        if sar:
            sar_map[name] = sar_map.get(name, 0) + int(sar)
        if ps_sv is not None:
            ps_sv_map[name] = ps_sv_map.get(name, 0) + int(ps_sv)
        if ps_att is not None:
            ps_att_map[name] = ps_att_map.get(name, 0) + int(ps_att)
        if so_flag:
            so_map[name] = so_map.get(name, 0) + 1

        if result == "W":
            w_map[name] = w_map.get(name, 0) + 1
        elif result in ("OTL", "SOL"):
            otl_map[name] = otl_map.get(name, 0) + 1
        elif result == "L":
            l_map[name] = l_map.get(name, 0) + 1

    for raw in lines:
        line = raw.rstrip()
        if not line.strip():
            continue

        mteam = GOALIES_TEAM_HEADER_RE.match(line)
        if mteam:
            in_section = True
            current_team = normalize_ws(mteam.group("team"))
            continue

        if not mteam and GOALIES_GENERIC_HEADER_RE.search(line):
            in_section = True
            current_team = None
            continue

        if in_section and re.match(
            r"(Players\s+Stats|Goalies?\s+Stats(?:istics)?\s+for\s+.+|Team\s+Stats|Scoring\s+Summary|Penalties|Shots|Goals)\b",
            line,
            re.I,
        ):
            m2 = GOALIES_TEAM_HEADER_RE.match(line)
            if m2:
                in_section = True
                current_team = normalize_ws(m2.group("team"))
            else:
                in_section = False
                current_team = None
            continue

        if not in_section:
            continue

        mm = GOALIE_LINE_RE.match(line)
        if mm:
            gname = mm.group("name")
            abbr = mm.group("abbr")
            rest = mm.group("rest") or ""
            team_to_use = current_team or (abbr if abbr else None)
            flush_goalie(gname, team_to_use, rest)
            continue

        mf = GOALIE_LINE_FALLBACK_RE.match(line)
        if mf:
            gname = mf.group("name")
            rest = mf.group("rest") or ""
            flush_goalie(gname, current_team, rest)

    return {
        "team_map": team_map,
        "mp_map": mp_map,
        "ga_map": ga_map,
        "sa_map": sa_map,
        "pim_map": pim_map,
        "w_map": w_map,
        "l_map": l_map,
        "otl_map": otl_map,
        "so_map": so_map,
        "a_map": a_map,
        "eg_map": eg_map,
        "ps_sv_map": ps_sv_map,
        "ps_att_map": ps_att_map,
        "sar_map": sar_map,
    }


# ---------------- PAGE-LEVEL PRECOMPUTATION ----------------
GAME_MARKER_RE = re.compile(
    r"\b\d{1,2}:\d{2}\s+of\s+(?:1st|2nd|3rd|OT\d*|OT|Overtime)\s+period\s*-\s*",
    re.IGNORECASE,
)
GIVEAWAY_RE = re.compile(
    r"^\s*(?P<victim>.+?)\s+is\s+hit\s+by\s+(?P<hitter>.+?)\s+and\s+loses(?:\s+the)?\s+puck\b",
    re.IGNORECASE,
)
HIT_ANY_RE = re.compile(
    r"^\s*(?P<victim>.+?)\s+is\s+hit\s+by\s+(?P<hitter>.+?)\b", re.IGNORECASE
)

SUCCESSFUL_HIT_RE = re.compile(
    r"^\s*(?P<victim>.+?)\s+is\s+hit\s+by\s+(?P<hitter>.+?)\s+and\s+loses(?:\s+the)?\s+puck\b",
    re.IGNORECASE,
)


def precompute_successful_hits_by_hitter(
    sents_tp: List[Tuple[str, str, str]],
) -> Dict[str, int]:
    """
    Counts 'Successful hits' by HITTER (not victim), deduped by (time, period, victim, hitter).
    (Includes hitter in signature to avoid cross-hitter collapses.)
    """
    seen: Set[Tuple[str, str, str, str]] = set()
    counts: Dict[str, int] = defaultdict(int)
    for t, p, sent in sents_tp:
        m = SUCCESSFUL_HIT_RE.search(sent)
        if not m:
            continue
        victim = normalize_ws(m.group("victim"))
        hitter = normalize_ws(m.group("hitter"))
        sig = (t, p.lower(), victim.lower(), hitter.lower())
        if sig in seen:
            continue
        seen.add(sig)
        counts[hitter] += 1
    return dict(counts)


def precompute_giveaways_by_victim(
    sents_tp: List[Tuple[str, str, str]],
) -> Dict[str, int]:
    seen = set()
    counts: Dict[str, int] = defaultdict(int)
    for t, p, sent in sents_tp:
        m = GIVEAWAY_RE.search(sent)
        if not m:
            continue
        victim = normalize_ws(m.group("victim"))
        hitter = normalize_ws(m.group("hitter"))
        sig = (t, p.lower(), hitter.lower())
        if sig in seen:
            continue
        seen.add(sig)
        counts[victim] += 1
    return counts


def precompute_hits_taken_by_victim(
    sents_tp: List[Tuple[str, str, str]],
) -> Dict[str, int]:
    seen = set()
    counts: Dict[str, int] = defaultdict(int)
    for t, p, sent in sents_tp:
        m = HIT_ANY_RE.search(sent)
        if not m:
            continue
        victim = normalize_ws(m.group("victim"))
        hitter = normalize_ws(m.group("hitter"))
        sig = (t, p.lower(), hitter.lower())
        if sig in seen:
            continue
        seen.add(sig)
        counts[victim] += 1
    return counts


RETRIEVE_ANY_RE = re.compile(
    r"\b(?:(?:Free|Loose)\s+)?Puck\s+Retr(?:ie|ei)ved\s+by\s+(?P<who>[^.]+?)(?=\.|$)",
    re.IGNORECASE,
)
INTERCEPT_BY_ANY_RE = re.compile(
    r"\bintercepted\s+by\s+([^.]+?)(?:\s+in\s+[A-Za-z '\-]+(?:\s+zone)?)?(?:\.|$)",
    re.IGNORECASE,
)

# --- Intercepted pass (credited against the passer) patterns ---
PASS_INT_AGAINST_PATTERNS = [
    re.compile(
        r"\bpass(?:\s+attempt|\s+attempted)?\s+by\s+(?P<passer>[^.]+?)\s+(?:is\s+)?intercepted\s+by\s+(?P<intcpt>[^.]+?)(?:\s+in\s+[A-Za-z '\-]+(?:\s+zone)?)?(?:\.|$)",
        re.IGNORECASE,
    ),
    re.compile(
        r"\b(?P<passer>[^.]+?)'?s\s+pass\s+(?:is\s+)?intercepted\s+by\s+(?P<intcpt>[^.]+?)(?:\s+in\s+[A-Za-z '\-]+(?:\s+zone)?)?(?:\.|$)",
        re.IGNORECASE,
    ),
    re.compile(
        r"\b(?P<passer>[^.]+?)\s+(?:tries|attempts)\s+to\s+pass\b[^.]*?\bintercepted\s+by\s+(?P<intcpt>[^.]+?)(?:\.|$)",
        re.IGNORECASE,
    ),
    re.compile(
        r"\b(?P<passer>[^.]+?)\s+passes\s+[^.]*?\bintercepted\s+by\s+(?P<intcpt>[^.]+?)(?:\.|$)",
        re.IGNORECASE,
    ),
]

REBOUND_SAVE_RE = re.compile(
    r"\bStopped\s+by\s+(?P<goalie>[A-Za-z][A-Za-z .'\-]+?)\s+with\s+a\s+rebound\b",
    re.IGNORECASE,
)

# --- Shootout PSA/PS%: header token to locate SO sections ---
SHOOTOUT_TOKEN_RE = re.compile(r"\bShoot\s*Out\b|\bShootout\b|\bSO\b", re.IGNORECASE)

# Require explicit "Round #n" markers for shootout parsing
SHOOTOUT_ROUND_RE = re.compile(r"\bRound\s*#?\s*\d+\b", re.IGNORECASE)

SO_PAIR_RE = re.compile(
    r"(?P<team>[^,:\-–—]+)\s*[,:\-–—]\s*(?P<shooter>[A-Za-z][A-Za-z .'\-]+)\.\s*",
    re.IGNORECASE,
)

# A single shootout attempt: "Team, Shooter." followed by outcome.
# Handles: "Stopped by <goalie>.", "Shot Misses the Net.", "Goal...", "Scores...", "Hits the Post."
SO_ATTEMPT_RE = re.compile(
    r"""
    (?P<team>[^,:\-–—]+)\s*[,:\-–—]\s*
    (?P<shooter>[A-Za-z][A-Za-z .'\-]+)\.\s*
    (?P<outcome>
        (?:Stopped\s+by\s+(?P<goalie>[^.]+)\.)
        |(?:Shot\s+Miss(?:es|ed)\s+(?:the\s+)?Net\.)
        |(?:Goal\b[^.]*\.)
        |(?:Scores?\b[^.]*\.)
        |(?:Hits?\s+the\s+Post\.)
    )
    """,
    re.IGNORECASE | re.VERBOSE | re.DOTALL,
)

# Optional "time of … -" prefix and flexible separators (-, – , —, :, ,)
_SO_PREFIX = r"(?:\d{1,2}:\d{2}\s+of\s+.*?[-–—]\s*)?"

# Anchored helpers (still used elsewhere sometimes)
SO_SHOOTER_LINE_RE = re.compile(
    rf"""^\s*{_SO_PREFIX}   # optional time header
         (?:Round\s*\d+\s*[:\-–—]\s*)?   # optional 'Round 1:'
         (?P<team>[^,:\-–—]+)\s*[,:\-–—]\s*(?P<shooter>[^.]+)\.\s*$""",
    re.IGNORECASE | re.VERBOSE,
)
SO_STOPPED_BY_RE = re.compile(
    rf"""^\s*{_SO_PREFIX}
         (?:Save|Saved|Stopped)\s+by\s+(?P<goalie>[^.]+)\.\s*$""",
    re.IGNORECASE | re.VERBOSE,
)
SO_MISSES_NET_RE = re.compile(
    rf"""^\s*{_SO_PREFIX}
         (?:Shot\s+)?Miss(?:es|ed)\s+the\s+Net\.?\s*$""",
    re.IGNORECASE | re.VERBOSE,
)
SO_GOAL_BY_RE = re.compile(
    rf"""^\s*{_SO_PREFIX}
         (?:Goal\s+by|Scores?\s+by)\s+(?P<shooter>[^.]+)\.?\s*$""",
    re.IGNORECASE | re.VERBOSE,
)

# Flexible outcome detectors used by the buffered parser
SO_SAVE_ANY_RE = re.compile(
    r"(?:Save|Saved|Stopped)\s+by\s+(?P<goalie>[^.]+)", re.IGNORECASE
)
SO_MISS_ANY_RE = re.compile(
    r"\b(?:Miss(?:es|ed)\s+(?:the\s+)?Net|Shot\s+Miss(?:es|ed)\s+(?:the\s+)?Net)\b",
    re.IGNORECASE,
)
SO_POST_ANY_RE = re.compile(r"\b(?:Hits?|Hit)\s+the\s+Post\b", re.IGNORECASE)
SO_GOAL_ANY_RE = re.compile(
    r"\b(?:Goal|Scores?)\b(?:\s+by\s+(?P<shooter>[^.]+))?(?:\s+on\s+(?P<goalie>[^.]+))?",
    re.IGNORECASE,
)

# Shooter line (may include outcome) — compiled at module scope
SHOOTER_ANY_LINE = re.compile(
    r"^\s*(?P<team>[^,:\-–—]+)\s*[,:\-–—]\s*(?P<shooter>[A-Za-z][A-Za-z .'\-]+)\s*(?:[.\-–—:]\s*(?P<rest>.*))?$",
    re.IGNORECASE,
)


def parse_shootout_ps_from_fpp(
    fpp_text: str,
    team_map_hint: Dict[str, str],
    goalie_team_map: Dict[str, str],
    goalie_mp_map: Dict[str, int],
    goalie_canon_to_display: Dict[str, str],
    goalie_w_map: Optional[Dict[str, int]] = None,
    goalie_l_map: Optional[Dict[str, int]] = None,
    goalie_otl_map: Optional[Dict[str, int]] = None,
) -> Tuple[Dict[str, int], Dict[str, int]]:
    """
    STRICT, sentence-aware FPP parser:
    - Only count PSA/PSSV that occur AFTER the first 'Round #n' marker inside a Shootout/SO block.
    - Pair each 'Team, Shooter.' with the *immediately following* outcome:
        Saved/Stopped by <goalie>  -> ATT + SV (that named goalie)
        Shot Misses the Net        -> ATT + SV (defending finishing goalie)
        Goal/Scores / Hits Post    -> ATT only  (defending finishing goalie)
    - Once a defending goalie is identified for a team, lock it for the rest of the shootout.
    - Emits [TRACE SO_FPP] lines when DEBUG_PS_TRACE is True for the configured DEBUG_GOALIE_NAME.
    """
    ps_att: Dict[str, int] = defaultdict(int)
    ps_sv: Dict[str, int] = defaultdict(int)

    goalie_w_map = goalie_w_map or {}
    goalie_l_map = goalie_l_map or {}
    goalie_otl_map = goalie_otl_map or {}

    teams_in_game_full = {normalize_ws(t) for t in (team_map_hint or {}).values() if t}
    abbr_to_team_full = _build_abbrev_to_team(fpp_text, team_map_hint or {})

    def norm_team(tok: Optional[str]) -> Optional[str]:
        if not tok:
            return None
        t = normalize_ws(tok)
        up = t.upper()
        if re.fullmatch(r"[A-Z]{2,4}", up) and up in abbr_to_team_full:
            return normalize_ws(abbr_to_team_full[up])
        for full in teams_in_game_full:
            if t == normalize_ws(full) or t.lower() in full.lower():
                return full
        return t

    def other_team(team: Optional[str]) -> Optional[str]:
        if not team or not teams_in_game_full:
            return None
        t_key = normalize_ws(team)
        for full in teams_in_game_full:
            if normalize_ws(full) != t_key:
                return full
        return None

    def finishing_goalie_for_team(
        team: Optional[str], locked: Dict[str, str]
    ) -> Optional[str]:
        """Prefer goalie with W/L/OTL; else nonzero MP; else max MP."""
        if not team:
            return None
        t = normalize_ws(team)
        if t in locked:
            return locked[t]
        candidates = [
            g for g, tm in (goalie_team_map or {}).items() if normalize_ws(tm) == t
        ]
        if not candidates:
            return None

        def has_result(g: str) -> int:
            return int(
                goalie_w_map.get(g, 0)
                + goalie_l_map.get(g, 0)
                + goalie_otl_map.get(g, 0)
                > 0
            )

        def mp(g: str) -> int:
            return int(goalie_mp_map.get(g, 0))

        return max(candidates, key=lambda g: (has_result(g), mp(g) > 0, mp(g)))

    found_any_attempt = False

    # Walk each Shootout/SO block in the FPP tail
    for m_so in SHOOTOUT_TOKEN_RE.finditer(fpp_text):
        block = fpp_text[m_so.end() :]

        # Require a Round marker in this block
        m_first_round = SHOOTOUT_ROUND_RE.search(block)
        if not m_first_round:
            continue

        # Only consider text from the first "Round #n" onward
        segment = block[m_first_round.start() :]

        # Lock the defending goalie per team within this shootout
        locked_goalie_by_team: Dict[str, str] = {}

        idx = 0
        while True:
            m_pair = SO_PAIR_RE.search(segment, idx)
            if not m_pair:
                break

            shooting_team = norm_team(m_pair.group("team"))
            def_team = other_team(shooting_team)

            # Slice to the outcome chunk (up to next pair/next round/end)
            after = m_pair.end()
            m_next_pair = SO_PAIR_RE.search(segment, after)
            m_next_round = SHOOTOUT_ROUND_RE.search(segment, after)
            cut = len(segment)
            if m_next_pair:
                cut = min(cut, m_next_pair.start())
            if m_next_round:
                cut = min(cut, m_next_round.start())
            chunk = segment[after:cut].strip()

            # Outcome 1: explicit save line
            m_save = SO_SAVE_ANY_RE.search(chunk)
            if m_save:
                found_any_attempt = True
                raw_g = normalize_ws(m_save.group("goalie"))
                gdisp = goalie_canon_to_display.get(canonical_name(raw_g), raw_g)
                ps_att[gdisp] += 1
                ps_sv[gdisp] += 1
                if DEBUG_PS_TRACE and DEBUG_GOALIE_NAME and gdisp == DEBUG_GOALIE_NAME:
                    print(
                        f"[TRACE SO_FPP] {DEBUG_GOALIE_NAME}: +ATT, +SV :: {m_pair.group(0).strip()}Stopped by {gdisp}."
                    )
                if def_team:
                    locked_goalie_by_team[normalize_ws(def_team)] = gdisp

            else:
                # Outcome 2/3/4: Miss / Post / Goal — need defending goalie
                gdisp = finishing_goalie_for_team(def_team, locked_goalie_by_team)
                if gdisp and def_team:
                    locked_goalie_by_team[normalize_ws(def_team)] = gdisp

                if gdisp:
                    if SO_MISS_ANY_RE.search(chunk):
                        found_any_attempt = True
                        ps_att[gdisp] += 1
                        ps_sv[gdisp] += 1
                        if (
                            DEBUG_PS_TRACE
                            and DEBUG_GOALIE_NAME
                            and gdisp == DEBUG_GOALIE_NAME
                        ):
                            print(
                                f"[TRACE SO_FPP] {DEBUG_GOALIE_NAME}: +ATT, +SV :: {m_pair.group(0).strip()}Shot Misses the Net."
                            )
                    elif SO_POST_ANY_RE.search(chunk):
                        found_any_attempt = True
                        ps_att[gdisp] += 1
                        if (
                            DEBUG_PS_TRACE
                            and DEBUG_GOALIE_NAME
                            and gdisp == DEBUG_GOALIE_NAME
                        ):
                            print(
                                f"[TRACE SO_FPP] {DEBUG_GOALIE_NAME}: +ATT :: {m_pair.group(0).strip()}Hits the Post."
                            )
                    elif SO_GOAL_ANY_RE.search(chunk):
                        found_any_attempt = True
                        ps_att[gdisp] += 1
                        if (
                            DEBUG_PS_TRACE
                            and DEBUG_GOALIE_NAME
                            and gdisp == DEBUG_GOALIE_NAME
                        ):
                            print(
                                f"[TRACE SO_FPP] {DEBUG_GOALIE_NAME}: +ATT :: {m_pair.group(0).strip()}Goal."
                            )

            # advance
            idx = after

    # Always return dicts, but if no attempts found in any shootout block, return empties
    return (dict(ps_att) if found_any_attempt else {}), (
        dict(ps_sv) if found_any_attempt else {}
    )


# --- Penalty shots during regulation/OT from Full PBP (unchanged logic, unified tracing) ---
PS_COMBINED_SAVE_RE = re.compile(
    rf"""{_SO_PREFIX}?Penalty\s*Shot[^.]*?\bby\s+(?P<shooter>[^.]+?)
        (?:\s*\((?P<abbr>[A-Z]{2,4})\))?
        [^.]*
        (?:Saved|Stopped)\s+by\s+(?P<goalie>[^.]+)""",
    re.IGNORECASE | re.VERBOSE,
)
PS_COMBINED_GOAL_RE = re.compile(
    rf"""{_SO_PREFIX}?Penalty\s*Shot[^.]*?\bby\s+(?P<shooter>[^.]+?)
        (?:\s*\((?P<abbr>[A-Z]{2,4})\))?
        [^.]*\b(?:Goal|Scores?)\b""",
    re.IGNORECASE | re.VERBOSE,
)
PS_COMBINED_MISS_RE = re.compile(
    rf"""{_SO_PREFIX}?Penalty\s*Shot[^.]*?\bby\s+(?P<shooter>[^.]+?)
        (?:\s*\((?P<abbr>[A-Z]{2,4})\))?
        [^.]*\bMiss(?:es|ed)\s+the\s+Net""",
    re.IGNORECASE | re.VERBOSE,
)
PS_COMBINED_POST_RE = re.compile(
    rf"""{_SO_PREFIX}?Penalty\s*Shot[^.]*?\bby\s+(?P<shooter>[^.]+?)
        (?:\s*\((?P<abbr>[A-Z]{2,4})\))?
        [^.]*\bhits?\s+the\s+post\b""",
    re.IGNORECASE | re.VERBOSE,
)


def parse_penalty_shots_from_fpp(
    fpp_text: str,
    team_map_hint: Dict[str, str],
    goalie_team_map: Dict[str, str],
    goalie_mp_map: Dict[str, int],
    goalie_canon_to_display: Dict[str, str],
) -> Tuple[Dict[str, int], Dict[str, int]]:
    ps_att: Dict[str, int] = defaultdict(int)
    ps_sv: Dict[str, int] = defaultdict(int)

    teams_in_game = {normalize_ws(t) for t in (team_map_hint or {}).values() if t}
    lines = []
    for ln in fpp_text.splitlines():
        s = normalize_ws(ln)
        if s:
            lines.append(s)

    for sent in lines:
        if "penalty shot" not in sent.lower():
            continue

        m = PS_COMBINED_SAVE_RE.search(sent)
        if m:
            goalie_raw = normalize_ws(m.group("goalie"))
            gdisp = goalie_canon_to_display.get(canonical_name(goalie_raw), goalie_raw)
            ps_att[gdisp] += 1
            ps_sv[gdisp] += 1
            _dbg_ps_trace("PS_REG/OT", gdisp, saved=True, line=sent)
            continue

        m = PS_COMBINED_GOAL_RE.search(sent)
        if m:
            shooter = normalize_ws(m.group("shooter"))
            sh_team = normalize_ws((team_map_hint or {}).get(shooter, ""))
            def_team = None
            if teams_in_game and sh_team:
                for t in teams_in_game:
                    if t != sh_team:
                        def_team = t
                        break
            candidates = [
                g
                for g, tm in (goalie_team_map or {}).items()
                if normalize_ws(tm) == normalize_ws(def_team or "")
            ]
            if candidates:
                gdisp = max(candidates, key=lambda g: int(goalie_mp_map.get(g, 0)))
                ps_att[gdisp] += 1
                _dbg_ps_trace("PS_REG/OT", gdisp, saved=False, line=sent)
            continue

        m = PS_COMBINED_MISS_RE.search(sent)
        if m:
            shooter = normalize_ws(m.group("shooter"))
            sh_team = normalize_ws((team_map_hint or {}).get(shooter, ""))
            def_team = None
            if teams_in_game and sh_team:
                for t in teams_in_game:
                    if t != sh_team:
                        def_team = t
                        break
            candidates = [
                g
                for g, tm in (goalie_team_map or {}).items()
                if normalize_ws(tm) == normalize_ws(def_team or "")
            ]
            if candidates:
                gdisp = max(candidates, key=lambda g: int(goalie_mp_map.get(g, 0)))
                ps_att[gdisp] += 1
                ps_sv[gdisp] += 1
                _dbg_ps_trace("PS_REG/OT", gdisp, saved=True, line=sent)
            continue

        m = PS_COMBINED_POST_RE.search(sent)
        if m:
            shooter = normalize_ws(m.group("shooter"))
            sh_team = normalize_ws((team_map_hint or {}).get(shooter, ""))
            def_team = None
            if teams_in_game and sh_team:
                for t in teams_in_game:
                    if t != sh_team:
                        def_team = t
                        break
            candidates = [
                g
                for g, tm in (goalie_team_map or {}).items()
                if normalize_ws(tm) == normalize_ws(def_team or "")
            ]
            if candidates:
                gdisp = max(candidates, key=lambda g: int(goalie_mp_map.get(g, 0)))
                ps_att[gdisp] += 1
                _dbg_ps_trace("PS_REG/OT", gdisp, saved=False, line=sent)
            continue

    return dict(ps_att), dict(ps_sv)


def precompute_pia_by_passer_from_fpp(fpp_text: str) -> Dict[str, int]:
    counts: Dict[str, int] = defaultdict(int)
    chunks = [m.group(0) for m in _SENT_RE.finditer(fpp_text)]
    tail = fpp_text[fpp_text.rfind(".") + 1 :] if "." in fpp_text else fpp_text
    if tail.strip():
        chunks.append(tail.strip())

    cur_time = None
    cur_period = None
    seen_tp: Set[Tuple[str, str, str, str]] = set()
    last_seen_idx: Dict[Tuple[str, str], int] = {}

    for idx, raw in enumerate(chunks):
        sent = normalize_ws(raw)
        mh = FPP_TIMEHDR_RE.search(sent)
        if mh:
            cur_time = mh.group("time")
            cur_period = mh.group("period")
        lower = sent.lower()
        if "intercept" not in lower or "pass" not in lower:
            continue
        passer = intcpt = None
        for pat in PASS_INT_AGAINST_PATTERNS:
            mm = pat.search(sent)
            if mm:
                passer = normalize_ws(mm.group("passer"))
                intcpt = normalize_ws(mm.group("intcpt"))
                break
        if not passer or not intcpt:
            continue
        if cur_time and cur_period:
            k = (cur_period.lower(), cur_time, passer.lower(), intcpt.lower())
            if k in seen_tp:
                continue
            seen_tp.add(k)
        else:
            k2 = (passer.lower(), intcpt.lower())
            last_idx = last_seen_idx.get(k2, -(10**9))
            if idx - last_idx <= NEAR_DUP_WINDOW:
                continue
            last_seen_idx[k2] = idx
        counts[passer] += 1
    return counts


# ---------------- CONNECTION POOLING + RETRIES ----------------
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.8",
    "Connection": "keep-alive",
}
SESSION = requests.Session()
adapter = HTTPAdapter(
    pool_connections=PARALLEL_WORKERS * 4,
    pool_maxsize=PARALLEL_WORKERS * 8,
    max_retries=Retry(
        total=MAX_RETRIES,
        connect=MAX_RETRIES,
        read=MAX_RETRIES,
        backoff_factor=0.2,  # slightly gentler retry cadence
        status_forcelist=[403, 429, 500, 502, 503, 504, 520, 522, 524],
        allowed_methods=frozenset(["HEAD", "GET", "OPTIONS"]),
        raise_on_status=False,
        respect_retry_after_header=True,
    ),
)
# Avoid brotli unless brotli/brotlicffi is installed
HEADERS.update({"Accept-Encoding": "gzip, deflate"})
SESSION.mount("https://", adapter)
SESSION.mount("http://", adapter)


def http_get_with_retries(url: str, *, timeout=REQUEST_TIMEOUT):
    return SESSION.get(url, headers=HEADERS, timeout=timeout)


# ---------------- FETCH PAGES ----------------


def extract_text_from_html(content: bytes) -> str:
    """
    Fast HTML->text extractor using lxml.html when available; falls back to BeautifulSoup.
    Joins text nodes with newlines to preserve line-oriented parsing while trimming whitespace.
    Set env VHLM_PARSER_FORCE_BS4=1 to force BeautifulSoup path (useful for worker stability on Windows).
    """
    if os.environ.get("VHLM_PARSER_FORCE_BS4") != "1" and LH is not None:
        try:
            tree = LH.fromstring(content)
            parts: List[str] = []
            for t in tree.itertext():
                ts = t.strip()
                if ts:
                    parts.append(ts)
            return "\n".join(parts)
        except Exception:
            pass
    parser = "html.parser" if os.environ.get("VHLM_PARSER_FORCE_BS4") == "1" else "lxml"
    soup = BeautifulSoup(content, parser)
    return soup.get_text(separator="\n")


def parse_shootout_ps_from_any_text(
    full_text_lines: str,
    team_map_hint: Dict[str, str],
    goalie_team_map: Dict[str, str],
    goalie_mp_map: Dict[str, int],
    goalie_canon_to_display: Dict[str, str],
    goalie_w_map: Optional[Dict[str, int]] = None,
    goalie_l_map: Optional[Dict[str, int]] = None,
    goalie_otl_map: Optional[Dict[str, int]] = None,
) -> Tuple[Dict[str, int], Dict[str, int]]:
    """
    STRICT shootout parser:
      - Only counts attempts inside 'Shootout/SO' blocks *after* a 'Round #' header is seen.
      - 'Shot Misses the Net' counts as attempt + save for the defending goalie.
      - 'Goal by' counts as attempt only.
      - 'Stopped/Saved by <goalie>' counts as attempt + save on that named goalie.
      - Keeps 'current_shooting_team' across sentences so multi-sentence rounds work.
      - Emits [TRACE SO_ANY] lines when DEBUG_GOALIE_NAME matches the credited goalie.
    """
    ps_att: Dict[str, int] = defaultdict(int)
    ps_sv: Dict[str, int] = defaultdict(int)

    goalie_w_map = goalie_w_map or {}
    goalie_l_map = goalie_l_map or {}
    goalie_otl_map = goalie_otl_map or {}

    # Helper regexes (reuse module-level compiled patterns to avoid re-compilation on every call)
    ROUND_TOKEN = SHOOTOUT_ROUND_RE
    SHOOTER_ANYWHERE = SHOOTER_ANY_LINE
    SAVE_ANY = SO_SAVE_ANY_RE
    MISS_ANY = SO_MISS_ANY_RE
    POST_ANY = SO_POST_ANY_RE
    GOAL_ANY = SO_GOAL_ANY_RE

    # Hard filter: ignore any “time of OT period …” noise just in case it leaks in
    OT_TIME_NOISE = re.compile(
        r"\b\d{1,2}:\d{2}\s+of\s+\d+(?:st|nd|rd|th)\s+overtime\s+period\b",
        re.IGNORECASE,
    )

    # Build team set and abbr→full mapping
    teams_in_game_full = {normalize_ws(t) for t in (team_map_hint or {}).values() if t}
    abbr_to_team_full = _build_abbrev_to_team(full_text_lines, team_map_hint or {})

    def norm_team(tok: Optional[str]) -> Optional[str]:
        if not tok:
            return None
        t = normalize_ws(tok)
        up = t.upper()
        if re.fullmatch(r"[A-Z]{2,4}", up) and up in abbr_to_team_full:
            return normalize_ws(abbr_to_team_full[up])
        for full in teams_in_game_full:
            if t == normalize_ws(full) or t.lower() in full.lower():
                return full
        return t

    def other_team(team: Optional[str]) -> Optional[str]:
        if not team or not teams_in_game_full:
            return None
        t_key = normalize_ws(team)
        for full in teams_in_game_full:
            if normalize_ws(full) != t_key:
                return full
        return None

    def finishing_goalie_for_team(
        team: Optional[str], goalie_by_team_current: Dict[str, str]
    ) -> Optional[str]:
        """Prefer goalie with W/L/OTL on the Goalie Stats line; else nonzero MP; else max MP."""
        if not team:
            return None
        t = normalize_ws(team)
        if t in goalie_by_team_current:
            return goalie_by_team_current[t]
        candidates = [
            g for g, tm in (goalie_team_map or {}).items() if normalize_ws(tm) == t
        ]
        if not candidates:
            return None

        def has_result(g: str) -> int:
            return int(
                goalie_w_map.get(g, 0)
                + goalie_l_map.get(g, 0)
                + goalie_otl_map.get(g, 0)
                > 0
            )

        def mp(g: str) -> int:
            return int(goalie_mp_map.get(g, 0))

        return max(candidates, key=lambda g: (has_result(g), mp(g) > 0, mp(g)))

    # Sentence splitter (robust when many sentences live on one line)
    SENT_SPLIT = re.compile(r"[^.]+(?:\.)")

    spans = __period_spans_for_summary(full_text_lines)

    for start, end, label in spans:
        if not re.search(r"\b(Shoot\s*Out|Shootout|SO)\b", label or "", re.I):
            continue

        chunk = full_text_lines[start:end]
        # Require at least one “Round #” in this shootout block
        if not ROUND_TOKEN.search(chunk):
            continue

        current_shooting_team: Optional[str] = None
        goalie_by_team_current: Dict[str, str] = {}
        in_round = (
            False  # ← gate: only count after we see a “Round #” header in this block
        )

        for m in SENT_SPLIT.finditer(chunk):
            sent = normalize_ws(m.group(0))
            if not sent:
                continue

            # Flip on when we see a round header; do not count earlier sentences
            if ROUND_TOKEN.search(sent):
                in_round = True
                current_shooting_team = None
                continue

            # Ignore any OT-timestamp noise even inside the block
            if OT_TIME_NOISE.search(sent):
                continue

            # Do not process anything until a Round header has appeared
            if not in_round:
                continue

            # If the sentence introduces a shooter (e.g., "Team, Player.")
            ms = SHOOTER_ANYWHERE.search(sent)
            if ms:
                current_shooting_team = norm_team(ms.group("team"))

            # Outcomes (can be same or next sentence)
            m_save = SAVE_ANY.search(sent)
            if m_save:
                raw_g = normalize_ws(m_save.group("goalie"))
                gdisp = goalie_canon_to_display.get(canonical_name(raw_g), raw_g)
                ps_att[gdisp] += 1
                ps_sv[gdisp] += 1
                if DEBUG_PS_TRACE and DEBUG_GOALIE_NAME and gdisp == DEBUG_GOALIE_NAME:
                    print(f"[TRACE SO_ANY] {DEBUG_GOALIE_NAME}: +ATT, +SV :: {sent}")
                # lock defending goalie for future attempts
                def_team = other_team(current_shooting_team)
                if def_team:
                    goalie_by_team_current[normalize_ws(def_team)] = gdisp
                current_shooting_team = None
                continue

            if MISS_ANY.search(sent):
                def_team = other_team(current_shooting_team)
                gdisp = finishing_goalie_for_team(def_team, goalie_by_team_current)
                if gdisp:
                    goalie_by_team_current[normalize_ws(def_team)] = gdisp
                    ps_att[gdisp] += 1
                    ps_sv[gdisp] += 1
                    if (
                        DEBUG_PS_TRACE
                        and DEBUG_GOALIE_NAME
                        and gdisp == DEBUG_GOALIE_NAME
                    ):
                        print(
                            f"[TRACE SO_ANY] {DEBUG_GOALIE_NAME}: +ATT, +SV :: {sent}"
                        )
                current_shooting_team = None
                continue

            if POST_ANY.search(sent):
                def_team = other_team(current_shooting_team)
                gdisp = finishing_goalie_for_team(def_team, goalie_by_team_current)
                if gdisp:
                    goalie_by_team_current[normalize_ws(def_team)] = gdisp
                    ps_att[gdisp] += 1
                    if (
                        DEBUG_PS_TRACE
                        and DEBUG_GOALIE_NAME
                        and gdisp == DEBUG_GOALIE_NAME
                    ):
                        print(f"[TRACE SO_ANY] {DEBUG_GOALIE_NAME}: +ATT :: {sent}")
                current_shooting_team = None
                continue

            m_goal = GOAL_ANY.search(sent)
            if m_goal:
                def_team = other_team(current_shooting_team)
                gdisp = finishing_goalie_for_team(def_team, goalie_by_team_current)
                if gdisp:
                    goalie_by_team_current[normalize_ws(def_team)] = gdisp
                    ps_att[gdisp] += 1
                    if (
                        DEBUG_PS_TRACE
                        and DEBUG_GOALIE_NAME
                        and gdisp == DEBUG_GOALIE_NAME
                    ):
                        print(f"[TRACE SO_ANY] {DEBUG_GOALIE_NAME}: +ATT :: {sent}")
                current_shooting_team = None
                continue

    return dict(ps_att), dict(ps_sv)


def fetch_and_parse_page(
    game_no: int, content_bytes: Optional[bytes] = None
) -> Optional[dict]:
    url = (
        f"https://vhlportal.com/vhlm/{SEASON}/Playoffs/"
        f"VHLM{SEASON}-PLF-{game_no}.html"
    )
    try:
        content = content_bytes
        if content is None:
            resp = http_get_with_retries(url)
            if resp.status_code != 200:
                print(f"[HTTP {resp.status_code}] {url}")
                return None
            content = resp.content

        raw_text = extract_text_from_html(content)
        full_text_lines = raw_text
        full_text_flat = normalize_ws(
            raw_text
        )  # make the space-collapsed version from the same text

        # Trim PBP area if STRICT_PBP_ONLY
        text_for_pbp = full_text_flat
        if STRICT_PBP_ONLY:
            m = re.search(
                r"\bFull\s+Play-?by-?Play\b", full_text_flat, flags=re.IGNORECASE
            )
            if m:
                text_for_pbp = full_text_flat[: m.start()]

        # Collect (time, period, sentence) triplets from PBP
        sents_tp: List[Tuple[str, str, str]] = []
        if GAME_MARKER_RE.search(text_for_pbp):
            for t, p, blk in extract_pbp_blocks(text_for_pbp):
                for sent in iter_sentences(blk):
                    sents_tp.append((t, p, sent))

        success_hits_by_hitter = (
            precompute_successful_hits_by_hitter(sents_tp) if sents_tp else {}
        )

        # Tables
        (
            pm_map,
            hits_map,
            pim_map,
            shots_map,
            sb_map,
            mp_map,
            team_map,
            fow_map,
            fot_map,
        ) = parse_player_stats_table(full_text_lines)
        goalie_maps = parse_goalies_stats_table(full_text_lines)

        # Table values (if any)
        _debug_ps_emit(
            game_no,
            url,
            "TABLE",
            goalie_maps.get("ps_att_map", {}),
            goalie_maps.get("ps_sv_map", {}),
        )

        # Map goalie team abbreviations to full team names using skater team_map hints on this page
        try:
            abbr_to_team_full = _build_abbrev_to_team(full_text_lines, team_map)
            gtm = goalie_maps.get("team_map", {})
            if gtm and abbr_to_team_full:
                remapped = {}
                for gdisp, t in gtm.items():
                    ab = (t or "").strip().upper()
                    remapped[gdisp] = abbr_to_team_full.get(ab, t)
                goalie_maps["team_map"] = remapped
        except Exception:
            pass

        # Canonical -> display for goalies on this page
        goalies_present_display = list(goalie_maps.get("team_map", {}).keys())
        goalie_canon_to_display = {
            canonical_name(n): n for n in goalies_present_display
        }

        # Goalie PIM from 'Penalties :' blocks (even if served by others)
        goalie_pim_from_pen = parse_goalie_pim_from_penalties(
            full_text_lines, goalie_canon_to_display
        )

        # Scoring summary-derived maps
        goals_map, assists_map = parse_goals_assists_from_summary(full_text_flat)
        ppg_map, ppa_map = parse_pp_from_scoring_summary(full_text_flat)
        pkg_map, pka_map = parse_sh_from_scoring_summary(full_text_flat)

        # GWG/GTG + debug
        try:
            gwg_map, gtg_map, gtg_debug = compute_gwg_gtg_and_debug(
                full_text_flat, full_text_lines=full_text_lines, team_map_hint=team_map
            )
        except Exception as e:
            print(f"[GWG/GTG ERROR] game {game_no}: {e}")
            gwg_map, gtg_map, gtg_debug = {}, {}, {}

        # Derived counts from PBP
        giveaways_by_victim = (
            precompute_giveaways_by_victim(sents_tp) if sents_tp else {}
        )
        hits_taken_by_victim = (
            precompute_hits_taken_by_victim(sents_tp) if sents_tp else {}
        )

        # Tail text (Full PBP)
        fpp_text = extract_full_pbp_tail(full_text_lines)

        # --- Goalie SAR from PBP ("Stopped by <goalie> with a rebound") ---
        sar_counts: Dict[str, int] = defaultdict(int)
        for m in REBOUND_SAVE_RE.finditer(fpp_text):
            gnm = normalize_ws(m.group("goalie"))
            if gnm:
                sar_counts[gnm] += 1

        if sar_counts:
            g_sar = goalie_maps.get("sar_map", {})
            for raw_name, cnt in sar_counts.items():
                disp = goalie_canon_to_display.get(canonical_name(raw_name), raw_name)
                g_sar[disp] = g_sar.get(disp, 0) + cnt
            goalie_maps["sar_map"] = g_sar

        # === Shootout + Penalty Shot parsing and merge =======================
        so_ps_att, so_ps_sv = {}, {}
        ps_ps_att, ps_ps_sv = {}, {}

        # (A) STRICT FPP-tail shootout parser (Rounds only) — run FIRST so attempt-level traces print
        try:
            so_ps_att, so_ps_sv = parse_shootout_ps_from_fpp(
                fpp_text,
                team_map or {},
                goalie_maps.get("team_map", {}),
                goalie_maps.get("mp_map", {}),
                goalie_canon_to_display,
                goalie_w_map=goalie_maps.get("w_map", {}),
                goalie_l_map=goalie_maps.get("l_map", {}),
                goalie_otl_map=goalie_maps.get("otl_map", {}),
            )
        except Exception as e:
            print(f"[SHOOTOUT FPP WARN] game {game_no}: {e}")
        _debug_ps_emit(game_no, url, "SO_FPP", so_ps_att, so_ps_sv)

        # (B) Fallback: ANY-TEXT parser (only if FPP found nothing)
        if not so_ps_att and not so_ps_sv:
            try:
                so_ps_att, so_ps_sv = parse_shootout_ps_from_any_text(
                    full_text_lines,
                    team_map or {},
                    goalie_maps.get("team_map", {}),
                    goalie_maps.get("mp_map", {}),
                    goalie_canon_to_display,
                    goalie_w_map=goalie_maps.get("w_map", {}),
                    goalie_l_map=goalie_maps.get("l_map", {}),
                    goalie_otl_map=goalie_maps.get("otl_map", {}),
                )
            except Exception as e:
                print(f"[SHOOTOUT ANY WARN] game {game_no}: {e}")
            _debug_ps_emit(game_no, url, "SO_ANY", so_ps_att, so_ps_sv)

        # (C) Penalty shots during reg/OT
        try:
            ps_ps_att, ps_ps_sv = parse_penalty_shots_from_fpp(
                fpp_text,
                team_map or {},
                goalie_maps.get("team_map", {}),
                goalie_maps.get("mp_map", {}),
                goalie_canon_to_display,
            )
        except Exception as e:
            print(f"[PENSHOT PARSE WARN] game {game_no}: {e}")
        _debug_ps_emit(game_no, url, "PS_REG/OT", ps_ps_att, ps_ps_sv)

        # (D) Combine & merge (max, no downgrades)
        comb_att: Dict[str, int] = defaultdict(int)
        comb_sv: Dict[str, int] = defaultdict(int)
        for d in (so_ps_att, ps_ps_att):
            for k, v in (d or {}).items():
                comb_att[k] += int(v)
        for d in (so_ps_sv, ps_ps_sv):
            for k, v in (d or {}).items():
                comb_sv[k] += int(v)

        _debug_ps_emit(game_no, url, "COMBINED", comb_att, comb_sv)

        if comb_att or comb_sv:
            g_ps_att = dict(goalie_maps.get("ps_att_map", {}) or {})
            g_ps_sv = dict(goalie_maps.get("ps_sv_map", {}) or {})
            # debug pre-merge view
            _debug_ps_emit(game_no, url, "PRE-MERGE", g_ps_att, g_ps_sv)

            for gdisp, att in comb_att.items():
                g_ps_att[gdisp] = max(int(g_ps_att.get(gdisp, 0)), int(att))
            for gdisp, sv in comb_sv.items():
                g_ps_sv[gdisp] = max(int(g_ps_sv.get(gdisp, 0)), int(sv))

            # debug post-merge view
            _debug_ps_emit(game_no, url, "POST-MERGE", g_ps_att, g_ps_sv)

            goalie_maps["ps_att_map"] = g_ps_att
            goalie_maps["ps_sv_map"] = g_ps_sv
        # === END NEW =========================================================

        # --- Retrievals, interceptions, PIA, 3 Stars ---
        retr_counts: Dict[str, int] = defaultdict(int)
        for mr in RETRIEVE_ANY_RE.finditer(fpp_text):
            who = normalize_ws(mr.group("who"))
            retr_counts[who] += 1

        # IMPORTANT: count interceptions on tail only to avoid doubles
        intc_counts: Dict[str, int] = defaultdict(int)
        for mi in INTERCEPT_BY_ANY_RE.finditer(fpp_text):
            who = normalize_ws(mi.group(1))
            intc_counts[who] += 1

        pia_counts = precompute_pia_by_passer_from_fpp(fpp_text)
        s1_map, s2_map, s3_map = parse_three_stars(full_text_lines)

        # --- Team results (W/L/OTL) for skaters ---
        team_results_by_key: Dict[str, Dict[str, object]] = {}

        def _team_key(name: Optional[str]) -> str:
            return normalize_ws(name or "").lower()

        def _ensure_team_entry(team_name: Optional[str]) -> Optional[Dict[str, object]]:
            if not team_name:
                return None
            key = _team_key(team_name)
            if not key:
                return None
            entry = team_results_by_key.setdefault(
                key,
                {
                    "team_display": normalize_ws(team_name),
                    "W": 0,
                    "L": 0,
                    "OTL": 0,
                },
            )
            disp = normalize_ws(team_name)
            cur_disp = normalize_ws(str(entry.get("team_display", "")))
            if disp and len(disp) > len(cur_disp):
                entry["team_display"] = team_name
            return entry

        goalie_team_map = dict(goalie_maps.get("team_map", {}) or {})
        for result_key, map_name in (
            ("W", "w_map"),
            ("L", "l_map"),
            ("OTL", "otl_map"),
        ):
            res_map = goalie_maps.get(map_name, {}) or {}
            for gname, val in res_map.items():
                try:
                    count = int(val)
                except Exception:
                    count = 0
                if count <= 0:
                    continue
                team_name = goalie_team_map.get(gname)
                if not team_name:
                    continue
                entry = _ensure_team_entry(team_name)
                if entry is not None:
                    entry[result_key] = max(int(entry.get(result_key, 0)), 1)

        if not team_results_by_key:
            team_goals_raw: Dict[str, int] = defaultdict(int)
            team_display_raw: Dict[str, str] = {}
            for skater_name, goals_scored in goals_map.items():
                team_name = team_map.get(skater_name)
                if not team_name:
                    continue
                key = _team_key(team_name)
                try:
                    goals_int = int(goals_scored)
                except Exception:
                    goals_int = 0
                team_goals_raw[key] += goals_int
                existing = team_display_raw.get(key, "")
                if not existing or len(normalize_ws(team_name)) > len(
                    normalize_ws(existing)
                ):
                    team_display_raw[key] = team_name

            if team_goals_raw:
                winner_key: Optional[str] = None
                debug_dict = gtg_debug if isinstance(gtg_debug, dict) else {}
                so_team = debug_dict.get("so_winner_team") if debug_dict else None
                if so_team:
                    candidate = _team_key(so_team)
                    if candidate in team_goals_raw:
                        winner_key = candidate
                if winner_key is None:
                    winner_key = max(team_goals_raw.items(), key=lambda kv: kv[1])[0]

                loser_keys = [k for k in team_goals_raw.keys() if k != winner_key]
                loser_key = loser_keys[0] if loser_keys else None

                if winner_key is not None:
                    entry = _ensure_team_entry(
                        team_display_raw.get(winner_key, winner_key)
                    )
                    if entry is not None:
                        entry["W"] = max(int(entry.get("W", 0)), 1)

                if loser_key is not None:
                    entry = _ensure_team_entry(
                        team_display_raw.get(loser_key, loser_key)
                    )
                    if entry is not None:
                        went_ot = (
                            bool(debug_dict.get("went_to_ot")) if debug_dict else False
                        )
                        went_so = (
                            bool(debug_dict.get("went_to_so")) if debug_dict else False
                        )
                        if went_ot or went_so:
                            entry["OTL"] = max(int(entry.get("OTL", 0)), 1)
                        else:
                            entry["L"] = max(int(entry.get("L", 0)), 1)

        team_results_payload = {
            key: {
                "team": normalize_ws(str(data.get("team_display", ""))),
                "W": 1 if int(data.get("W", 0)) > 0 else 0,
                "L": 1 if int(data.get("L", 0)) > 0 else 0,
                "OTL": 1 if int(data.get("OTL", 0)) > 0 else 0,
            }
            for key, data in team_results_by_key.items()
        }

        return {
            "url": url,
            "tables": (
                pm_map,
                hits_map,
                pim_map,
                shots_map,
                sb_map,
                mp_map,
                team_map,
                fow_map,
                fot_map,
            ),
            "goals_assists": (goals_map, assists_map),
            "ppg_by_player": ppg_map,
            "ppa_by_player": ppa_map,
            "pkg_by_player": pkg_map,
            "pka_by_player": pka_map,
            "gwg_by_player": gwg_map,
            "gtg_by_player": gtg_map,
            "giveaways_by_victim": giveaways_by_victim,
            "hits_taken_by_victim": hits_taken_by_victim,
            "retrievals_by_player": retr_counts,
            "intercepts_by_player": intc_counts,
            "pia_by_passer": pia_counts,
            "s1_by_player": s1_map,
            "s2_by_player": s2_map,
            "s3_by_player": s3_map,
            "debug_gtg": gtg_debug,
            "goalies": goalie_maps,
            "goalie_pim_from_pen": goalie_pim_from_pen,  # NEW
            "success_hits_by_hitter": success_hits_by_hitter,  # NEW
            "team_results": team_results_payload,
        }

    except Exception as e:
        print(f"[PARSE ERROR] game {game_no}: {e}")
        return None


# Async I/O fetch + CPU-bound parse (hybrid). Falls back to threads if aiohttp missing.
async def _async_fetch_one(
    session: "aiohttp.ClientSession", game_no: int
) -> tuple[int, Optional[bytes]]:
    url = (
        f"https://vhlportal.com/vhlm/{SEASON}/Playoffs/"
        f"VHLM{SEASON}-PLF-{game_no}.html"
    )
    status_forcelist = {403, 429, 500, 502, 503, 504, 520, 522, 524}
    backoff = 0.2
    for attempt in range(MAX_RETRIES):
        try:
            async with session.get(url, headers=HEADERS) as resp:
                if resp.status == 200:
                    return game_no, await resp.read()
                if resp.status in status_forcelist:
                    await asyncio.sleep(backoff * (2**attempt))
                    continue
                print(f"[HTTP {resp.status}] {url}")
                return game_no, None
        except Exception:
            await asyncio.sleep(backoff * (2**attempt))
    print(f"[HTTP FAIL] {url}")
    return game_no, None


async def _async_fetch_all(start_no: int, end_no: int) -> list[tuple[int, bytes]]:
    connector = aiohttp.TCPConnector(limit=ASYNC_MAX_CONNECTIONS)
    timeout = aiohttp.ClientTimeout(total=REQUEST_TIMEOUT + 5)
    async with aiohttp.ClientSession(connector=connector, timeout=timeout) as session:
        tasks = [_async_fetch_one(session, n) for n in range(start_no, end_no + 1)]
        results = await asyncio.gather(*tasks, return_exceptions=True)
    out: list[tuple[int, bytes]] = []
    for r in results:
        if isinstance(r, Exception):
            continue
        if not r:
            continue
        n, content = r
        if content:
            out.append((n, content))
    return out


def fetch_range_async_hybrid(start_no: int, end_no: int) -> Dict[int, dict]:
    if aiohttp is None:
        # Fallback to existing threaded approach
        return fetch_range(start_no, end_no)

    print(
        f"Async fetching pages {start_no}..{end_no} with {ASYNC_MAX_CONNECTIONS} connections; parsing on {CPU_WORKERS} threads."
    )
    try:
        loop = asyncio.new_event_loop()
        try:
            asyncio.set_event_loop(loop)
            fetched = loop.run_until_complete(_async_fetch_all(start_no, end_no))
        finally:
            loop.close()
    except Exception as e:
        print(f"[ASYNC ERROR] Falling back to threads: {e}")
        return fetch_range(start_no, end_no)

    total = len(fetched)
    if total == 0:
        print("No pages fetched via async; falling back to threads.")
        return fetch_range(start_no, end_no)

    page_data: Dict[int, dict] = {}
    with ThreadPoolExecutor(max_workers=CPU_WORKERS) as ex:
        futures = {
            ex.submit(fetch_and_parse_page, n, content): n for (n, content) in fetched
        }
        done = 0
        for fut in as_completed(futures):
            n = futures[fut]
            try:
                data = fut.result()
                if data:
                    page_data[n] = data
            except Exception as te:
                print(f"[THREAD PARSE ERROR] game {n}: {te}")
            done += 1
            if done % 25 == 0 or done == total:
                print(f"  parsed (threads) {done}/{total} ...")

    print(f"Parsed (threads) {len(page_data)} / {total} pages.")
    return page_data


def fetch_range(start_no: int, end_no: int) -> Dict[int, dict]:
    page_data: Dict[int, dict] = {}
    total = end_no - start_no + 1
    print(f"Fetching pages {start_no}..{end_no} (total {total}) ...")

    with ThreadPoolExecutor(max_workers=PARALLEL_WORKERS) as ex:
        futures = {}
        for game_no in range(start_no, end_no + 1):
            futures[ex.submit(fetch_and_parse_page, game_no)] = game_no

        done = 0
        for fut in as_completed(futures):
            n = futures[fut]
            try:
                data = fut.result()
                if data:
                    page_data[n] = data
            except Exception as e:
                print(f"[FUTURE ERROR] game {n}: {e}")
            done += 1
            if done % 25 == 0 or done == total:
                print(f"  fetched {done}/{total} ...")

    print(f"Fetched {len(page_data)} / {total} pages.")
    return page_data


# ---------------- FETCH ALL PAGES ----------------
# Prefer hybrid async I/O + process pool, fallback to threaded fetch+parse


def main():
    # Ensure safe start method for multiprocessing on Windows
    try:
        mp.set_start_method("spawn", force=True)
    except RuntimeError:
        pass
    t0_fetch = time.perf_counter()
    page_data_by_no = fetch_range_async_hybrid(START_GAME_NO, END_GAME_NO)
    print(f"[TIME] fetch+parse: {time.perf_counter() - t0_fetch:.2f}s")
    global page_player_tables, page_goals_assists, page_ppg_by_player, page_ppa_by_player, page_pkg_by_player, page_pka_by_player, page_gwg_by_player, page_gtg_by_player, page_giveaways_by_victim, page_hits_taken_by_victim, page_retrievals_by_player, page_intercepts_by_player, page_pia_by_passer, page_s1_by_player, page_s2_by_player, page_s3_by_player, page_success_hits_by_hitter, page_team_results, page_goalie_team_map, page_goalie_mp_map, page_goalie_ga_map, page_goalie_sa_map, page_goalie_pim_map, page_goalie_w_map, page_goalie_l_map, page_goalie_otl_map, page_goalie_so_map, page_goalie_a_map, page_goalie_eg_map, page_goalie_ps_sv_map, page_goalie_ps_att_map, page_goalie_sar_map, page_goalie_pim_from_pen_map, gtg_debug_records

    page_player_tables = {}
    page_goals_assists = {}
    page_ppg_by_player = {}
    page_ppa_by_player = {}
    page_pkg_by_player = {}
    page_pka_by_player = {}
    page_gwg_by_player = {}
    page_gtg_by_player = {}
    page_giveaways_by_victim = {}
    page_hits_taken_by_victim = {}
    page_retrievals_by_player = {}
    page_intercepts_by_player = {}
    page_pia_by_passer = {}
    page_s1_by_player = {}
    page_s2_by_player = {}
    page_s3_by_player = {}
    page_success_hits_by_hitter = {}
    page_team_results = {}

    # --- Goalie page maps ---
    page_goalie_team_map = {}
    page_goalie_mp_map = {}
    page_goalie_ga_map = {}
    page_goalie_sa_map = {}
    page_goalie_pim_map = {}
    page_goalie_w_map = {}
    page_goalie_l_map = {}
    page_goalie_otl_map = {}
    page_goalie_so_map = {}
    page_goalie_a_map = {}
    page_goalie_eg_map = {}
    page_goalie_ps_sv_map = {}
    page_goalie_ps_att_map = {}
    page_goalie_sar_map = {}
    page_goalie_pim_from_pen_map = {}  # NEW

    gtg_debug_records = []  # or simply: gtg_debug_records = []

    for no, data in sorted(page_data_by_no.items()):
        url = data["url"]
        page_player_tables[url] = data["tables"]
        page_goals_assists[url] = data["goals_assists"]
        page_ppg_by_player[url] = data["ppg_by_player"]
        page_ppa_by_player[url] = data["ppa_by_player"]
        page_pkg_by_player[url] = data["pkg_by_player"]
        page_pka_by_player[url] = data["pka_by_player"]
        page_gwg_by_player[url] = data["gwg_by_player"]
        page_gtg_by_player[url] = data["gtg_by_player"]
        page_giveaways_by_victim[url] = data["giveaways_by_victim"]
        page_hits_taken_by_victim[url] = data["hits_taken_by_victim"]
        page_retrievals_by_player[url] = data["retrievals_by_player"]
        page_intercepts_by_player[url] = data["intercepts_by_player"]
        page_pia_by_passer[url] = data["pia_by_passer"]
        page_s1_by_player[url] = data["s1_by_player"]
        page_s2_by_player[url] = data["s2_by_player"]
        page_s3_by_player[url] = data["s3_by_player"]
        page_success_hits_by_hitter[url] = data.get("success_hits_by_hitter", {})
        page_team_results[url] = data.get("team_results", {})

        g = data.get("goalies", {})
        if g:
            page_goalie_team_map[url] = g.get("team_map", {})
            page_goalie_mp_map[url] = g.get("mp_map", {})
            page_goalie_ga_map[url] = g.get("ga_map", {})
            page_goalie_sa_map[url] = g.get("sa_map", {})
            page_goalie_pim_map[url] = g.get("pim_map", {})
            page_goalie_w_map[url] = g.get("w_map", {})
            page_goalie_l_map[url] = g.get("l_map", {})
            page_goalie_otl_map[url] = g.get("otl_map", {})
            page_goalie_so_map[url] = g.get("so_map", {})
            page_goalie_a_map[url] = g.get("a_map", {})
            page_goalie_eg_map[url] = g.get("eg_map", {})
            page_goalie_ps_sv_map[url] = g.get("ps_sv_map", {})
            page_goalie_ps_att_map[url] = g.get("ps_att_map", {})
            page_goalie_sar_map[url] = g.get("sar_map", {})
            page_goalie_pim_from_pen_map[url] = data.get("goalie_pim_from_pen", {})

        dbg = data.get("debug_gtg")
        if dbg:
            item = dict(dbg)
            item["url"] = url
            gtg_debug_records.append(item)


if __name__ == "__main__":
    main()

# ---------------------------------------------------------------------
# DYNAMIC NAME LIST FROM GAMES (+ POS enrichment via Players Info)
#   *** Only include players with POS enrichment (filters out bots).
# ---------------------------------------------------------------------


# Build goalie default team mapping (moved from main() to global scope)
def _looks_abbr(s: str) -> bool:
    return bool(re.fullmatch(r"[A-Z]{2,4}", (s or "").strip()))


goalie_default_team: Dict[str, str] = {}
_counts = defaultdict(Counter)
for umap in page_goalie_team_map.values():
    for gdisp, t in umap.items():
        if t and not _looks_abbr(t):
            _counts[gdisp][normalize_ws(t)] += 1
for gdisp, ctr in _counts.items():
    if ctr:
        goalie_default_team[gdisp] = ctr.most_common(1)[0][0]


def find_players_info_url(soup: BeautifulSoup, page_url: str) -> Optional[str]:
    for a in soup.find_all("a", href=True):
        txt = (a.get_text() or "").strip().lower()
        if re.search(r"players?\s*'?s?\s*info", txt):
            return requests.compat.urljoin(page_url, a["href"])
    for a in soup.find_all("a", href=True):
        href = a["href"].lower()
        if any(
            k in href
            for k in [
                "playersinfo",
                "playerinfo",
                "players.php",
                "playerinfo.php",
                "player.php",
            ]
        ):
            return requests.compat.urljoin(page_url, a["href"])
    for a in soup.find_all("a", href=True):
        href = a["href"].lower()
        if "player" in href and "info" in href:
            return requests.compat.urljoin(page_url, a["href"])
    return None


def _find_players_info_table_and_cols(soup: BeautifulSoup):
    def _norm(s: str) -> str:
        return normalize_ws(s).lower()

    def _header_map(headers: List[str]) -> Dict[str, int]:
        m: Dict[str, int] = {}
        for i, h in enumerate(headers):
            n = _norm(h)
            if n in ("player name", "name", "player"):
                m["player name"] = i
            elif n in ("pos", "position"):
                m["pos"] = i
            elif n == "link":
                m["link"] = i
        return m

    for tbl in soup.find_all("table"):
        header_cells: List[str] = []
        thead = tbl.find("thead")
        if thead and thead.find("tr"):
            header_cells = [
                normalize_ws(c.get_text(" "))
                for c in thead.find("tr").find_all(["th", "td"])
            ]
        if not header_cells:
            header_tr = None
            for tr in tbl.find_all("tr"):
                if tr.find("th"):
                    header_tr = tr
                    break
            if header_tr:
                header_cells = [
                    normalize_ws(c.get_text(" "))
                    for c in header_tr.find_all(["th", "td"])
                ]
        if not header_cells:
            continue
        cmap = _header_map(header_cells)
        if "player name" in cmap:
            return tbl, cmap
    return None, {}


def extract_names_and_positions_from_players_info(html: bytes) -> List[Tuple[str, str]]:
    soup = BeautifulSoup(html, "lxml")  # <<< switch parser
    tbl, cmap = _find_players_info_table_and_cols(soup)
    if not tbl:
        return []
    header_tr = None
    thead = tbl.find("thead")
    if thead and thead.find("tr"):
        header_tr = thead.find("tr")
    if not header_tr:
        for tr in tbl.find_all("tr"):
            if tr.find("th"):
                header_tr = tr
                break
    out: List[Tuple[str, str]] = []
    for tr in tbl.find_all("tr"):
        if header_tr and tr is header_tr:
            continue
        tds = tr.find_all("td")
        if not tds or len(tds) < 2:
            continue
        name_idx = cmap.get("player name", 0)
        if name_idx >= len(tds):
            continue
        td_name = tds[name_idx]
        raw_name = normalize_ws(td_name.get_text(" "))
        if not raw_name or "team average" in raw_name.lower():
            continue
        pos_idx = cmap.get("pos", 1)
        if pos_idx >= len(tds):
            pos_idx = 1
        pos_val = normalize_ws(tds[pos_idx].get_text(" "))
        link_ok = False
        link_idx = cmap.get("link", None)
        if link_idx is not None and link_idx < len(tds):
            if tds[link_idx].find("a", href=True):
                link_ok = True
        else:
            for j, td in enumerate(tds):
                if j == name_idx:
                    continue
                if td.find("a", href=True):
                    link_ok = True
                    break
        if not link_ok:
            continue
        name_clean = re.sub(r"\([^)]*\)", "", raw_name)
        name_clean = normalize_ws(name_clean)
        out.append((name_clean, pos_val))
    return out


def is_goalie_position(pos: str) -> bool:
    s = (pos or "").strip().lower()
    return bool(re.search(r"\b(g|goalie|goaltender)\b", s))


def build_names_from_games_and_enrich_pos() -> (
    Tuple[List[str], Dict[str, str], List[str]]
):
    """
    Collect names from game pages and enrich POS via Players Info pages (concurrently).
    Returns (skater_names, name_to_position, goalie_names).
    """

    t0_names = time.perf_counter()

    # 1) Build display-name set from parsed game pages first
    cname_to_display: Dict[str, str] = {}
    for url, tables in page_player_tables.items():
        team_map = tables[6]
        for display_name in team_map.keys():
            c = canonical_name(display_name)
            if c and c not in cname_to_display:
                cname_to_display[c] = display_name

    for url, umap in page_goalie_team_map.items():
        for g_display in umap.keys():
            c = canonical_name(g_display)
            if c and c not in cname_to_display:
                cname_to_display[c] = g_display

    # 2) Concurrently fetch Players Info pages to get POS
    import threading

    cname_to_pos: Dict[str, str] = {}
    lock = threading.Lock()

    def fetch_one(team_id: int):
        team_url = f"https://vhlportal.com/vhlm/{SEASON}/ProTeam.php?Team={team_id}"
        try:
            r = http_get_with_retries(team_url)
            if r.status_code != 200 or not r.content:
                return
            soup = BeautifulSoup(r.content, "lxml")
            info_url = find_players_info_url(soup, team_url)
            if not info_url:
                return
            r_info = http_get_with_retries(info_url)
            if r_info.status_code != 200 or not r_info.content:
                return
            rows = extract_names_and_positions_from_players_info(r_info.content)
            for display_name, pos in rows:
                c = canonical_name(display_name)
                if c and c in cname_to_display:
                    with lock:
                        if c not in cname_to_pos:
                            cname_to_pos[c] = pos
        except Exception:
            pass

    with ThreadPoolExecutor(max_workers=8) as ex:
        list(ex.map(fetch_one, range(1, TEAM_PAGE_MAX_ID + 1)))

    # 3) Build final lists
    skater_names: List[str] = []
    goalie_names: List[str] = []
    name_to_position: Dict[str, str] = {}

    for c, display in cname_to_display.items():
        pos = cname_to_pos.get(c)
        if not pos:
            continue  # filters bots & unlinked rows
        name_to_position[display] = pos
        if is_goalie_position(pos):
            goalie_names.append(display)
        else:
            skater_names.append(display)

    print(
        f"Dynamic NameList (from games): {len(skater_names)} skaters + {len(goalie_names)} goalies = {len(skater_names)+len(goalie_names)} players."
    )
    print(
        f"[TIME] build_names_from_games_and_enrich_pos: {time.perf_counter() - t0_names:.2f}s"
    )
    return skater_names, name_to_position, goalie_names


# ---------------- AGGREGATE PER SKATER ----------------
overall_rows = []
team_split_rows = []

skater_names, name_to_position, goalie_names_from_nameslist = (
    build_names_from_games_and_enrich_pos()
)
player_names = skater_names


def get_search_phrases(player_name: str):
    name_lc = player_name.lower()
    return {"Successful hits": f"is hit by {name_lc} and loses puck"}


for idx, player_name in enumerate(skater_names, start=1):
    print(f"Processing player: {player_name} ({idx}/{len(player_names)})")
    position_str = name_to_position.get(player_name, "")
    phrases = get_search_phrases(player_name)

    games = goals_total = assists_total = plus_minus_total = 0
    wins_total = losses_total = otl_total = 0
    pen_minutes = hits_all = shots_total = shots_blocked = minutes_played_total_sec = 0
    hits_success = schta_total = 0
    hits_taken_all = 0
    retrievals = interceptions = passes_intercepted_against = 0
    ppg_total = 0
    ppa_total = 0
    pkg_total = 0
    pka_total = 0
    gwg_total = 0
    gtg_total = 0
    s1_total = 0
    s2_total = 0
    s3_total = 0
    fo_wins_total = 0
    fo_taken_total = 0
    last_seen_team = ""

    team_acc = defaultdict(
        lambda: {
            "GP": 0,
            "W": 0,
            "L": 0,
            "OTL": 0,
            "G": 0,
            "A": 0,
            "P": 0,
            "PM": 0,
            "PIM": 0,
            "HIT": 0,
            "HTT": 0,
            "SHT": 0,
            "SB": 0,
            "MP": 0,
            "SCHT": 0,
            "SCHTA": 0,
            "PRET": 0,
            "PI": 0,
            "PIA": 0,
            "PPG": 0,
            "PPA": 0,
            "PKG": 0,
            "PKA": 0,
            "GWG": 0,
            "GTG": 0,
            "S1": 0,
            "S2": 0,
            "S3": 0,
            "FOW": 0,
            "FOT": 0,
        }
    )

    for url, (
        pm_map,
        hits_map,
        pim_map,
        shots_map,
        sb_map,
        mp_map,
        team_map,
        fow_map,
        fot_map,
    ) in page_player_tables.items():
        team_val = team_map.get(player_name)
        wins_inc = losses_inc = otl_inc = 0
        if player_name in mp_map:
            games += 1
            last_seen_team = team_val or last_seen_team

            team_key = normalize_ws(team_val or "").lower()
            if team_key:
                team_result_entry = (page_team_results.get(url, {}) or {}).get(
                    team_key, {}
                )
                wins_inc = int(team_result_entry.get("W", 0) or 0)
                losses_inc = int(team_result_entry.get("L", 0) or 0)
                otl_inc = int(team_result_entry.get("OTL", 0) or 0)
                wins_total += wins_inc
                losses_total += losses_inc
                otl_total += otl_inc

        pm_val = int(pm_map.get(player_name, 0))
        h_val = int(hits_map.get(player_name, 0))
        pim_v = int(pim_map.get(player_name, 0))
        s_val = int(shots_map.get(player_name, 0))
        sb_v = int(sb_map.get(player_name, 0))
        mp_sec = int(mp_map.get(player_name, 0))
        fo_wins_total += int(fow_map.get(player_name, 0))
        fo_taken_total += int(fot_map.get(player_name, 0))

        plus_minus_total += pm_val
        hits_all += h_val
        pen_minutes += pim_v
        shots_total += s_val
        shots_blocked += sb_v
        minutes_played_total_sec += mp_sec

        s1_total += fast_get(page_s1_by_player.get(url, {}), player_name)
        s2_total += fast_get(page_s2_by_player.get(url, {}), player_name)
        s3_total += fast_get(page_s3_by_player.get(url, {}), player_name)

        pg_goals, pg_assists = page_goals_assists.get(url, ({}, {}))
        goals_total += pg_goals.get(player_name, 0)
        assists_total += pg_assists.get(player_name, 0)

        schta_this_page = fast_get(page_giveaways_by_victim.get(url, {}), player_name)
        htt_this_page = fast_get(page_hits_taken_by_victim.get(url, {}), player_name)
        schta_total += schta_this_page
        hits_taken_all += htt_this_page

        retrievals += fast_get(page_retrievals_by_player.get(url, {}), player_name)
        interceptions += fast_get(page_intercepts_by_player.get(url, {}), player_name)
        passes_intercepted_against += fast_get(
            page_pia_by_passer.get(url, {}), player_name
        )

        ppg_total += fast_get(page_ppg_by_player.get(url, {}), player_name)
        ppa_total += fast_get(page_ppa_by_player.get(url, {}), player_name)
        pkg_total += fast_get(page_pkg_by_player.get(url, {}), player_name)
        pka_total += fast_get(page_pka_by_player.get(url, {}), player_name)

        gwg_total += fast_get(page_gwg_by_player.get(url, {}), player_name)
        gtg_total += fast_get(page_gtg_by_player.get(url, {}), player_name)

        if team_val:
            acc = team_acc[team_val]
            acc["GP"] += 1
            acc["W"] += wins_inc
            acc["L"] += losses_inc
            acc["OTL"] += otl_inc
            acc["PM"] += pm_val
            acc["HIT"] += h_val
            acc["HTT"] += htt_this_page
            acc["PIM"] += pim_v
            acc["SHT"] += s_val
            acc["SB"] += sb_v
            acc["MP"] += mp_sec
            acc["FOW"] += int(fow_map.get(player_name, 0))
            acc["FOT"] += int(fot_map.get(player_name, 0))
            acc["S1"] += fast_get(page_s1_by_player.get(url, {}), player_name)
            acc["S2"] += fast_get(page_s2_by_player.get(url, {}), player_name)
            acc["S3"] += fast_get(page_s3_by_player.get(url, {}), player_name)

            g_add = pg_goals.get(player_name, 0)
            a_add = pg_assists.get(player_name, 0)
            acc["G"] += g_add
            acc["A"] += a_add

            acc["SCHT"] += int(
                page_success_hits_by_hitter.get(url, {}).get(player_name, 0)
            )

            acc["SCHTA"] += schta_this_page
            acc["PRET"] += fast_get(page_retrievals_by_player.get(url, {}), player_name)
            acc["PI"] += fast_get(page_intercepts_by_player.get(url, {}), player_name)
            acc["PIA"] += fast_get(page_pia_by_passer.get(url, {}), player_name)

            acc["PPG"] += fast_get(page_ppg_by_player.get(url, {}), player_name)
            acc["PPA"] += fast_get(page_ppa_by_player.get(url, {}), player_name)
            acc["PKG"] += fast_get(page_pkg_by_player.get(url, {}), player_name)
            acc["PKA"] += fast_get(page_pka_by_player.get(url, {}), player_name)

            acc["GWG"] += fast_get(page_gwg_by_player.get(url, {}), player_name)
            acc["GTG"] += fast_get(page_gtg_by_player.get(url, {}), player_name)

    # total per player
    for url in page_success_hits_by_hitter:
        hits_success += int(page_success_hits_by_hitter[url].get(player_name, 0))

    team_final = last_seen_team or ""
    points_total = goals_total + assists_total
    shot_pct = (goals_total / shots_total * 100.0) if shots_total > 0 else 0.0
    shot_pct_str = f"{round(shot_pct, 2)}%"

    fo_pct = (fo_wins_total / fo_taken_total * 100.0) if fo_taken_total > 0 else 0.0
    fo_pct_str = f"{round(fo_pct, 2)}%"

    total_minutes_trunc = int(minutes_played_total_sec // 60)
    total_minutes_rounded = int(round(minutes_played_total_sec / 60.0))
    avg_minutes_played_min = (
        round((total_minutes_rounded / games), 2) if games > 0 else 0.0
    )

    takeaways_total = hits_success + interceptions
    giveaways_total = schta_total + passes_intercepted_against
    to_total = giveaways_total - takeaways_total

    ppp_total = ppg_total + ppa_total
    pkp_total = pkg_total + pka_total
    p_per_20 = (
        round((points_total * 1200.0) / minutes_played_total_sec, 2)
        if minutes_played_total_sec > 0
        else 0.0
    )

    overall_rows.append(
        [
            player_name,
            team_final,
            position_str,
            int(games),
            int(wins_total),
            int(losses_total),
            int(otl_total),
            int(goals_total),
            int(assists_total),
            int(points_total),
            int(plus_minus_total),
            int(pen_minutes),
            int(hits_all),
            int(hits_taken_all),
            int(shots_total),
            shot_pct_str,
            int(shots_blocked),
            int(total_minutes_trunc),
            float(avg_minutes_played_min),
            int(ppg_total),
            int(ppa_total),
            int(ppp_total),
            int(pkg_total),
            int(pka_total),
            int(pkp_total),
            int(gwg_total),
            int(gtg_total),
            fo_pct_str,
            int(fo_taken_total),
            float(p_per_20),
            int(s1_total),
            int(s2_total),
            int(s3_total),
            int(hits_success),
            int(schta_total),
            int(takeaways_total),
            int(giveaways_total),
            int(to_total),
            int(retrievals),
            int(interceptions),
            int(passes_intercepted_against),
        ]
    )

    for team_name, acc in team_acc.items():
        gp_t = acc["GP"]
        g_t = acc["G"]
        a_t = acc["A"]
        p_t = g_t + a_t
        pm_t = acc["PM"]
        pim_t = acc["PIM"]
        hit_t = acc["HIT"]
        htt_t = acc["HTT"]
        sht_t = acc["SHT"]
        sb_t = acc["SB"]
        mp_t = acc["MP"]
        scht_t = acc["SCHT"]
        schta_t = acc["SCHTA"]
        pret_t = acc["PRET"]
        pi_t = acc["PI"]
        pia_t = acc["PIA"]
        ppg_t = acc["PPG"]
        ppa_t = acc["PPA"]
        ppp_t = ppg_t + ppa_t
        pkg_t = acc["PKG"]
        pka_t = acc["PKA"]
        pkp_t = pkg_t + pka_t
        p_per_20_t = round((p_t * 1200.0) / mp_t, 2) if mp_t > 0 else 0.0
        gwg_t = acc["GWG"]
        gtg_t = acc["GTG"]
        s1_t = acc["S1"]
        s2_t = acc["S2"]
        s3_t = acc["S3"]

        shot_pct_t = (g_t / sht_t * 100.0) if sht_t > 0 else 0.0
        shot_pct_t_str = f"{round(shot_pct_t, 2)}%"

        mp_trunc_t = int(mp_t // 60)
        mp_round_t = int(round(mp_t / 60.0))
        amg_t = round((mp_round_t / gp_t), 2) if gp_t > 0 else 0.0

        ta_t = scht_t + pi_t
        ga_t = schta_t + pia_t
        to_t = ga_t - ta_t

        fow_t = acc["FOW"]
        fot_t = acc["FOT"]
        fo_pct_t = (fow_t / fot_t * 100.0) if fot_t > 0 else 0.0
        fo_pct_t_str = f"{round(fo_pct_t, 2)}%"

        team_split_rows.append(
            [
                player_name,
                team_name,
                position_str,
                int(gp_t),
                int(acc["W"]),
                int(acc["L"]),
                int(acc["OTL"]),
                int(g_t),
                int(a_t),
                int(p_t),
                int(pm_t),
                int(pim_t),
                int(hit_t),
                int(htt_t),
                int(sht_t),
                shot_pct_t_str,
                int(sb_t),
                int(mp_trunc_t),
                float(amg_t),
                int(ppg_t),
                int(ppa_t),
                int(ppp_t),
                int(pkg_t),
                int(pka_t),
                int(pkp_t),
                int(gwg_t),
                int(gtg_t),
                fo_pct_t_str,
                int(fot_t),
                float(p_per_20_t),
                int(s1_t),
                int(s2_t),
                int(s3_t),
                int(scht_t),
                int(schta_t),
                int(ta_t),
                int(ga_t),
                int(to_t),
                int(pret_t),
                int(pi_t),
                int(pia_t),
            ]
        )

overall_columns = [
    "Player Name",
    "Team Name",
    "POS",
    "GP",
    "W",
    "L",
    "OTL",
    "G",
    "A",
    "P",
    "+/-",
    "PIM",
    "HIT",
    "HTT",
    "SHT",
    "SHT%",
    "SB",
    "MP",
    "AMG",
    "PPG",
    "PPA",
    "PPP",
    "PKG",
    "PKA",
    "PKP",
    "GW",
    "GT",
    "FO%",
    "FOT",
    "P/20",
    "S1",
    "S2",
    "S3",
    "SCHT",
    "SCHTA",
    "TA",
    "GA",
    "TO",
    "PRET",
    "PI",
    "PIA",
]
team_columns = [
    "Player Name",
    "Team Name",
    "POS",
    "GP",
    "W",
    "L",
    "OTL",
    "G",
    "A",
    "P",
    "+/-",
    "PIM",
    "HIT",
    "HTT",
    "SHT",
    "SHT%",
    "SB",
    "MP",
    "AMG",
    "PPG",
    "PPA",
    "PPP",
    "PKG",
    "PKA",
    "PKP",
    "GW",
    "GT",
    "FO%",
    "FOT",
    "P/20",
    "S1",
    "S2",
    "S3",
    "SCHT",
    "SCHTA",
    "TA",
    "GA",
    "TO",
    "PRET",
    "PI",
    "PIA",
]


# ---------------- MVP_v3 (balanced, position-standardized, stats-only) ----------------
def add_mvp_v3_balanced(
    df: pd.DataFrame, col_name: str = "MVP_v3", *, parity_align: bool = True
) -> pd.DataFrame:
    """
    Position-standardized MVP:
      - Offense: P/20
      - Power play: PPP per 20
      - Puck play: (TA - GA), SCHT - SCHTA, PI - PIA, SB per 20
      - Physical: HIT - 0.5*HTT per 20
      - Retrievals: PRET per 20
      - Discipline: -PIM per 20
      - Clutch: 5/3/1 star weighting + GW/GT per MP
      - Faceoffs: boosted for C, tiny for W, none for D
    All per-20 stats divide by MP (minutes played), with safe zeros.
    Z-scores computed within buckets {C, W, D}.
    """

    import numpy as np
    import pandas as pd

    df = df.copy()

    # Safe numerics
    mp = df["MP"].replace(0, np.nan)
    gp = df["GP"].replace(0, np.nan)
    amg = df["AMG"].replace(0, np.nan)

    def per20(col):
        return (20.0 * df[col] / mp).replace([np.inf, -np.inf], np.nan).fillna(0.0)

    # Inputs
    P20 = df["P/20"].astype(float).replace([np.inf, -np.inf], 0.0).fillna(0.0)
    PPP20 = per20("PPP")
    TA20 = per20("TA")
    GA20 = per20("GA")
    SCHT20 = per20("SCHT")
    SCHTA20 = per20("SCHTA")
    PI20 = per20("PI")
    PIA20 = per20("PIA")
    SB20 = per20("SB")
    HIT20 = per20("HIT")
    HTT20 = per20("HTT")
    PRET20 = per20("PRET")
    PIM20 = per20("PIM")

    StarPM = (
        ((5 * df["S1"] + 3 * df["S2"] + df["S3"]) / df["MP"])
        .replace([np.inf, -np.inf], 0.0)
        .fillna(0.0)
    )
    GWpm = (df["GW"] / df["MP"]).replace([np.inf, -np.inf], 0.0).fillna(0.0)
    GTpm = (df["GT"] / df["MP"]).replace([np.inf, -np.inf], 0.0).fillna(0.0)

    # Bucket by position (C/W/D)
    def _bucket(row):
        cats = extract_pos_categories(str(row.get("POS", "")))
        if "D" in cats:
            return "D"
        if "C" in cats:
            return "C"
        if "LW" in cats or "RW" in cats:
            return "W"
        try:
            if row["GP"] > 0 and (row["FOT"] / row["GP"]) >= 5:
                return "C"  # fallback: heavy draw-takers are centers
        except Exception:
            pass
        return "W"

    pos = df.apply(_bucket, axis=1)

    def z_by_pos(series, cap=2.5):
        z = pd.Series(index=series.index, dtype=float)
        for p in ("C", "W", "D"):
            idx = pos == p
            s = series[idx]
            mu = s.mean()
            sd = s.std(ddof=0)
            z.loc[idx] = (s - mu) / sd if sd and sd > 1e-9 else 0.0
        return z.clip(-cap, cap).fillna(0.0)

    Z_OFF = z_by_pos(P20)
    Z_PP = z_by_pos(PPP20)
    Z_PLAY = z_by_pos((TA20 - GA20) + (SCHT20 - SCHTA20) + (PI20 - PIA20))
    Z_BLOCK = z_by_pos(SB20)
    Z_PHYS = z_by_pos(HIT20 - 0.5 * HTT20)
    Z_RETR = z_by_pos(PRET20)
    Z_DISC = z_by_pos(-PIM20)
    Z_CLUTCH = z_by_pos(0.7 * GWpm + 0.2 * GTpm + 0.3 * StarPM)

    # Faceoffs: parse percent safely
    fo_pct = (
        df["FO%"]
        .astype(str)
        .str.replace("%", "", regex=False)
        .replace("", np.nan)
        .astype(float)
        .clip(lower=0, upper=100)
        .fillna(50.0)
    )
    draws_pg = (df["FOT"] / gp).replace([np.inf, -np.inf], np.nan).fillna(0.0)
    fo_vol = np.sqrt(draws_pg.clip(lower=0.0))  # diminishing returns
    fo_norm = (20.0 / amg).replace([np.inf, -np.inf], np.nan).fillna(0.0)
    fo_core = ((fo_pct - 50.0) / 10.0) * fo_vol * fo_norm  # ±1 per meaningful edge

    fo_scale = pos.map({"C": 1.0, "W": 0.2, "D": 0.0}).astype(float)
    FO_TERM = (fo_scale * fo_core).clip(-1.5, 1.5).fillna(0.0)

    # Weights per bucket (small, stable differences)
    W = {
        "C": {
            "off": 1.55,
            "pp": 0.55,
            "play": 1.05,
            "block": 0.25,
            "phys": 0.25,
            "disc": 0.15,
            "retr": 0.30,
            "clutch": 0.55,
            "fo": 1.00,
        },
        "W": {
            "off": 1.65,
            "pp": 0.65,
            "play": 0.85,
            "block": 0.25,
            "phys": 0.30,
            "disc": 0.15,
            "retr": 0.30,
            "clutch": 0.55,
            "fo": 0.20,
        },
        "D": {
            "off": 1.05,
            "pp": 0.45,
            "play": 1.30,
            "block": 0.80,
            "phys": 0.25,
            "disc": 0.20,
            "retr": 0.35,
            "clutch": 0.45,
            "fo": 0.00,
        },
    }

    mvp = []
    for i in df.index:
        p = pos.loc[i]
        w = W[p]
        score = (
            w["off"] * Z_OFF.loc[i]
            + w["pp"] * Z_PP.loc[i]
            + w["play"] * Z_PLAY.loc[i]
            + w["block"] * Z_BLOCK.loc[i]
            + w["phys"] * Z_PHYS.loc[i]
            + w["disc"] * Z_DISC.loc[i]
            + w["retr"] * Z_RETR.loc[i]
            + w["clutch"] * Z_CLUTCH.loc[i]
            + w["fo"] * FO_TERM.loc[i]
        )
        mvp.append(score)

    df[col_name] = pd.Series(mvp, index=df.index)

    if parity_align:
        # Light centering so each team’s average is ~0 (prevents single-team skew)
        df[col_name] = df.groupby("Team Name")[col_name].transform(
            lambda s: s - s.mean()
        )

    return df


def add_goalie_mvp_v1(
    df: pd.DataFrame, col_name: str = "MVP_vG", *, parity_align: bool = False
) -> pd.DataFrame:
    """
    League-standardized goalie MVP.
      Inputs expected in df: PCT (SV%), GAA, GP, W, SO, MP (minutes), PIM, GA, SA, SAR, PS% (float/str),
      S1,S2,S3.
    - Builds stable per-60/ per-GP rates with safe zeros
    - Z-scores across the entire league of goalies (not per team)
    - Optional parity_align per team is OFF by default for goalies (turning it on caused your zeros)
    """
    import numpy as _np

    d = df.copy()

    # Safe numerics
    def _num(s, default=0.0):
        return pd.to_numeric(d.get(s, default), errors="coerce").fillna(default)

    GP = _num("GP")
    MPm = _num("MP")  # MP is stored as minutes
    MPs = MPm * 60.0
    SVP = _num("PCT")
    GAA = _num("GAA")
    W = _num("W")
    SO = _num("SO")
    PIM = _num("PIM")
    GA = _num("GA")
    SA = _num("SA")
    SAR = _num("SAR")
    # PS% might be a string like "0.800"
    PSR = pd.to_numeric(d.get("PS%", 0.0), errors="coerce").fillna(0.0)

    # Derived rates
    def safe_div(n, den):
        out = n.copy()
        den = den.replace(0, _np.nan)
        out = n / den
        return out.fillna(0.0).replace([_np.inf, -_np.inf], 0.0)

    Wpm = safe_div(W, MPs)  # wins per minute
    SOpm = safe_div(SO, MPs)  # shutouts per minute
    StarPM = safe_div(
        5 * _num("S1") + 3 * _num("S2") + _num("S3"), MPs
    )  # stars per minute

    SAR60 = safe_div(SAR * 3600.0, MPs)  # per 60 minutes
    PIM60 = safe_div(PIM * 60.0, MPs)  # penalties per 60 (to penalize)
    GA60 = safe_div(GA * 3600.0, MPs)  # essentially GAA with the exact same units

    # Z-score helper across ALL goalies (not per team)
    def z(series, cap=3.0):
        mu = series.mean()
        sd = series.std(ddof=0)
        if sd and sd > 1e-9:
            out = (series - mu) / sd
        else:
            out = series * 0.0
        return out.clip(-cap, cap).fillna(0.0)

    # Components (higher is better unless prefixed by minus)
    Z_SV = z(SVP)  # save%
    Z_GAA = z(-GAA)  # lower GAA is better
    Z_W = z(Wpm)  # wins per minute
    Z_SO = z(SOpm)  # shutouts per minute
    Z_SAR = z(SAR60)  # rebounds controlled (proxy for workload/shot quality mgmt)
    Z_PS = z(PSR)  # shootout/penalty-shot performance
    Z_DISC = z(-PIM60)  # discipline
    Z_GA60 = z(-GA60)  # extra stability on goals allowed rate
    Z_STAR = z(StarPM)  # clutch

    # Weights (tuned to give a bit more spread; adjust if you want more goalie presence in Top MVPs)
    score = (
        1.50 * Z_SV
        + 1.10 * Z_GAA
        + 0.80 * Z_W
        + 0.50 * Z_SO
        + 0.60 * Z_SAR
        + 0.50 * Z_PS
        + 0.30 * Z_DISC
        + 0.80 * Z_GA60
        + 0.40 * Z_STAR
    )

    d[col_name] = score

    # DO NOT center by team when there is only one recognized goalie; it collapses them to 0
    if parity_align:
        d[col_name] = d.groupby("Team Name")[col_name].transform(lambda s: s - s.mean())

    return d


overall_df = pd.DataFrame(overall_rows, columns=overall_columns)
team_df = pd.DataFrame(team_split_rows, columns=team_columns)

# --- Add MVP and sort ---
overall_df = add_mvp_v3_balanced(overall_df, col_name="MVP_v3", parity_align=True)
overall_df = overall_df.sort_values("MVP_v3", ascending=False).reset_index(drop=True)
team_df = add_mvp_v3_balanced(team_df, col_name="MVP_v3", parity_align=True)

overall_df = overall_df.sort_values("MVP_v3", ascending=False)
# We’ll sort per-team sheets individually when writing.


# ---------------- ROSTER ATTRIBUTES SCRAPER ----------------
def find_roster_table(
    soup: BeautifulSoup,
) -> Optional[Tuple[List[str], List[List[str]]]]:
    tables = soup.find_all("table")
    for tbl in tables:
        header_cells = []
        thead = tbl.find("thead")
        if thead:
            tr = thead.find("tr")
            if tr:
                header_cells = [
                    normalize_ws(c.get_text(" ")) for c in tr.find_all(["th", "td"])
                ]
        if not header_cells:
            tr = tbl.find("tr")
            if tr:
                header_cells = [
                    normalize_ws(c.get_text(" ")) for c in tr.find_all(["th", "td"])
                ]
        if not header_cells:
            continue
        header_set = {h.lower() for h in header_cells}
        if "player name" in header_set and "ck" in header_set and "ov" in header_set:
            data_rows: List[List[str]] = []
            for tr in tbl.find_all("tr"):
                cells = tr.find_all("td")
                if not cells:
                    continue
                row = [normalize_ws(td.get_text(" ")) for td in cells]
                if row:
                    data_rows.append(row)
            if data_rows:
                return header_cells, data_rows
    return None


def parse_roster_attributes_from_html(html: bytes) -> Dict[str, Dict[str, int]]:
    soup = BeautifulSoup(html, "lxml")  # <<< switch parser
    found = find_roster_table(soup)
    if not found:
        return {}
    headers, rows = found
    col_idx = {h: i for i, h in enumerate(headers)}
    required = ["Player Name"] + ATTR_COLUMNS
    for k in required:
        if k not in col_idx:
            for h, i in list(col_idx.items()):
                if h.lower() == k.lower():
                    col_idx[k] = i
                    break
    out: Dict[str, Dict[str, int]] = {}
    for row in rows:
        if len(row) < max(col_idx.values(), default=0) + 1:
            continue
        name_raw = row[col_idx.get("Player Name", 1)]
        if not name_raw or "TEAM AVERAGE" in name_raw.upper():
            continue
        name_clean = re.sub(r"\([^)]*\)", "", name_raw)
        name_clean = normalize_ws(name_clean)
        if not name_clean:
            continue
        attrs: Dict[str, int] = {}
        ok = True
        for k in ATTR_COLUMNS:
            idx = col_idx.get(k)
            if idx is None or idx >= len(row):
                ok = False
                break
            v = row[idx]
            try:
                val = int(round(float(v)))
            except Exception:
                val = 0
            attrs[k] = val
        if not ok:
            continue
        out[canonical_name(name_clean)] = attrs
    return out


def find_goalie_roster_table(
    soup: BeautifulSoup,
) -> Optional[Tuple[List[str], List[List[str]]]]:
    tables = soup.find_all("table")
    for tbl in tables:
        header_cells = []
        thead = tbl.find("thead")
        if thead and thead.find("tr"):
            tr = thead.find("tr")
            header_cells = [
                normalize_ws(td.get_text(" ")) for td in tr.find_all(["th", "td"])
            ]
        if not header_cells:
            tr = tbl.find("tr")
            if tr:
                header_cells = [
                    normalize_ws(td.get_text(" ")) for td in tr.find_all(["th", "td"])
                ]
        if not header_cells:
            continue
        lower = [h.lower() for h in header_cells]
        if ("goalie name" in lower) and ("ov" in [h.lower() for h in header_cells]):
            data_rows: List[List[str]] = []
            for tr in tbl.find_all("tr"):
                cells = tr.find_all("td")
                if not cells:
                    continue
                row = [normalize_ws(td.get_text(" ")) for td in cells]
                if row:
                    data_rows.append(row)
            if data_rows:
                return header_cells, data_rows
    return None


def parse_goalie_roster_attributes_from_html(html: bytes) -> Dict[str, Dict[str, int]]:
    soup = BeautifulSoup(html, "lxml")  # <<< switch parser
    found = find_goalie_roster_table(soup)
    if not found:
        return {}
    headers, rows = found
    idx_by_header = {h: i for i, h in enumerate(headers)}
    name_idx = None
    for h, i in list(idx_by_header.items()):
        if h.lower() == "goalie name":
            name_idx = i
            break
    if name_idx is None:
        name_idx = 0
    col_idx: Dict[str, int] = {}
    for k in GOALIE_ATTR_COLUMNS:
        match = [i for h, i in idx_by_header.items() if h.strip().lower() == k.lower()]
        if match:
            col_idx[k] = match[0]
    out: Dict[str, Dict[str, int]] = {}
    for row in rows:
        if name_idx >= len(row):
            continue
        raw_name = row[name_idx]
        if not raw_name or "TEAM AVERAGE" in raw_name.upper():
            continue
        name_clean = re.sub(r"\([^)]*\)", "", raw_name)
        name_clean = normalize_ws(name_clean)
        if not name_clean:
            continue
        fallback_start = name_idx + 2
        attrs: Dict[str, int] = {}
        ok_any = False
        for j, k in enumerate(GOALIE_ATTR_COLUMNS):
            if k in col_idx and col_idx[k] < len(row):
                v = row[col_idx[k]]
            else:
                jj = fallback_start + j
                if jj < len(row):
                    v = row[jj]
                else:
                    v = "0"
            try:
                val = int(round(float(v)))
            except Exception:
                val = 0
            attrs[k] = val
            if val:
                ok_any = True
        if ok_any:
            out[canonical_name(name_clean)] = attrs
    return out


def scrape_all_roster_attributes() -> (
    Tuple[Dict[str, Dict[str, int]], Dict[str, Dict[str, int]]]
):
    t0_roster = time.perf_counter()
    print("Discovering roster pages and scraping attributes (skaters + goalies)...")
    sk_attr_by_name: Dict[str, Dict[str, int]] = {}
    g_attr_by_name: Dict[str, Dict[str, int]] = {}
    misses = 0
    for team_id in range(1, TEAM_PAGE_MAX_ID + 1):
        url = f"https://vhlportal.com/vhlm/{SEASON}/ProTeam.php?Team={team_id}"
        try:
            r = http_get_with_retries(url)
        except Exception:
            misses += 1
            if misses >= STOP_AFTER_CONSEC_FAILS:
                break
            continue
        if r.status_code != 200 or not r.content:
            misses += 1
            if misses >= STOP_AFTER_CONSEC_FAILS:
                break
            continue
        misses = 0
        page_sk = parse_roster_attributes_from_html(r.content) or {}
        for cname, attrs in page_sk.items():
            if cname not in sk_attr_by_name:
                sk_attr_by_name[cname] = attrs
        page_g = parse_goalie_roster_attributes_from_html(r.content) or {}
        for cname, attrs in page_g.items():
            if cname not in g_attr_by_name:
                g_attr_by_name[cname] = attrs
        time.sleep(0.05)  # be kind
    print(
        f"Collected attributes for ~{len(sk_attr_by_name)} skaters and ~{len(g_attr_by_name)} goalies from roster pages."
    )
    print(
        f"[TIME] scrape_all_roster_attributes: {time.perf_counter() - t0_roster:.2f}s"
    )
    return sk_attr_by_name, g_attr_by_name


# ---------------- Players Info + player FINAL page (trained skills) ----------------
def extract_name_to_intermediate_links_from_players_info(
    html: bytes, base_url: str
) -> Dict[str, str]:
    soup = BeautifulSoup(html, "lxml")  # <<< switch parser
    out: Dict[str, str] = {}
    for tbl in soup.find_all("table"):
        any_row = False
        for tr in tbl.find_all("tr"):
            tds = tr.find_all("td")
            if not tds:
                continue
            first_td = tds[0]
            a = first_td.find("a", href=True)
            if not a:
                continue
            name_raw = normalize_ws(a.get_text(" "))
            if not name_raw or "team average" in name_raw.lower():
                continue
            name_clean = re.sub(r"\([^)]*\)", "", name_raw)
            name_clean = normalize_ws(name_clean)
            cname = canonical_name(name_clean)
            if not cname:
                continue
            href = a["href"]
            abs_url = requests.compat.urljoin(base_url, href)
            if cname not in out:
                out[cname] = abs_url
                any_row = True
        if any_row:
            return out
    return out


def extract_final_player_link_from_intermediate_page(
    html: bytes, base_url: str
) -> Optional[str]:
    soup = BeautifulSoup(html, "lxml")  # <<< switch parser
    for tbl in soup.find_all("table"):
        header_tr = None
        for tr in tbl.find_all("tr"):
            if tr.find("th"):
                header_tr = tr
                break
        if not header_tr:
            header_tr = tbl.find("tr")
            if not header_tr:
                continue
        headers = [
            normalize_ws(x.get_text(" ")).lower()
            for x in header_tr.find_all(["th", "td"])
        ]
        if not headers:
            continue
        try:
            link_idx = headers.index("link")
        except ValueError:
            link_idx = None
        has_position = any(h in ("position",) for h in headers)
        has_age = any(h in ("age",) for h in headers)
        looks_like_meta = has_position and (has_age or len(headers) >= 5)
        if not looks_like_meta and link_idx is None:
            continue
        started = False
        for tr in tbl.find_all("tr"):
            if not started:
                if tr is header_tr:
                    started = True
                continue
            tds = tr.find_all("td")
            if not tds:
                continue
            cell = (
                tds[-1] if (link_idx is None or link_idx >= len(tds)) else tds[link_idx]
            )
            a = cell.find("a", href=True)
            if a:
                href = a["href"]
                return requests.compat.urljoin(base_url, href)
    a_txt = soup.find("a", string=re.compile(r"^\s*Link\s*$", re.I))
    if a_txt and a_txt.get("href"):
        return requests.compat.urljoin(base_url, a_txt["href"])
    for a in soup.find_all("a", href=True):
        href = a["href"]
        lch = href.lower()
        if any(
            k in lch
            for k in ["player", "players.php", "player.php", "playerpage", "playerid"]
        ):
            return requests.compat.urljoin(base_url, href)
    return None


CODE_SET = set(TRAINED_SKILL_COLUMNS)
SYNONYMS: Dict[str, str] = {}


def _add_synonyms(code: str, names: List[str]):
    for n in names:
        SYNONYMS[re.sub(r"[^a-z0-9]+", "", n.lower())] = code


_add_synonyms("DK", ["DK", "Deking", "Deke", "Dekes"])
_add_synonyms("SH", ["SH", "Shooting"])
_add_synonyms("PA", ["PA", "Passing"])
_add_synonyms("BC", ["BC", "Body Checking", "BodyChecking", "Checking"])
_add_synonyms("GR", ["GR", "Getting Open", "GettingOpen"])
_add_synonyms("FO", ["FO", "Faceoffs", "Face-Offs", "Face Offs"])
_add_synonyms(
    "PC", ["PC", "Puck Control", "PuckControl", "Stickhandling", "Stick Handling"]
)
_add_synonyms(
    "DC",
    [
        "DC",
        "Defensive Coverage",
        "Defense Coverage",
        "Def Coverage",
        "Defence Coverage",
    ],
)
_add_synonyms("OV", ["OV", "OVR", "Overall"])
_add_synonyms("SP", ["SP", "Skating", "Speed", "Skating Speed"])
_add_synonyms("SS", ["SS", "Slap Shot", "SlapShot"])
_add_synonyms("WS", ["WS", "Wrist Shot", "WristShot", "Wrist"])
_add_synonyms("LD", ["LD", "Leadership"])
_add_synonyms("FG", ["FG", "Fighting"])
_add_synonyms("PO", ["PO", "Positioning"])
_add_synonyms("EX", ["EX", "Experience"])


def _parse_ratings_table(soup: BeautifulSoup) -> Tuple[Dict[str, int], Optional[int]]:
    skills = {f"{TR_PREFIX}{k}": 0 for k in TRAINED_SKILL_COLUMNS}
    tpa_val: Optional[int] = None
    tbl = soup.find("table", id="ratings")
    if not tbl:
        tbl = soup.find("table", attrs={"class": re.compile(r"\bratings\b", re.I)})
    if not tbl:
        return skills, tpa_val
    hdrs: List[str] = []
    thead = tbl.find("thead")
    if thead:
        hdrs = [normalize_ws(th.get_text(" ")) for th in thead.find_all("th")]
    if not hdrs:
        first_tr = tbl.find("tr")
        if first_tr:
            hdrs = [normalize_ws(th.get_text(" ")) for th in first_tr.find_all("th")]
    body = tbl.find("tbody") or tbl
    cells = body.find_all(["td", "th"])
    vals: List[Optional[int]] = []
    for c in cells:
        raw = normalize_ws(c.get_text(" "))
        if (c.get("id") or "").strip().upper() == "TPA":
            try:
                tpa_val = int(raw.replace(",", ""))
            except Exception:
                pass
        m = re.fullmatch(r"\d{1,3}", raw)
        if m:
            vals.append(int(raw))
    n = min(len(hdrs), len(vals))
    for i in range(n):
        lab = hdrs[i].strip().upper()
        if lab == "TPA":
            continue
        if lab in TRAINED_SKILL_COLUMNS:
            skills[f"{TR_PREFIX}{lab}"] = int(vals[i])
    return skills, tpa_val


def _parse_goalie_ratings_table(soup: BeautifulSoup) -> Dict[str, int]:
    gskills = {f"{GTR_PREFIX}{k}": 0 for k in GOALIE_TRAINED_SKILL_COLUMNS}
    best_tbl = None
    best_hits = -1
    for tbl in soup.find_all("table"):
        hdrs = []
        thead = tbl.find("thead")
        if thead and thead.find("tr"):
            hdrs = [normalize_ws(x.get_text(" ")) for x in thead.find_all("th")]
        if not hdrs:
            first_tr = tbl.find("tr")
            if first_tr:
                hdrs = [
                    normalize_ws(x.get_text(" "))
                    for x in first_tr.find_all(["th", "td"])
                ]
        if not hdrs:
            continue
        hits = sum(1 for h in hdrs if h.strip().upper() in GOALIE_TRAINED_SKILL_COLUMNS)
        if hits > best_hits:
            best_hits = hits
            best_tbl = tbl
    if not best_tbl or best_hits <= 0:
        return gskills
    hdrs = []
    thead = best_tbl.find("thead")
    if thead and thead.find("tr"):
        hdrs = [normalize_ws(x.get_text(" ")) for x in thead.find_all("th")]
    if not hdrs:
        first_tr = best_tbl.find("tr")
        if first_tr:
            hdrs = [
                normalize_ws(x.get_text(" ")) for x in first_tr.find_all(["th", "td"])
            ]
    body = best_tbl.find("tbody") or best_tbl
    vals = []
    for c in body.find_all(["td", "th"]):
        raw = normalize_ws(c.get_text(" "))
        if re.fullmatch(r"\d{1,3}", raw):
            vals.append(int(raw))
    vi = 0
    for h in hdrs:
        lab = h.strip().upper()
        if lab in GOALIE_TRAINED_SKILL_COLUMNS and vi < len(vals):
            gskills[f"{GTR_PREFIX}{lab}"] = int(vals[vi])
            vi += 1
    return gskills


def parse_player_training_page(
    html: bytes,
) -> Tuple[Dict[str, int], str, Optional[int]]:
    soup = BeautifulSoup(html, "lxml")  # <<< switch parser
    text = soup.get_text(separator="\n")
    user_name = ""
    tpe_val: Optional[int] = None
    m_user = re.search(r"User\s*:\s*([^\|\n\r]+)", text, re.IGNORECASE)
    if m_user:
        user_name = normalize_ws(m_user.group(1))
    m_tpe = re.search(r"\bTPE\s*:\s*([\d,]+)", text, re.IGNORECASE)
    if m_tpe:
        try:
            tpe_val = int(m_tpe.group(1).replace(",", ""))
        except Exception:
            tpe_val = None

    out = {f"{TR_PREFIX}{k}": 0 for k in TRAINED_SKILL_COLUMNS}
    out.update({f"{GTR_PREFIX}{k}": 0 for k in GOALIE_TRAINED_SKILL_COLUMNS})

    sk_skills, _ = _parse_ratings_table(soup)
    for k, v in sk_skills.items():
        out[k] = v

    g_skills = _parse_goalie_ratings_table(soup)
    for k, v in g_skills.items():
        out[k] = v

    return out, user_name, tpe_val


def scrape_all_trained_skills(target_cnames: Set[str]) -> Dict[str, Dict[str, object]]:
    t0_trained = time.perf_counter()
    print(
        "Discovering Players Info pages and scraping trained skills (filtered by NameList)..."
    )
    name_to_intermediate: Dict[str, str] = {}
    misses = 0
    for team_id in range(1, TEAM_PAGE_MAX_ID + 1):
        team_url = f"https://vhlportal.com/vhlm/{SEASON}/ProTeam.php?Team={team_id}"
        try:
            r = http_get_with_retries(team_url)
        except Exception:
            misses += 1
            if misses >= STOP_AFTER_CONSEC_FAILS:
                break
            continue
        if r.status_code != 200 or not r.content:
            misses += 1
            if misses >= STOP_AFTER_CONSEC_FAILS:
                break
            continue
        misses = 0
        soup = BeautifulSoup(r.content, "lxml")  # <<< switch parser
        info_url = find_players_info_url(soup, team_url)
        if not info_url:
            continue
        try:
            r_info = http_get_with_retries(info_url)
        except Exception:
            continue
        if r_info.status_code != 200 or not r.content:
            continue
        links_map = extract_name_to_intermediate_links_from_players_info(
            r_info.content, info_url
        )
        for cname, inter_url in links_map.items():
            if cname in target_cnames and cname not in name_to_intermediate:
                name_to_intermediate[cname] = inter_url
    print(
        f"Found {len(name_to_intermediate)} intermediate player pages to follow (from {len(target_cnames)} targets)."
    )

    def _scrape_one(
        cname: str, inter_url: str
    ) -> Optional[Tuple[str, Dict[str, object]]]:
        try:
            r_mid = http_get_with_retries(inter_url)
            if r_mid.status_code != 200 or not r_mid.content:
                return None
            final_url = extract_final_player_link_from_intermediate_page(
                r_mid.content, inter_url
            )
            if not final_url:
                return None
            rp = http_get_with_retries(final_url)
            if rp.status_code != 200 or not rp.content:
                return None
            skills_map, user_name, tpe_val = parse_player_training_page(rp.content)
            out = dict(skills_map)
            out["User"] = user_name or ""
            out["TPE"] = int(tpe_val or 0)
            return (cname, out)
        except Exception:
            return None

    trained: Dict[str, Dict[str, object]] = {}
    with ThreadPoolExecutor(max_workers=min(PARALLEL_WORKERS, 12)) as ex:
        futures = {
            ex.submit(_scrape_one, cname, url): cname
            for cname, url in name_to_intermediate.items()
        }
        done = 0
        total = len(futures)
        for fut in as_completed(futures):
            done += 1
            if done % 25 == 1 or done == total:
                print(f"  scraping player pages... {done}/{total}")
            res = fut.result()
            if res:
                cname, out = res
                trained[cname] = out
    print(f"Collected trained skills for ~{len(trained)} players (targets only).")
    print(f"[TIME] scrape_all_trained_skills: {time.perf_counter() - t0_trained:.2f}s")
    return trained


# -------- Scrape roster attributes --------
sk_attr_by_name, g_attr_by_name = scrape_all_roster_attributes()

# -------- Merge attributes into SKATERS (OVERALL/TEAM) --------
overall_df["__norm_name__"] = overall_df["Player Name"].apply(canonical_name)
team_df["__norm_name__"] = team_df["Player Name"].apply(canonical_name)

rows_attrs = []
for cname, amap in sk_attr_by_name.items():
    row = {"__norm_name__": cname}
    row.update(amap)
    rows_attrs.append(row)

attrs_df = (
    pd.DataFrame(rows_attrs)
    if rows_attrs
    else pd.DataFrame(columns=["__norm_name__"] + ATTR_COLUMNS)
)
for c in ATTR_COLUMNS:
    if c not in attrs_df.columns:
        attrs_df[c] = 0

overall_df = overall_df.merge(attrs_df, on="__norm_name__", how="left")
for c in ATTR_COLUMNS:
    overall_df[c] = overall_df[c].fillna(0).astype(int)
overall_df = overall_df.drop(columns=["__norm_name__"])
overall_df = overall_df[overall_columns + ATTR_COLUMNS]

team_df = team_df.merge(attrs_df, on="__norm_name__", how="left")
for c in ATTR_COLUMNS:
    team_df[c] = team_df[c].fillna(0).astype(int)
team_df = team_df.drop(columns=["__norm_name__"])
team_columns_with_attrs = team_columns + ATTR_COLUMNS
team_df = team_df[team_columns_with_attrs]

# ----------------  GOALIE AGGREGATION (overall + team)  ----------------
goalie_overall_rows = []
goalie_team_rows = []

goalie_columns = [
    "Goalie Name",
    "Team Name",
    "GP",
    "W",
    "L",
    "OTL",
    "PCT",
    "GAA",
    "MP",
    "PIM",
    "SO",
    "GA",
    "SA",
    "SAR",
    "A",
    "EG",
    "PS%",
    "PSA",
    "S1",
    "S2",
    "S3",
]

# Filter bot goalies: only those with a verified goalie POS from Players Info (via name_to_position)
goalie_targets_set: Set[str] = set()
for disp, pos in name_to_position.items():
    if is_goalie_position(pos):
        goalie_targets_set.add(canonical_name(disp))

# Build goalie name list from pages, but keep only verified (filters out bot goalies)
goalie_display_names: Set[str] = set()
for umap in page_goalie_team_map.values():
    for gname in umap.keys():
        if canonical_name(gname) in goalie_targets_set:
            goalie_display_names.add(gname)

for idx, gname in enumerate(sorted(goalie_display_names), start=1):
    print(f"Processing goalie: {gname} ({idx}/{len(goalie_display_names)})")
    gp = w = l = otl = pim = so = ga = sa = sar = a = eg = 0
    mp_sec = 0
    ps_sv = 0
    ps_att = 0
    s1 = s2 = s3 = 0
    last_team = ""

    g_team_acc = defaultdict(
        lambda: {
            "GP": 0,
            "W": 0,
            "L": 0,
            "OTL": 0,
            "MP": 0,
            "PIM": 0,
            "SO": 0,
            "GA": 0,
            "SA": 0,
            "SAR": 0,
            "A": 0,
            "EG": 0,
            "PS_SV": 0,
            "PSA": 0,
            "S1": 0,
            "S2": 0,
            "S3": 0,
        }
    )

    for url in page_goalie_team_map.keys():
        tmap = page_goalie_team_map.get(url, {})
        if gname not in tmap:
            continue

        team = tmap.get(gname, "")

        # Prefer a canonical full team name for this goalie if we only saw an abbr here
        if _looks_abbr(team):
            team = goalie_default_team.get(gname, team)

        team = normalize_ws(team)
        last_team = team or last_team
        gp += 1

        if team:
            acc = g_team_acc[team]
            acc["GP"] += 1

            mp_add = int(page_goalie_mp_map.get(url, {}).get(gname, 0))
            ga_add = int(page_goalie_ga_map.get(url, {}).get(gname, 0))
            sa_add = int(page_goalie_sa_map.get(url, {}).get(gname, 0))

            pim_from_table = int(page_goalie_pim_map.get(url, {}).get(gname, 0))
            pim_from_pen = int(page_goalie_pim_from_pen_map.get(url, {}).get(gname, 0))
            pim_add = pim_from_table if pim_from_table > 0 else pim_from_pen

            w_add = int(page_goalie_w_map.get(url, {}).get(gname, 0))
            l_add = int(page_goalie_l_map.get(url, {}).get(gname, 0))
            otl_add = int(page_goalie_otl_map.get(url, {}).get(gname, 0))
            so_add = int(page_goalie_so_map.get(url, {}).get(gname, 0))
            a_add = int(page_goalie_a_map.get(url, {}).get(gname, 0))
            eg_add = int(page_goalie_eg_map.get(url, {}).get(gname, 0))
            sar_add = int(page_goalie_sar_map.get(url, {}).get(gname, 0))
            pssv_add = int(page_goalie_ps_sv_map.get(url, {}).get(gname, 0))
            psatt_add = int(page_goalie_ps_att_map.get(url, {}).get(gname, 0))

            # >>> DEBUG per-page contribution
            if (
                DEBUG_GOALIE_NAME
                and gname == DEBUG_GOALIE_NAME
                and (psatt_add or pssv_add)
            ):
                print(
                    f"[DEBUG AGG PS] {gname} :: {url} :: PSA={psatt_add}, PSSV={pssv_add}"
                )
                DEBUG_PS_ROWS.append(
                    {
                        "game_no": game_no_from_url(url),
                        "url": url,
                        "phase": "AGG_PAGE",
                        "goalie": gname,
                        "psa": int(psatt_add),
                        "pssv": int(pssv_add),
                    }
                )

            mp_sec += mp_add
            ga += ga_add
            sa += sa_add
            pim += pim_add
            w += w_add
            l += l_add
            otl += otl_add
            so += so_add
            a += a_add
            eg += eg_add
            sar += sar_add
            ps_sv += pssv_add
            ps_att += psatt_add

            s1_add = fast_get(page_s1_by_player.get(url, {}), gname)
            s2_add = fast_get(page_s2_by_player.get(url, {}), gname)
            s3_add = fast_get(page_s3_by_player.get(url, {}), gname)
            s1 += s1_add
            s2 += s2_add
            s3 += s3_add

            acc["MP"] += mp_add
            acc["GA"] += ga_add
            acc["SA"] += sa_add
            acc["PIM"] += pim_add
            acc["W"] += w_add
            acc["L"] += l_add
            acc["OTL"] += otl_add
            acc["SO"] += so_add
            acc["A"] += a_add
            acc["EG"] += eg_add
            acc["SAR"] += sar_add
            acc["PS_SV"] += pssv_add
            acc["PSA"] += psatt_add
            acc["S1"] += s1_add
            acc["S2"] += s2_add
            acc["S3"] += s3_add

    pct = ((sa - ga) / sa) if sa > 0 else 0.0
    gaa = (ga * 3600.0 / mp_sec) if mp_sec > 0 else 0.0
    ps_ratio = (ps_sv / ps_att) if ps_att > 0 else 0.0
    ps_str = f"{ps_ratio:.3f}"

    if DEBUG_GOALIE_NAME and gname == DEBUG_GOALIE_NAME:
        print(
            f"[DEBUG AGG TOTAL] {gname} :: PSA={ps_att}, PSSV={ps_sv} :: PS%={ps_str}"
        )
        DEBUG_PS_ROWS.append(
            {
                "game_no": 0,
                "url": "",
                "phase": "AGG_TOTAL",
                "goalie": gname,
                "psa": int(ps_att),
                "pssv": int(ps_sv),
            }
        )

    mp_min_trunc = int(mp_sec // 60)

    goalie_overall_rows.append(
        [
            gname,
            last_team,
            int(gp),
            int(w),
            int(l),
            int(otl),
            round(pct, 3),
            round(gaa, 2),
            int(mp_min_trunc),
            int(pim),
            int(so),
            int(ga),
            int(sa),
            int(sar),
            int(a),
            int(eg),
            ps_str,
            int(ps_att),
            int(s1),
            int(s2),
            int(s3),
        ]
    )

    for t, acc in g_team_acc.items():
        mp_min_t = int(acc["MP"] // 60)
        pct_t = ((acc["SA"] - acc["GA"]) / acc["SA"]) if acc["SA"] > 0 else 0.0
        gaa_t = (acc["GA"] * 3600.0 / acc["MP"]) if acc["MP"] > 0 else 0.0
        ps_ratio_t = (acc["PS_SV"] / acc["PSA"]) if acc["PSA"] > 0 else 0.0
        ps_str_t = f"{ps_ratio_t:.3f}"
        goalie_team_rows.append(
            [
                gname,
                t,
                int(acc["GP"]),
                int(acc["W"]),
                int(acc["L"]),
                int(acc["OTL"]),
                round(pct_t, 3),
                round(gaa_t, 2),
                int(mp_min_t),
                int(acc["PIM"]),
                int(acc["SO"]),
                int(acc["GA"]),
                int(acc["SA"]),
                int(acc["SAR"]),
                int(acc["A"]),
                int(acc["EG"]),
                ps_str_t,
                int(acc["PSA"]),
                int(acc["S1"]),
                int(acc["S2"]),
                int(acc["S3"]),
            ]
        )

goalie_overall_df = pd.DataFrame(goalie_overall_rows, columns=goalie_columns)
goalie_team_df = pd.DataFrame(goalie_team_rows, columns=goalie_columns)


# ------------- PS DEBUG CSV DUMP -------------
if DEBUG_GOALIE_NAME and DEBUG_PS_ROWS:
    try:
        df_ps = pd.DataFrame(DEBUG_PS_ROWS)
        # prefer a nice phase ordering for readability
        phase_order = [
            "TABLE",
            "SO_ANY",
            "SO_FPP",
            "PS_REG/OT",
            "COMBINED",
            "PRE-MERGE",
            "POST-MERGE",
            "AGG_PAGE",
            "AGG_TOTAL",
        ]
        if "phase" in df_ps.columns:
            df_ps["phase"] = pd.Categorical(
                df_ps["phase"], categories=phase_order, ordered=True
            )
        # sort by game then phase (AGG_TOTAL has game_no 0 and will sort last if you want; adjust as you like)
        sort_cols = [c for c in ["game_no", "phase"] if c in df_ps.columns]
        if sort_cols:
            df_ps = df_ps.sort_values(sort_cols, kind="mergesort")
        df_ps.to_csv(DEBUG_PS_CSV_PATH, index=False)
        print(f"Saved PS debug log to {DEBUG_PS_CSV_PATH}")
    except Exception as e:
        print(f"[PS DEBUG CSV WARN] Could not write CSV: {e}")

# ---------- Merge goalie roster attributes ----------
if not goalie_overall_df.empty:
    goalie_overall_df["__norm_name__"] = goalie_overall_df["Goalie Name"].apply(
        canonical_name
    )

    g_rows_attrs = []
    for cname, amap in g_attr_by_name.items():
        row = {"__norm_name__": cname}
        for c in GOALIE_ATTR_COLUMNS:
            row[c] = int(amap.get(c, 0))
        g_rows_attrs.append(row)

    g_attrs_df = (
        pd.DataFrame(g_rows_attrs)
        if g_rows_attrs
        else pd.DataFrame(columns=["__norm_name__"] + GOALIE_ATTR_COLUMNS)
    )

    for c in GOALIE_ATTR_COLUMNS:
        if c not in g_attrs_df.columns:
            g_attrs_df[c] = 0

    goalie_overall_df = goalie_overall_df.merge(
        g_attrs_df, on="__norm_name__", how="left"
    )
    for c in GOALIE_ATTR_COLUMNS:
        goalie_overall_df[c] = goalie_overall_df[c].fillna(0).astype(int)
    goalie_overall_df = goalie_overall_df.drop(columns=["__norm_name__"])

if not goalie_team_df.empty:
    goalie_team_df["__norm_name__"] = goalie_team_df["Goalie Name"].apply(
        canonical_name
    )

    g_rows_attrs = []
    for cname, amap in g_attr_by_name.items():
        row = {"__norm_name__": cname}
        for c in GOALIE_ATTR_COLUMNS:
            row[c] = int(amap.get(c, 0))
        g_rows_attrs.append(row)

    g_attrs_df = (
        pd.DataFrame(g_rows_attrs)
        if g_rows_attrs
        else pd.DataFrame(columns=["__norm_name__"] + GOALIE_ATTR_COLUMNS)
    )

    for c in GOALIE_ATTR_COLUMNS:
        if c not in g_attrs_df.columns:
            g_attrs_df[c] = 0

    goalie_team_df = goalie_team_df.merge(g_attrs_df, on="__norm_name__", how="left")
    for c in GOALIE_ATTR_COLUMNS:
        goalie_team_df[c] = goalie_team_df[c].fillna(0).astype(int)
    goalie_team_df = goalie_team_df.drop(columns=["__norm_name__"])

# ---------- Optional: GTG debug CSV ----------
if DEBUG_GTG and gtg_debug_records:
    try:
        pd.DataFrame(gtg_debug_records).to_csv(DEBUG_GTG_CSV_PATH, index=False)
        print(f"Saved GTG debug to {DEBUG_GTG_CSV_PATH}")
    except Exception as e:
        print(f"[GTG DEBUG CSV WARN] {e}")

# ---------- Trained skills scrape & merge (TR_* skaters, GTR_* goalies, plus User/TPE) ----------
# Build target canonical-name set from the data we just aggregated
sk_cnames = (
    set(overall_df["Player Name"].map(canonical_name))
    if not overall_df.empty
    else set()
)
g_cnames = (
    set(goalie_overall_df["Goalie Name"].map(canonical_name))
    if not goalie_overall_df.empty
    else set()
)
target_cnames = sk_cnames | g_cnames

trained_by_cname = scrape_all_trained_skills(target_cnames)

# Normalize to a DataFrame
tr_rows = []
for cname, m in (trained_by_cname or {}).items():
    row = {"__norm_name__": cname}
    # copy known skill keys, default 0
    for k in TR_COLUMNS + GTR_COLUMNS:
        row[k] = int(m.get(k, 0))
    row["User"] = str(m.get("User", "") or "")
    row["TPE"] = int(m.get("TPE", 0) or 0)
    tr_rows.append(row)

trained_df = (
    pd.DataFrame(tr_rows)
    if tr_rows
    else pd.DataFrame(
        columns=["__norm_name__"] + TR_COLUMNS + GTR_COLUMNS + ["User", "TPE"]
    )
)

# Ensure all columns exist even if scraping found nothing
for k in TR_COLUMNS + GTR_COLUMNS:
    if k not in trained_df.columns:
        trained_df[k] = 0
if "User" not in trained_df.columns:
    trained_df["User"] = ""
if "TPE" not in trained_df.columns:
    trained_df["TPE"] = 0

# --- Merge into SKATERS (overall + team): keep TR_* + User + TPE
if not overall_df.empty:
    overall_df["__norm_name__"] = overall_df["Player Name"].apply(canonical_name)
    sk_tr = trained_df[["__norm_name__"] + TR_COLUMNS + ["User", "TPE"]].copy()
    overall_df = overall_df.merge(sk_tr, on="__norm_name__", how="left")
    for c in TR_COLUMNS:
        overall_df[c] = overall_df[c].fillna(0).astype(int)
    overall_df["User"] = overall_df["User"].fillna("")
    overall_df["TPE"] = overall_df["TPE"].fillna(0).astype(int)
    overall_df = overall_df.drop(columns=["__norm_name__"])

if not team_df.empty:
    team_df["__norm_name__"] = team_df["Player Name"].apply(canonical_name)
    sk_tr = trained_df[["__norm_name__"] + TR_COLUMNS + ["User", "TPE"]].copy()
    team_df = team_df.merge(sk_tr, on="__norm_name__", how="left")
    for c in TR_COLUMNS:
        team_df[c] = team_df[c].fillna(0).astype(int)
    team_df["User"] = team_df["User"].fillna("")
    team_df["TPE"] = team_df["TPE"].fillna(0).astype(int)
    team_df = team_df.drop(columns=["__norm_name__"])

# --- Merge into GOALIES (overall + team): keep GTR_* + User + TPE
if not goalie_overall_df.empty:
    goalie_overall_df["__norm_name__"] = goalie_overall_df["Goalie Name"].apply(
        canonical_name
    )
    gk_tr = trained_df[["__norm_name__"] + GTR_COLUMNS + ["User", "TPE"]].copy()
    goalie_overall_df = goalie_overall_df.merge(gk_tr, on="__norm_name__", how="left")
    for c in GTR_COLUMNS:
        goalie_overall_df[c] = goalie_overall_df[c].fillna(0).astype(int)
    goalie_overall_df["User"] = goalie_overall_df["User"].fillna("")
    goalie_overall_df["TPE"] = goalie_overall_df["TPE"].fillna(0).astype(int)
    goalie_overall_df = goalie_overall_df.drop(columns=["__norm_name__"])

if not goalie_team_df.empty:
    goalie_team_df["__norm_name__"] = goalie_team_df["Goalie Name"].apply(
        canonical_name
    )
    gk_tr = trained_df[["__norm_name__"] + GTR_COLUMNS + ["User", "TPE"]].copy()
    goalie_team_df = goalie_team_df.merge(gk_tr, on="__norm_name__", how="left")
    for c in GTR_COLUMNS:
        goalie_team_df[c] = goalie_team_df[c].fillna(0).astype(int)
    goalie_team_df["User"] = goalie_team_df["User"].fillna("")
    goalie_team_df["TPE"] = goalie_team_df["TPE"].fillna(0).astype(int)
    goalie_team_df = goalie_team_df.drop(columns=["__norm_name__"])

# ---------- Column ordering so trained skills/User/TPE appear in Excel ----------
# Ensure MVP exists before we lock column order
if "MVP_v3" not in overall_df.columns:
    overall_df = add_mvp_v3_balanced(overall_df, col_name="MVP_v3", parity_align=True)
if "MVP_v3" not in team_df.columns:
    team_df = add_mvp_v3_balanced(team_df, col_name="MVP_v3", parity_align=True)

# Goalies: add POS = 'G'
for gdf in (goalie_overall_df, goalie_team_df):
    if gdf is not None and not gdf.empty:
        if "POS" not in gdf.columns:
            gdf["POS"] = "G"

# Present trained/extra columns
sk_extra_cols = TR_COLUMNS + ["User", "TPE"]
gk_extra_cols = GTR_COLUMNS + ["User", "TPE"]

present_sk_extra_overall = [c for c in sk_extra_cols if c in overall_df.columns]
present_sk_extra_team = [c for c in sk_extra_cols if c in team_df.columns]

present_gk_extra_overall = [c for c in gk_extra_cols if c in goalie_overall_df.columns]
present_gk_extra_team = [c for c in gk_extra_cols if c in goalie_team_df.columns]


# --- Skaters: User 2nd; MVP_v3 right after POS
def _skater_order(base_cols, df):
    # Start with desired head: Name, User, Team, POS, MVP
    head = ["Player Name", "User", "Team Name", "POS", "MVP_v3"]
    # Then the rest of the original stat columns (excluding the ones we just placed)
    tail_stats = [c for c in base_cols if c not in {"Player Name", "Team Name", "POS"}]
    # Attributes and trained columns (exclude User here to avoid duplication; keep TPE)
    tr_cols = [c for c in sk_extra_cols if c in df.columns and c != "User"]
    # Build final list, only keeping columns that exist
    desired = (
        [c for c in head if c in df.columns]
        + [c for c in tail_stats if c in df.columns]
        + [c for c in ATTR_COLUMNS if c in df.columns]
        + tr_cols
    )
    # Append any leftover columns to be safe
    leftovers = [c for c in df.columns if c not in desired]
    return desired + leftovers


overall_df = overall_df[_skater_order(overall_columns, overall_df)]
team_df = team_df[_skater_order(team_columns, team_df)]


# --- Goalies: Name, User, Team, POS, then the rest
def _goalie_order(base_cols, df):
    head = ["Goalie Name", "User", "Team Name", "POS"]
    tail_stats = [c for c in base_cols if c not in {"Goalie Name", "Team Name"}]
    tr_cols = [c for c in gk_extra_cols if c in df.columns and c != "User"]
    desired = (
        [c for c in head if c in df.columns]
        + [c for c in tail_stats if c in df.columns]
        + [c for c in GOALIE_ATTR_COLUMNS if c in df.columns]
        + tr_cols
    )
    leftovers = [c for c in df.columns if c not in desired]
    return desired + leftovers


goalie_overall_df = goalie_overall_df[_goalie_order(goalie_columns, goalie_overall_df)]
goalie_team_df = goalie_team_df[_goalie_order(goalie_columns, goalie_team_df)]


# ---------- Inline per-minute columns (right after each stat) ----------
def _ensure_per_minute_inline(
    df: pd.DataFrame, minutes_col: str, stats: list[str], *, suffix="/M", decimals=3
) -> pd.DataFrame:
    """
    Create/update per-minute columns right after each base stat. Idempotent:
    - If <stat>/M exists, update it and move it right after <stat>.
    - If it doesn't, insert it.
    """
    dfo = df.copy()
    if minutes_col not in dfo.columns:
        return dfo

    m = pd.to_numeric(dfo[minutes_col], errors="coerce").replace(0, pd.NA)

    for col in stats:
        if col not in dfo.columns:
            continue

        per_name = f"{col}{suffix}"
        base = pd.to_numeric(dfo[col], errors="coerce")
        per = base.divide(m).fillna(0)  # Now dividing by minutes instead of games
        if decimals is not None:
            per = per.round(decimals)

        target_idx = dfo.columns.get_loc(col) + 1

        if per_name in dfo.columns:
            # Update values
            dfo[per_name] = per
            # Move next to its base stat if needed
            cur_idx = dfo.columns.get_loc(per_name)
            if cur_idx != target_idx:
                series = dfo.pop(per_name)
                dfo.insert(target_idx, per_name, series)
        else:
            dfo.insert(target_idx, per_name, per)

    return dfo


# ==== GOALIE MVP + CROSS-ROLE NORMALIZATION + COLUMN ORDER + PER-MINUTE (insert above goalie per-minute) ====

# Ensure POS column for goalies (right after Team Name)
for _df in (goalie_overall_df, goalie_team_df):
    if (
        _df is not None
        and not _df.empty
        and "POS" not in _df.columns
        and "Team Name" in _df.columns
    ):
        _df.insert(_df.columns.get_loc("Team Name") + 1, "POS", "G")


# Goalie MVP (balanced; team-parity aligned like skaters)
def add_goalie_mvp(
    df: pd.DataFrame, col_name: str = "MVP_G", *, parity_align: bool = True
) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    d = df.copy()

    gp = pd.to_numeric(d.get("GP", 0), errors="coerce").replace(0, np.nan)
    mp = pd.to_numeric(d.get("MP", 0), errors="coerce").replace(0, np.nan)
    pct = pd.to_numeric(d.get("PCT", 0.0), errors="coerce").fillna(0.0)
    gaa = pd.to_numeric(d.get("GAA", np.nan), errors="coerce")
    psr = pd.to_numeric(d.get("PS%", 0.0), errors="coerce").fillna(0.0)
    so = pd.to_numeric(d.get("SO", 0), errors="coerce").fillna(0.0)
    w = pd.to_numeric(d.get("W", 0), errors="coerce").fillna(0.0)
    pim = pd.to_numeric(d.get("PIM", 0), errors="coerce").fillna(0.0)
    sar = pd.to_numeric(d.get("SAR", 0), errors="coerce").fillna(0.0)
    s1 = pd.to_numeric(d.get("S1", 0), errors="coerce").fillna(0.0)
    s2 = pd.to_numeric(d.get("S2", 0), errors="coerce").fillna(0.0)
    s3 = pd.to_numeric(d.get("S3", 0), errors="coerce").fillna(0.0)

    sar60 = (sar / mp * 60.0).replace([np.inf, -np.inf], np.nan).fillna(0.0)
    pim60 = (pim / mp * 60.0).replace([np.inf, -np.inf], np.nan).fillna(0.0)
    so_pg = (so / gp).replace([np.inf, -np.inf], np.nan).fillna(0.0)
    w_pg = (w / gp).replace([np.inf, -np.inf], np.nan).fillna(0.0)
    stars_pg = (
        ((5 * s1 + 3 * s2 + s3) / gp).replace([np.inf, -np.inf], np.nan).fillna(0.0)
    )

    def z(series):
        s = series.astype(float)
        mu = s.mean()
        sd = s.std(ddof=0)
        if sd and sd > 1e-9:
            out = (s - mu) / sd
        else:
            out = s * 0.0
        return out.clip(-3, 3).fillna(0.0)

    score = (
        2.0 * z(pct)
        + 1.5 * z(-gaa)
        + 0.7 * z(psr)
        + 0.6 * z(w_pg)
        + 0.6 * z(so_pg)
        + 0.6 * z(stars_pg)
        + 0.4 * z(-sar60)
        + 0.2 * z(-pim60)
    )
    d[col_name] = score
    if parity_align and "Team Name" in d.columns:
        d[col_name] = d.groupby("Team Name")[col_name].transform(lambda s: s - s.mean())
    return d


goalie_overall_df = add_goalie_mvp(goalie_overall_df, "MVP_G", parity_align=True)
goalie_team_df = add_goalie_mvp(goalie_team_df, "MVP_G", parity_align=True)

goalie_overall_df = add_goalie_mvp_v1(
    goalie_overall_df, col_name="MVP_vG", parity_align=False
)
goalie_team_df = add_goalie_mvp_v1(
    goalie_team_df, col_name="MVP_vG", parity_align=False
)


def _zcol(s: pd.Series) -> pd.Series:
    mu = s.mean()
    sd = s.std(ddof=0)
    return (s - mu) / sd if sd and sd > 1e-9 else s * 0.0


# skaters
overall_df["MVP_Z"] = _zcol(overall_df["MVP_v3"])
# goalies
goalie_overall_df["MVP_Z"] = _zcol(goalie_overall_df["MVP_vG"])


# Cross-role normalization so we can pick a Team MVP across skaters & goalies
def _std_norm(s: pd.Series) -> pd.Series:
    s = pd.to_numeric(s, errors="coerce")
    mu = s.mean()
    sd = s.std(ddof=0)
    return (
        ((s - mu) / sd).fillna(0.0)
        if (sd and sd > 1e-9)
        else pd.Series(0.0, index=s.index)
    )


if "MVP_v3" in overall_df.columns:
    overall_df["MVP_norm"] = _std_norm(overall_df["MVP_v3"])
if "MVP_G" in goalie_overall_df.columns:
    goalie_overall_df["MVP_norm"] = _std_norm(goalie_overall_df["MVP_G"])


# Column placement: User after Name; MVP after POS
def _reorder_skaters_cols(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    cols = list(df.columns)

    def _move(c, after):
        if c in cols and after in cols:
            cols.insert(cols.index(after) + 1, cols.pop(cols.index(c)))

    if "Player Name" in cols and "User" in cols:
        _move("User", "Player Name")
    if "POS" in cols and "MVP_v3" in cols:
        _move("MVP_v3", "POS")
    return df[cols]


def _reorder_goalie_cols(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    cols = list(df.columns)

    def _move(c, after):
        if c in cols and after in cols:
            cols.insert(cols.index(after) + 1, cols.pop(cols.index(c)))

    if "Goalie Name" in cols and "User" in cols:
        _move("User", "Goalie Name")
    if "POS" in cols and "MVP_G" in cols:
        _move("MVP_G", "POS")
    return df[cols]


overall_df = _reorder_skaters_cols(overall_df)
team_df = _reorder_skaters_cols(team_df)
goalie_overall_df = _reorder_goalie_cols(goalie_overall_df)
goalie_team_df = _reorder_goalie_cols(goalie_team_df)


# Inline per-minute helper + SKATER per-minute columns
def _add_per_minute_inline(
    df: pd.DataFrame, mp_col: str, stats: list, *, suffix="/M", decimals=2
) -> pd.DataFrame:
    if df is None or df.empty or mp_col not in df.columns:
        return df
    dfo = df.copy()
    denom = pd.to_numeric(dfo[mp_col], errors="coerce").replace(0, np.nan).astype(float)
    for col in stats:
        if col not in dfo.columns:
            continue
        num = pd.to_numeric(dfo[col], errors="coerce").astype(float)
        per_minute = (num / denom).round(decimals).fillna(0.0)
        dfo.insert(dfo.columns.get_loc(col) + 1, f"{col}{suffix}", per_minute)
    return dfo


# Skaters
_skater_stats_for_pm = [
    "G",
    "A",
    "P",
    "+/-",
    "PIM",
    "HIT",
    "HTT",
    "SHT",
    "SB",
    "PPG",
    "PPA",
    "PPP",
    "PKG",
    "PKA",
    "PKP",
    "GW",
    "GT",
    "FOT",
    "S1",
    "S2",
    "S3",
    "SCHT",
    "SCHTA",
    "TA",
    "GA",
    "TO",
    "PRET",
    "PI",
    "PIA",
]
overall_df = _ensure_per_minute_inline(overall_df, "MP", _skater_stats_for_pm)
team_df = _ensure_per_minute_inline(team_df, "MP", _skater_stats_for_pm)

# Goalies
_goalie_stats_for_pm = [
    "W",
    "L",
    "OTL",
    "PIM",
    "GA",
    "SA",
    "SAR",
    "PSA",
    "S1",
    "S2",
    "S3",
]  # Removed MP since we're using it as denominator
goalie_overall_df = _ensure_per_minute_inline(
    goalie_overall_df, "MP", _goalie_stats_for_pm
)
goalie_team_df = _ensure_per_minute_inline(goalie_team_df, "MP", _goalie_stats_for_pm)


# ---------- Excel export helpers ----------
import re as _re


def _sanitize_sheet_name(name: str) -> str:
    s = _re.sub(r"[:\\/?*\[\]]", " ", str(name))
    return s[:31] if len(s) > 31 else s


def _pick_excel_engine() -> Optional[str]:
    try:
        import openpyxl  # noqa: F401

        return "openpyxl"
    except Exception:
        try:
            import xlsxwriter  # noqa: F401

            return "xlsxwriter"
        except Exception:
            return None


def export_excel_with_team_tabs(
    overall_df: pd.DataFrame,
    team_df: pd.DataFrame,
    goalie_overall_df: Optional[pd.DataFrame],
    goalie_team_df: Optional[pd.DataFrame],
    path: str,
    team_mvp_skaters_df: Optional[pd.DataFrame] = None,
    team_mvp_goalies_df: Optional[pd.DataFrame] = None,
):
    engine = _pick_excel_engine()
    used_sheet_names: set[str] = set()

    def _unique_sheet_name(proposed: str) -> str:
        base = _sanitize_sheet_name(proposed)
        name = base or "Sheet1"
        i = 1
        while name in used_sheet_names:
            name = _sanitize_sheet_name(f"{base[:28]}_{i}")
            i += 1
        used_sheet_names.add(name)
        return name

    def _strip_per_minute(df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return df
        per_minute_cols = [
            col for col in df.columns if isinstance(col, str) and col.endswith("/M")
        ]
        return df.drop(columns=per_minute_cols) if per_minute_cols else df

    with pd.ExcelWriter(path, engine=engine) as xw:
        # ---- Skaters: All teams combined -----------------------------------
        sk_source: Optional[pd.DataFrame] = None
        if team_df is not None and not team_df.empty:
            sk_source = team_df.copy()
        elif overall_df is not None and not overall_df.empty:
            sk_source = overall_df.copy()

        if sk_source is None or sk_source.empty:
            sk_all = pd.DataFrame()
        else:
            if "MVP_v3" not in sk_source.columns:
                sk_source = add_mvp_v3_balanced(
                    sk_source, col_name="MVP_v3", parity_align=True
                )
            sk_source["MVP_v3"] = pd.to_numeric(
                sk_source.get("MVP_v3", 0.0), errors="coerce"
            ).fillna(0.0)

            sort_cols: List[str] = []
            ascending: List[bool] = []
            if "Team Name" in sk_source.columns:
                sk_source["Team Name"] = sk_source["Team Name"].fillna("")
                sort_cols.append("Team Name")
                ascending.append(True)
            sort_cols.append("MVP_v3")
            ascending.append(False)

            sk_all = sk_source.sort_values(sort_cols, ascending=ascending).reset_index(
                drop=True
            )

        drop_skater_cols = [c for c in ("MVP_v3",) if c in sk_all.columns]
        if drop_skater_cols:
            sk_all = sk_all.drop(columns=drop_skater_cols)

        sk_all = _strip_per_minute(sk_all)

        sk_all.to_excel(
            xw, sheet_name=_unique_sheet_name("Skaters - All Teams"), index=False
        )

        # ---- Goalies: All teams combined -----------------------------------
        goalie_source: Optional[pd.DataFrame] = None
        if goalie_team_df is not None and not goalie_team_df.empty:
            goalie_source = goalie_team_df.copy()
        elif goalie_overall_df is not None and not goalie_overall_df.empty:
            goalie_source = goalie_overall_df.copy()

        if goalie_source is None or goalie_source.empty:
            g_all = pd.DataFrame()
        else:
            sort_cols_g: List[str] = []
            ascending_g: List[bool] = []
            if "Team Name" in goalie_source.columns:
                goalie_source["Team Name"] = goalie_source["Team Name"].fillna("")
                sort_cols_g.append("Team Name")
                ascending_g.append(True)
            if {"PCT", "GAA"}.issubset(goalie_source.columns):
                goalie_source["_PCT"] = pd.to_numeric(
                    goalie_source["PCT"], errors="coerce"
                ).fillna(0.0)
                goalie_source["_GAA"] = pd.to_numeric(
                    goalie_source["GAA"], errors="coerce"
                ).fillna(999.0)
                sort_cols_g.extend(["_PCT", "_GAA"])
                ascending_g.extend([False, True])
            elif "MP" in goalie_source.columns:
                goalie_source["MP"] = pd.to_numeric(
                    goalie_source["MP"], errors="coerce"
                ).fillna(0)
                sort_cols_g.append("MP")
                ascending_g.append(False)

            if sort_cols_g:
                g_all = goalie_source.sort_values(
                    sort_cols_g, ascending=ascending_g
                ).reset_index(drop=True)
            else:
                g_all = goalie_source.reset_index(drop=True)

            for col in ["_PCT", "_GAA"]:
                if col in g_all.columns:
                    g_all = g_all.drop(columns=col)

        drop_goalie_cols = [c for c in ("MVP_G", "MVP_vG") if c in g_all.columns]
        if drop_goalie_cols:
            g_all = g_all.drop(columns=drop_goalie_cols)

        g_all = _strip_per_minute(g_all)

        g_all.to_excel(
            xw, sheet_name=_unique_sheet_name("Goalies - All Teams"), index=False
        )

    def _zscore(series: pd.Series) -> pd.Series:
        s = pd.to_numeric(series, errors="coerce").fillna(0.0)
        mu = float(s.mean())
        sd = float(s.std(ddof=0))
        if sd <= 1e-9:
            return pd.Series(0.0, index=s.index)
        return (s - mu) / sd


# --- Skaters block
sk = overall_df.copy()
# Use MVP_v3 as the base; make a Z-score for skaters if not present
if "MVP_Z" not in sk.columns:
    sk["MVP_Z"] = _zscore(sk["MVP_v3"])
sk_long = sk.rename(columns={"Player Name": "Name"})[
    [
        c
        for c in ["Team Name", "Name", "User", "POS", "MVP_v3", "MVP_Z"]
        if c in sk.columns
    ]
].copy()
sk_long["Role"] = "Skater"
sk_long["Raw_MVP"] = sk_long.get("MVP_v3", 0.0)

# --- Goalies block
g = goalie_overall_df.copy()
# If you already computed a goalie MVP column (e.g., 'MVP_G'), use it.
goalie_mvp_col = "MVP_G" if "MVP_G" in g.columns else None
if goalie_mvp_col is None:
    # Minimal fallback so code never crashes; feel free to replace with your full formula.
    # Higher PCT, lower GAA, more SO and SAR are good.
    g_pct = pd.to_numeric(g.get("PCT", 0.0), errors="coerce").fillna(0.0)
    g_gaa = pd.to_numeric(g.get("GAA", 0.0), errors="coerce").fillna(0.0)
    g_so = pd.to_numeric(g.get("SO", 0), errors="coerce").fillna(0.0)
    g_sar = pd.to_numeric(g.get("SAR", 0), errors="coerce").fillna(0.0)
    g["__MVP_FALLBACK__"] = 3.0 * g_pct - 0.5 * g_gaa + 0.05 * g_so + 0.01 * g_sar
    goalie_mvp_col = "__MVP_FALLBACK__"

# Z-score within goalies if not present
if "MVP_Z" not in g.columns:
    g["MVP_Z"] = _zscore(g[goalie_mvp_col])

g_long = g.rename(columns={"Goalie Name": "Name"})[
    [
        c
        for c in ["Team Name", "Name", "User", goalie_mvp_col, "MVP_Z"]
        if c in g.columns
    ]
].copy()
g_long["POS"] = "G"
g_long["Role"] = "Goalie"
g_long["Raw_MVP"] = g_long[goalie_mvp_col]
g_long.drop(
    columns=[c for c in [goalie_mvp_col] if c != "Raw_MVP" and c in g_long.columns],
    inplace=True,
)

# --- Combine and pick top per team
combo = pd.concat([sk_long, g_long], ignore_index=True, sort=False)
# Make sure we don’t lose anyone due to NaNs
combo["MVP_Z"] = pd.to_numeric(combo["MVP_Z"], errors="coerce").fillna(0.0)
combo["Raw_MVP"] = pd.to_numeric(combo["Raw_MVP"], errors="coerce").fillna(0.0)

top_team_mvp = (
    combo.sort_values(["Team Name", "MVP_Z"], ascending=[True, False])
    .groupby("Team Name", as_index=False)
    .head(1)
    .reset_index(drop=True)
)
# (Optional) nice column order if present
cols_pref = ["Team Name", "Name", "User", "POS", "Role", "MVP_Z", "Raw_MVP"]
top_team_mvp = top_team_mvp[[c for c in cols_pref if c in top_team_mvp.columns]]


def _zscore_1d(s: pd.Series) -> pd.Series:
    s = pd.to_numeric(s, errors="coerce").fillna(0.0)
    mu = float(s.mean())
    sd = float(s.std(ddof=0))
    return pd.Series(0.0, index=s.index) if sd <= 1e-9 else (s - mu) / sd


# --- Skaters with z-scored MVP_v3 ------------------------------------------
_sk = overall_df.copy()
if "MVP_v3" not in _sk.columns:
    _sk = add_mvp_v3_balanced(_sk, col_name="MVP_v3", parity_align=True)
_sk["MVP_Z"] = _zscore_1d(_sk["MVP_v3"])
_sk["__ROLE__"] = "S"
_sk_cols_order = list(_sk.columns)  # preserve full skater columns

# --- Goalies with z-scored MVP_G -------------------------------------------
_g = goalie_overall_df.copy()
if "MVP_G" not in _g.columns:
    g_pct = pd.to_numeric(_g.get("PCT", 0.0), errors="coerce").fillna(0.0)
    g_gaa = pd.to_numeric(_g.get("GAA", 0.0), errors="coerce").fillna(0.0)
    g_so = pd.to_numeric(_g.get("SO", 0), errors="coerce").fillna(0.0)
    g_sar = pd.to_numeric(_g.get("SAR", 0), errors="coerce").fillna(0.0)
    _g["MVP_G"] = 3.0 * g_pct - 0.5 * g_gaa + 0.05 * g_so + 0.01 * g_sar
_g["MVP_Z"] = _zscore_1d(_g["MVP_G"])
_g["__ROLE__"] = "G"
_g_cols_order = list(_g.columns)  # preserve full goalie columns

# --- One winner per team: compare skater-vs-goalie MVP_Z --------------------
all_teams = sorted(
    set(_sk["Team Name"].dropna().unique()) | set(_g["Team Name"].dropna().unique())
)

winners_s_rows = []
winners_g_rows = []

for team in all_teams:
    s_best = _sk[_sk["Team Name"] == team].sort_values("MVP_Z", ascending=False).head(1)
    g_best = _g[_g["Team Name"] == team].sort_values("MVP_Z", ascending=False).head(1)

    candidates = []
    if not s_best.empty:
        candidates.append(("S", float(s_best["MVP_Z"].iloc[0])))
    if not g_best.empty:
        candidates.append(("G", float(g_best["MVP_Z"].iloc[0])))
    if not candidates:
        continue

    winner_role = max(candidates, key=lambda t: t[1])[0]
    if winner_role == "S":
        winners_s_rows.append(s_best[_sk_cols_order].iloc[0])
    else:
        winners_g_rows.append(g_best[_g_cols_order].iloc[0])

team_mvp_skaters = pd.DataFrame(winners_s_rows, columns=_sk_cols_order).reset_index(
    drop=True
)
team_mvp_goalies = pd.DataFrame(winners_g_rows, columns=_g_cols_order).reset_index(
    drop=True
)

# ============================================================================

# ---------- Write outputs (Excel only) ----------
try:
    t0_excel = time.perf_counter()
    export_excel_with_team_tabs(
        overall_df=overall_df,
        team_df=team_df,
        goalie_overall_df=goalie_overall_df,
        goalie_team_df=goalie_team_df,
        path=output_excel_path,
        team_mvp_skaters_df=team_mvp_skaters,
        team_mvp_goalies_df=team_mvp_goalies,
    )
    print(
        f"Wrote Excel workbook to {output_excel_path} ({time.perf_counter() - t0_excel:.2f}s)"
    )
except Exception as e:
    print(f"[EXCEL WRITE WARN] {e}")

print("All done.")
print(f"[TIME] TOTAL: {time.perf_counter() - SCRIPT_START_TIME:.2f}s")
