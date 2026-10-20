"""Canonical entity resolution for Ask Data (projects, countries, months).

Two distinct jobs, deliberately kept separate:

* resolving what a USER typed (fuzzy, with clarification)      -> resolve_project / resolve_country / resolve_month
* matching LEDGER rows to projects (deterministic, never fuzzy) -> ProjectIndex

Normalisation folds case, accents, punctuation, a few known spelling variants and Hebrew (niqqud, final
letters, glued one/two-letter prefixes, known country/month terms). Aliases only map spellings to an existing
canonical project; they never modify any data.

Thresholds (fuzzy user input only): score >= 0.90 accept (if unambiguous), 0.60 <= score < 0.90 ask the user to
clarify, score < 0.60 not found.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Callable, Iterable, Optional

from config.ask_data_aliases import (
    COUNTRY_SYNONYMS,
    HEBREW_PREFIXES,
    HEBREW_TERMS,
    PROJECT_ALIASES,
    SPELLING_VARIANTS,
)

ACCEPT_THRESHOLD = 0.90
CLARIFY_THRESHOLD = 0.60
MAX_CANDIDATES = 4

ACCEPTED = "accepted"
AMBIGUOUS = "ambiguous"
NOT_FOUND = "not_found"

_FINAL_LETTERS = str.maketrans("ךםןףץ", "כמנפצ")


def _fold(value: object) -> str:
    """Case/accents/punctuation/Hebrew-points fold. No vocabulary lookups."""
    if value is None:
        return ""
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")  # accents and Hebrew niqqud
    text = text.replace("\u05f3", "").replace("\u05f4", "").replace("'", "").replace("\u2019", "")  # geresh / gershayim / apostrophes
    text = text.casefold().translate(_FINAL_LETTERS)
    text = re.sub(r"[^\w\s]|_", " ", text)
    return re.sub(r"\s+", " ", text).strip()


_HEBREW = {_fold(k): v for k, v in HEBREW_TERMS.items()}
_HEBREW_PREFIX_SET = set(_fold(HEBREW_PREFIXES))


def _has_hebrew(token: str) -> bool:
    return any("\u0590" <= ch <= "\u05ff" for ch in token)


def _translate_token(token: str) -> str:
    if token in _HEBREW:
        return _HEBREW[token]
    if _has_hebrew(token):
        for prefix_len in (1, 2):
            rest = token[prefix_len:]
            if len(token) > prefix_len + 2 and set(token[:prefix_len]) <= _HEBREW_PREFIX_SET and rest in _HEBREW:
                return _HEBREW[rest]
    return token


def normalize_text(value: object) -> str:
    """Canonical comparison text: lower-case, accent/punctuation-free, Hebrew terms translated, variants unified."""
    tokens = []
    for token in _fold(value).split(" "):
        if not token:
            continue
        token = _translate_token(token)
        tokens.append(SPELLING_VARIANTS.get(token, token))
    return " ".join(tokens)


def compact_key(value: object) -> str:
    """Space-insensitive key used for exact matching ("AD Alpha" == "ad-alpha" == "ADAlpha")."""
    return normalize_text(value).replace(" ", "")


def classify_score(score: float) -> str:
    if score >= ACCEPT_THRESHOLD:
        return ACCEPTED
    if score >= CLARIFY_THRESHOLD:
        return AMBIGUOUS
    return NOT_FOUND


@dataclass(frozen=True)
class Resolution:
    status: str
    value: Optional[str] = None
    score: float = 0.0
    candidates: tuple = ()
    reason: str = ""

    @property
    def accepted(self) -> bool:
        return self.status == ACCEPTED


# ── aliases ──────────────────────────────────────────────────────────────────────────────────────
def project_alias_map(project_names: Iterable[str], aliases: Optional[dict] = None) -> dict[str, str]:
    """alias compact-key -> canonical project name, keeping only aliases whose target project exists."""
    names_by_key: dict[str, list[str]] = {}
    for name in project_names:
        if name:
            names_by_key.setdefault(compact_key(name), []).append(name)
    resolved: dict[str, str] = {}
    for alias, canonical in (PROJECT_ALIASES if aliases is None else aliases).items():
        targets = names_by_key.get(compact_key(canonical), [])
        if len(targets) == 1:
            resolved[compact_key(alias)] = targets[0]
    return resolved


def _alias_candidates_by_canonical(project_names: Iterable[str], aliases: Optional[dict] = None) -> dict[str, list[str]]:
    """canonical project name -> its own alias source texts (only aliases whose target project exists)."""
    names_by_key: dict[str, list[str]] = {}
    for name in project_names:
        if name:
            names_by_key.setdefault(compact_key(name), []).append(name)
    groups: dict[str, list[str]] = {}
    for alias, canonical in (PROJECT_ALIASES if aliases is None else aliases).items():
        targets = names_by_key.get(compact_key(canonical), [])
        if len(targets) == 1:
            groups.setdefault(targets[0], []).append(alias)
    return groups


# ── projects ─────────────────────────────────────────────────────────────────────────────────────
def _project_score(text: str, name: str) -> float:
    candidate_key, project_key = compact_key(text), compact_key(name)
    if candidate_key == project_key:
        return 1.0
    candidate_tokens, project_tokens = set(normalize_text(text).split()), set(normalize_text(name).split())
    if candidate_tokens and candidate_tokens <= project_tokens:
        return 0.95
    if project_tokens and project_tokens <= candidate_tokens:
        return 0.92
    return SequenceMatcher(None, candidate_key, project_key).ratio()


def resolve_project(text: object, project_names: Iterable[str], aliases: Optional[dict] = None) -> Resolution:
    names = sorted({n for n in project_names if n})
    key = compact_key(text)
    if not key or not names:
        return Resolution(NOT_FOUND, reason="empty")
    alias_hit = project_alias_map(names, aliases).get(key)
    if alias_hit:
        return Resolution(ACCEPTED, alias_hit, 1.0, (), "alias")

    # A project's fuzzy score also considers its own alias source spellings (e.g. a ledger/PPS spelling that
    # differs from the project master), so a word shared ONLY with an alias -- not with any real project name --
    # still surfaces that project as a genuine competing candidate, instead of silently losing to whichever real
    # name happens to contain the word.
    alias_groups = _alias_candidates_by_canonical(names, aliases)

    def best_score(name: str) -> float:
        return max(_project_score(text, candidate) for candidate in [name] + alias_groups.get(name, []))

    scored = sorted(((best_score(n), n) for n in names), key=lambda sn: (-sn[0], sn[1]))
    top = scored[0][0]
    if top >= 1.0:
        exact = [n for s, n in scored if s >= 1.0]
        if len(exact) == 1:
            return Resolution(ACCEPTED, exact[0], 1.0, (), "exact")
        return Resolution(AMBIGUOUS, None, 1.0, tuple(exact[:MAX_CANDIDATES]), "duplicate_names")
    if top >= ACCEPT_THRESHOLD:
        peers = [n for s, n in scored if s >= ACCEPT_THRESHOLD]
        if len(peers) == 1:
            return Resolution(ACCEPTED, peers[0], top, (), "fuzzy")
        return Resolution(AMBIGUOUS, None, top, tuple(peers[:MAX_CANDIDATES]), "several_close_matches")
    if top >= CLARIFY_THRESHOLD:
        close = [n for s, n in scored if s >= CLARIFY_THRESHOLD]
        return Resolution(AMBIGUOUS, None, top, tuple(close[:MAX_CANDIDATES]), "low_confidence")
    return Resolution(NOT_FOUND, None, top, (), "no_match")


# ── countries ────────────────────────────────────────────────────────────────────────────────────
def _country_names(known_codes: Iterable[str], full_name: Optional[Callable[[str], str]]) -> dict[str, set[str]]:
    names: dict[str, set[str]] = {}
    for code in known_codes:
        if not code:
            continue
        spellings = {normalize_text(code)}
        if full_name is not None:
            spellings.add(normalize_text(full_name(code)))
        spellings.update(normalize_text(s) for s in COUNTRY_SYNONYMS.get(str(code).lower(), []))
        names[code] = {s for s in spellings if s}
    return names


def resolve_country(text: object, known_codes: Iterable[str], full_name: Optional[Callable[[str], str]] = None) -> Resolution:
    names = _country_names(known_codes, full_name)
    normalized = normalize_text(text)
    if not normalized or not names:
        return Resolution(NOT_FOUND, reason="empty")

    exact = sorted(code for code, spellings in names.items() if normalized in spellings)
    if len(exact) == 1:
        return Resolution(ACCEPTED, exact[0], 1.0, (), "exact")
    if len(exact) > 1:
        return Resolution(AMBIGUOUS, None, 1.0, tuple(exact[:MAX_CANDIDATES]), "duplicate_names")

    tokens = set(normalized.split())
    by_token = sorted(code for code, spellings in names.items() if tokens & {s for s in spellings if len(s) >= 3})
    if len(by_token) == 1:
        return Resolution(ACCEPTED, by_token[0], 0.95, (), "token")
    if len(by_token) > 1:
        return Resolution(AMBIGUOUS, None, 0.95, tuple(by_token[:MAX_CANDIDATES]), "several_close_matches")

    scored = []
    for code, spellings in names.items():
        best = max((SequenceMatcher(None, normalized, s).ratio() for s in spellings if len(s) >= 4), default=0.0)
        scored.append((best, code))
    scored.sort(key=lambda sc: (-sc[0], sc[1]))
    top = scored[0][0]
    if top >= ACCEPT_THRESHOLD:
        peers = [c for s, c in scored if s >= ACCEPT_THRESHOLD]
        if len(peers) == 1:
            return Resolution(ACCEPTED, peers[0], top, (), "fuzzy")
        return Resolution(AMBIGUOUS, None, top, tuple(peers[:MAX_CANDIDATES]), "several_close_matches")
    if top >= CLARIFY_THRESHOLD:
        return Resolution(AMBIGUOUS, None, top, tuple(c for s, c in scored if s >= CLARIFY_THRESHOLD)[:MAX_CANDIDATES], "low_confidence")
    return Resolution(NOT_FOUND, None, top, (), "no_match")


# ── months ───────────────────────────────────────────────────────────────────────────────────────
def resolve_month(text: object, months: Iterable[str]) -> Optional[str]:
    """English full/3-letter names (as before), plus 'sept' and Hebrew month names. None when not a month."""
    normalized = normalize_text(text)
    if not normalized:
        return None
    for month in months:
        full = normalize_text(month)
        if normalized == full or normalized == full[:3] or (full == "september" and normalized == "sept"):
            return month
    return None


# ── ledger rows -> projects (deterministic) ────────────────────────────────────────────────────────
class ProjectIndex:
    """Maps ledger project names to canonical projects: exact (normalised) or alias only, never fuzzy.

    A name that matches no project, or several projects equally, is reported as unassigned (None) so it can
    never be silently attributed to the wrong country.
    """

    def __init__(self, projects, aliases: Optional[dict] = None):
        self._by_key: dict[str, list] = {}
        for project in projects:
            self._by_key.setdefault(compact_key(project.project_name), []).append(project)
        names = [p.project_name for p in projects]
        self._alias_key_to_name = project_alias_map(names, aliases)

    @property
    def collisions(self) -> dict[str, list[str]]:
        return {k: [p.project_name for p in v] for k, v in self._by_key.items() if len(v) > 1}

    def project(self, ledger_name: object):
        key = compact_key(ledger_name)
        if not key:
            return None
        hits = self._by_key.get(key)
        if hits is None and key in self._alias_key_to_name:
            hits = self._by_key.get(compact_key(self._alias_key_to_name[key]))
        return hits[0] if hits and len(hits) == 1 else None

    def canonical_name(self, ledger_name: object) -> Optional[str]:
        project = self.project(ledger_name)
        return project.project_name if project is not None else None

    def country(self, ledger_name: object) -> Optional[str]:
        project = self.project(ledger_name)
        return (project.country or None) if project is not None else None
