"""Alias and vocabulary data for the Ask Data entity resolver (data only, no logic, reviewed via PR).

Aliases only ever map a user/ledger spelling to an EXISTING canonical project; they never rename or modify
ledger rows. An alias whose target project does not exist is ignored by the resolver.
"""

# alias text -> canonical project name (matched after normalisation: case, accents, punctuation, spelling variants)
PROJECT_ALIASES: dict[str, str] = {
    # Invoice 8595 (Y1, 4 extra cameras, EUR 3,112): the complementary cameras of project "CM Essen".
    "CM Essen (coplementary)": "CM Essen",
    # Case/accent variant used on a ledger row; also resolved by normalisation, listed explicitly for clarity.
    "intermarché Heusy": "Intermarché Heusy",
    # The project master misspells "Landgraaf" as "langdraaf"; historical ledger rows (invoices 8130, 8194, 8292,
    # 8302, 8310, 8434, 8495, 8590) and PPS correspondence use the correct "Landgraaf" with either "Arts" or
    # "Aarts". Approved 2026-10-20 (docs/notes/2026-10-20-alias-proposal-plus-landgraaf-arts.md); the project
    # master's own typo is NOT corrected here. "Landgraaf" alone must stay ambiguous with "Plus Landgraaf
    # Wetzels" (a different store) -- see resolver.py's alias-aware fuzzy scoring.
    "Plus Landgraaf Arts": "Plus langdraaf Aarts",
    "Plus Landgraaf Aarts": "Plus langdraaf Aarts",
}

# token-level spelling variants applied during normalisation (variant -> canonical token)
SPELLING_VARIANTS: dict[str, str] = {
    "coplementary": "complementary",
    "complementry": "complementary",
    "complimentary": "complementary",
}

# country code (lower-case, as stored on projects) -> accepted names / spellings (normalised on load)
COUNTRY_SYNONYMS: dict[str, list[str]] = {
    "bel": ["belgium", "belgie", "belgique", "belgien", "belgië"],
    "ger": ["germany", "deutschland", "duitsland", "allemagne", "deu"],
    "hol": ["holland", "netherlands", "the netherlands", "nederland", "nld", "nl"],
    "lux": ["luxembourg", "luxemburg"],
}

# Hebrew term -> English canonical term (Hebrew side is normalised on load: no niqqud, final letters folded)
HEBREW_TERMS: dict[str, str] = {
    # countries
    "בלגיה": "belgium",
    "גרמניה": "germany",
    "הולנד": "holland",
    "לוקסמבורג": "luxembourg",
    # months
    "ינואר": "january",
    "פברואר": "february",
    "מרץ": "march",
    "מרס": "march",
    "אפריל": "april",
    "מאי": "may",
    "יוני": "june",
    "יולי": "july",
    "אוגוסט": "august",
    "ספטמבר": "september",
    "אוקטובר": "october",
    "נובמבר": "november",
    "דצמבר": "december",
}

# one-letter Hebrew prefixes that may be glued to a term ("in Germany" = בגרמניה, "from Holland" = מהולנד)
HEBREW_PREFIXES = "בלמהוכש"
