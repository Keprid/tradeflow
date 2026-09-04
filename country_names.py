#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
country_names.py
================

Shared module for professional display names.

ITC Trade Map downloads carry verbose official names - countries ("United
States of America", "China, Hong Kong SAR", ...) and products ("Plants and
parts of plants, incl. seeds and fruits, of a kind used primarily in
pharmacy...", ...). Reports and charts read better with the concise forms
used in business publications ("USA", "Hong Kong, China", "Medicinal
plants, n.e.s").

Two resolvers are provided:

* ``display_name()``      - countries / markets, via ``SHORT_NAMES``.
* ``short_product_name()``- product labels, via ``PRODUCT_SHORT_NAMES``
                            (keyed by HS code) with a clause-preserving
                            compressor as fallback.

Both fall back cleanly to the original name - tidied up (whitespace
collapsed, trailing punctuation removed). When a maximum length forces
truncation, whole clauses are kept wherever possible and elided content is
marked with ``,...`` so the reader can tell the description continues.
"""

import re

# Verbose / official name (lowercase) -> professional short display name.
SHORT_NAMES = {
    # A
    "bolivia, plurinational state of": "Bolivia",
    "bonaire, sint eustatius and saba": "Bonaire",
    "brunei darussalam": "Brunei",
    # C
    "cape verde": "Cabo Verde",
    "china, hong kong sar": "Hong Kong, China",
    "hong kong sar, china": "Hong Kong, China",
    "hong kong, china": "Hong Kong, China",
    "china, macao sar": "Macao, China",
    "macao sar, china": "Macao, China",
    "macao, china": "Macao, China",
    "chinese taipei": "Taiwan",
    "taiwan, province of china": "Taiwan",
    "taiwan, chinese taipei": "Taiwan",
    "congo, democratic republic of the": "DR Congo",
    "democratic republic of the congo": "DR Congo",
    "czech republic": "Czechia",
    "cote d'ivoire": "Côte d'Ivoire",
    "côte d'ivoire": "Côte d'Ivoire",
    "ivory coast": "Côte d'Ivoire",
    # E
    "east timor": "Timor-Leste",
    "swaziland": "Eswatini",
    # H
    "holy see (vatican city state)": "Holy See",
    # I
    "iran, islamic republic of": "Iran",
    # K
    "korea, republic of": "South Korea",
    "republic of korea": "South Korea",
    "korea, dem. people's rep. of": "North Korea",
    "korea, democratic people's republic of": "North Korea",
    "democratic people's republic of korea": "North Korea",
    "kyrgyzstan": "Kyrgyzstan",
    # L
    "lao people's democratic republic": "Laos",
    "lao pdr": "Laos",
    "libya, state of": "Libya",
    # M
    "macedonia, north": "North Macedonia",
    "macedonia, the former yugoslav republic of": "North Macedonia",
    "micronesia, federated states of": "Micronesia",
    "moldova, republic of": "Moldova",
    "myanmar (burma)": "Myanmar",
    # P
    "palestine, state of": "Palestine",
    "palestinian territory, occupied": "Palestine",
    # R
    "russian federation": "Russia",
    # S
    "saint martin (french part)": "St. Martin (Fr.)",
    "saint kitts and nevis": "St. Kitts and Nevis",
    "saint vincent and the grenadines": "St. Vincent and Gren.",
    "sint maarten (dutch part)": "St. Maarten",
    "sao tome and principe": "São Tomé and Príncipe",
    "syrian arab republic": "Syria",
    # T
    "tanzania, united republic of": "Tanzania",
    "timor-leste": "Timor-Leste",
    "turks and caicos islands": "Turks & Caicos Is.",
    # U
    "united arab emirates": "UAE",
    "united kingdom": "UK",
    "united kingdom of great britain and northern ireland": "UK",
    "united states": "USA",
    "united states of america": "USA",
    # V
    "venezuela, bolivarian republic of": "Venezuela",
    "viet nam": "Vietnam",
    # ITC aggregate rows
    "area nes": "Unspecified areas",
    "areas nes": "Unspecified areas",
    "european union nes": "European Union",
    "europe othr. nes": "Other Europe",
}


# Lowercase African country names as used by ITC Trade Map.  Used to compute
# "share of exports destined to Africa" and to isolate Kenya within a market's
# export-destination table.  Matches are attempted against the raw ITC label
# AND its short form, so variants ("Congo, Democratic Republic of the",
# "Tanzania, United Republic of") are all caught.
AFRICA_COUNTRIES = {
    "algeria", "angola", "benin", "botswana", "burkina faso",
    "burundi", "cabo verde", "cape verde", "cameroon", "central african republic",
    "chad", "comoros", "congo", "congo, democratic republic of the",
    "democratic republic of the congo", "congo, republic of", "republic of the congo",
    "cote d'ivoire", "côte d'ivoire", "ivory coast", "djibouti", "egypt",
    "equatorial guinea", "eritrea", "eswatini", "swaziland", "ethiopia", "gabon",
    "gambia", "ghana", "guinea", "guinea-bissau", "kenya", "lesotho", "liberia",
    "libya", "libya, state of", "madagascar", "malawi", "mali", "mauritania",
    "mauritius", "morocco", "mozambique", "namibia", "niger", "nigeria",
    "rwanda", "sao tome and principe", "senegal", "seychelles", "sierra leone",
    "somalia", "south africa", "south sudan", "sudan", "tanzania",
    "tanzania, united republic of", "togo", "tunisia", "uganda", "zambia",
    "zimbabwe", "saint helena",
}


def is_africa(name):
    """Return True if ``name`` (raw ITC label or short form) is an African
    country.  Handles the verbose ITC forms used for DR Congo and Tanzania."""
    if name is None:
        return False
    text = " ".join(str(name).lower().split())
    if text in AFRICA_COUNTRIES:
        return True
    return display_name(name).lower() in AFRICA_COUNTRIES


def display_name(name, maxlen=None):
    """Professional display name for a country / market label.

    Exact-match lookup against SHORT_NAMES (case-insensitive); anything else
    falls back cleanly to the original name tidied up (whitespace collapsed,
    trailing ellipsis/punctuation stripped). When ``maxlen`` is given and an
    unmapped name is longer, it is cut at a word boundary and marked with an
    ellipsis so it still reads as truncated rather than broken.
    """
    if name is None:
        return ""
    text = re.sub(r"\s+", " ", str(name)).strip().rstrip(" ,;:")
    mapped = SHORT_NAMES.get(text.lower())
    if mapped:
        return mapped
    if maxlen and len(text) > maxlen:
        head = text[: maxlen - 1]
        cut = head.rsplit(" ", 1)[0].rstrip(" ,;:")
        text = (cut or head).rstrip(" ,;:") + "..."
    return text


def title_partner(name, max_len=14):
    """Uppercase partner name for report titles ('KENYA-<PARTNER> ...').

    Prefers the professional short form (SHORT_NAMES) whenever the official
    name is long, e.g. 'United States of America' -> 'KENYA-USA TRADE FLOW'
    and 'United Arab Emirates' -> 'KENYA-UAE TRADE FLOW'. Short names pass
    through unchanged ('Saudi Arabia' -> 'SAUDI ARABIA'); a leading 'the'
    is dropped ('the World' -> 'WORLD').
    """
    text = re.sub(r"\s+", " ", str(name or "")).strip()
    disp = display_name(text)
    if disp and len(disp) < len(text) and len(text) > max_len:
        text = disp
    text = re.sub(r"^the\s+", "", text, flags=re.IGNORECASE)
    return text.upper()


def narrative_ref(name):
    """Partner reference for running prose, with the article abbreviations take.

    All-caps short forms read naturally with 'the' in sentences such as
    "Kenya's total imports from the USA in 2025"; regular country names do
    not take an article ('imports from Saudi Arabia').
    """
    disp = display_name(name)
    return f"the {disp}" if disp and disp.isupper() else disp


# ---------------------------------------------------------------------------
# Product labels (HS descriptions)
# ---------------------------------------------------------------------------
# Professional short names keyed by HS code (heading level, e.g. "0901", or
# chapter level, e.g. "09"). Codes are stable across HS revisions while the
# raw descriptions drift, so the code is the reliable lookup key.
PRODUCT_SHORT_NAMES = {
    # 01-05  animal & dairy
    "03": "Fish, crustaceans & molluscs",
    "04": "Dairy, eggs & honey",
    # 06-14  plants / food
    "06": "Live trees, plants & cut flowers",
    "0602": "Live plants, bulbs & roots",
    "0603": "Cut roses & buds, fresh",
    "07": "Edible vegetables",
    "0702": "Tomatoes, fresh or chilled",
    "0703": "Onions, shallots & leeks",
    "0710": "Frozen vegetables",
    "0712": "Dried vegetables",
    "0713": "Dried legumes (pulses)",
    "0714": "Cassava, sweet potatoes & yams",
    "08": "Edible fruit & nuts",
    "0801": "Coconuts, Brazil nuts & cashews",
    "0802": "Other nuts, fresh or dried",
    "0803": "Bananas, fresh or dried",
    "0804": "Dates, figs, mangoes & avocados",
    "0805": "Citrus fruit",
    "0806": "Grapes, fresh or dried",
    "0807": "Melons & papayas",
    "0808": "Apples, pears & quinces",
    "0809": "Apricots, cherries & peaches",
    "0810": "Other fresh fruit (berries etc.)",
    "0811": "Frozen fruit & nuts",
    "0812": "Provisionally preserved fruit & nuts",
    "0813": "Dried fruit & mixes",
    "09": "Coffee, tea & spices",
    "0901": "Coffee",
    "0902": "Tea",
    "0910": "Ginger & other spices",
    "10": "Cereals",
    "1001": "Wheat & meslin",
    "1005": "Maize (corn)",
    "1006": "Rice",
    "11": "Milling products, malt & starch",
    "1101": "Wheat flour",
    "12": "Oil seeds, grains & medicinal plants",
    "1201": "Soya beans",
    "1202": "Groundnuts (peanuts)",
    "1207": "Other oil seeds",
    "1211": "Medicinal plants, n.e.s",
    "13": "Lac, gums & resins",
    "14": "Vegetable plaiting materials",
    # 15-24  fats, food preparations, beverages
    "15": "Animal & vegetable oils and fats",
    "1511": "Palm oil",
    "16": "Preparations of meat & fish",
    "17": "Sugars & confectionery",
    "18": "Cocoa & cocoa preparations",
    "1801": "Cocoa beans",
    "19": "Cereals, flour & pastry preparations",
    "20": "Preparations of vegetables & fruit",
    "2008": "Prepared fruits & nuts",
    "2009": "Fruit & vegetable juices",
    "21": "Miscellaneous edible preparations",
    "2101": "Coffee & tea extracts",
    "22": "Beverages, spirits & vinegar",
    "2201": "Water (incl. mineral water)",
    "2202": "Sweetened & flavoured waters",
    "2204": "Wine of fresh grapes",
    "23": "Food industry residues & animal feed",
    "24": "Tobacco",
    # 25-27  minerals / fuels
    "25": "Salt, cement & stone",
    "2501": "Salt",
    "2523": "Portland cement",
    "26": "Metal ores & concentrates",
    "2601": "Iron ores & concentrates",
    "27": "Mineral fuels & oils",
    "2701": "Coal",
    "2709": "Crude petroleum",
    "2710": "Petroleum oils (refined)",
    "2711": "Petroleum gases",
    "2716": "Electrical energy",
    # 28-40  chemicals / plastics
    "28": "Inorganic chemicals",
    "29": "Organic chemicals",
    "30": "Pharmaceuticals",
    "3004": "Medicaments, dosed",
    "31": "Fertilizers",
    "3105": "Mixed fertilizers",
    "32": "Dyes, paints & inks",
    "33": "Essential oils & perfumery",
    "3301": "Essential oils",
    "3304": "Beauty, make-up & skin care preparations",
    "34": "Soaps & washing preparations",
    "3401": "Soap",
    "35": "Glues & enzymes",
    "38": "Miscellaneous chemical products",
    "39": "Plastics & articles thereof",
    "3901": "Polyethylene",
    "3902": "Polypropylene",
    "40": "Rubber & articles thereof",
    "4011": "New pneumatic tyres",
    # 41-43  hides / leather
    "41": "Raw hides & leather",
    "42": "Leather articles & travel goods",
    # 44-49  wood / paper
    "44": "Wood & articles of wood",
    "4407": "Sawn wood",
    "47": "Wood pulp",
    "48": "Paper & paperboard",
    "49": "Printed books & newspapers",
    # 50-63  textiles / apparel
    "52": "Cotton",
    "5201": "Raw cotton",
    "53": "Vegetable textile fibres",
    "54": "Man-made filaments",
    "55": "Man-made staple fibres",
    "58": "Special woven fabrics",
    "60": "Knitted fabrics",
    "61": "Apparel, knitted",
    "62": "Apparel, not knitted",
    "63": "Home textiles & worn clothing",
    # 64-67  footwear etc.
    "64": "Footwear",
    "65": "Headgear",
    # 68-71  stone / glass / precious
    "68": "Articles of stone & cement",
    "69": "Ceramic products",
    "70": "Glass & glassware",
    "71": "Precious stones & metals",
    # 72-83  base metals
    "72": "Iron & steel",
    "73": "Articles of iron & steel",
    "74": "Copper",
    "75": "Nickel",
    "76": "Aluminium",
    "78": "Lead",
    "79": "Zinc",
    "81": "Other base metals",
    "82": "Tools & cutlery",
    "83": "Miscellaneous base-metal articles",
    # 84-85  machinery / electronics
    "84": "Machinery & mechanical appliances",
    "8413": "Pumps & liquid elevators",
    "8432": "Agricultural machinery",
    "8471": "Computers & data-processing machines",
    "8501": "Electric motors & generators",
    "85": "Electrical machinery & electronics",
    "8517": "Phones & telecom equipment",
    "8528": "TV & monitor receivers",
    "8544": "Insulated wire & cable",
    # 86-89  transport
    "87": "Vehicles",
    "8703": "Motor cars",
    "8704": "Trucks & goods vehicles",
    "8708": "Vehicle parts & accessories",
    "8711": "Motorcycles",
    "88": "Aircraft & spacecraft",
    "89": "Ships & boats",
    # 90-97  instruments / misc
    "90": "Optical & medical instruments",
    "91": "Clocks & watches",
    "92": "Musical instruments",
    "93": "Arms & ammunition",
    "94": "Furniture & bedding",
    "9401": "Seats (chairs etc.)",
    "95": "Toys, games & sports equipment",
    "96": "Miscellaneous manufactures",
    "97": "Works of art & antiques",
}

_CLAUSE_SPLIT = re.compile(r"[;,]\s*")
_CODE_STRIP = re.compile(r"[^0-9]")

# Shortened labels must keep at least this many whole words (when the full
# label has them) so a stub like "Animals,..." never hides the meaning.
MIN_SHORT_WORDS = 4

# Do not leave a shortened label dangling on one of these connectors.
_TRAILING_STOPWORDS = {"and", "&", "of", "the", "for", "with", "in", "to",
                       "or", "other", "from"}


def _tidy_label(label):
    """Collapse whitespace and drop trailing ellipsis / punctuation."""
    return re.sub(r"\s+", " ", str(label)).strip().rstrip(" ,;:")


def _truncate_words(text, maxlen):
    """Keep whole words of ``text`` while they fit within ``maxlen`` chars.

    Always keeps at least ``MIN_SHORT_WORDS`` of them (unless the label is
    shorter), never ends on a dangling connector word, and appends no '...'
    marker: endings are complete words.
    """
    if not text:
        return ""
    text = re.sub(r"(?<=\w)\s+and\s+(?=\w)", " & ", text)
    if len(text) <= maxlen:
        return text
    words = text.split()
    out = words[0]
    for word in words[1:]:
        candidate = f"{out} {word}"
        if len(candidate) > maxlen and out.count(" ") + 1 >= MIN_SHORT_WORDS:
            break
        out = candidate
    # A label must not end on a dangling connector ("... fresh or").
    while " " in out and out.rsplit(" ", 1)[1].strip(",").lower() \
            in _TRAILING_STOPWORDS:
        out = out.rsplit(" ", 1)[0].rstrip(" ,;:")
    return out


def short_product_name(label=None, code=None, maxlen=None):
    """Professional short name for a product label.

    Resolution order:

    1. The informative first clause of the real label (everything up to the
       first semicolon) is preferred.  This keeps the useful detail -- e.g.
       "Coffee, whether or not roasted or decaffeinated" -- instead of
       collapsing to a coarse heading like "Coffee".  A concise clause is
       shown in full; only very long clauses are word-truncated.
    2. Exact HS-code lookup in PRODUCT_SHORT_NAMES, used only as a fallback
       when the label has no informative clause.
    3. Otherwise the label is tidied and word-truncated if needed.
    """
    if label:
        text = _tidy_label(label)
        if text and len(re.split(r";", text, maxsplit=1)[0].split()) \
                >= MIN_SHORT_WORDS:
            clause = _tidy_label(re.split(r";", text, maxsplit=1)[0])
            # keep a concise clause readable: truncate only beyond ~50 chars
            effective = max(maxlen or 50, 50)
            return _truncate_words(clause, effective)

    if code is not None:
        mapped = PRODUCT_SHORT_NAMES.get(_CODE_STRIP.sub("", str(code)))
        if mapped:
            return mapped

    text = _tidy_label(label)
    return _truncate_words(text, maxlen or 50) if text else ""


# Words conventionally left lowercase in Title Case headings/captions.
_TITLE_MINOR = {"a", "an", "the", "and", "but", "or", "nor",
                "for", "of", "on", "in", "to", "by", "with",
                "from", "at", "as", "via", "per"}


def title_case(text):
    """Convert a caption/title string to Title Case.

    Respects a leading "Table N:" / "Figure N:" prefix, keeps acronyms and
    words with existing internal capitals (USD, UNCTAD, EAC, Kenya's) intact,
    and leaves minor words (prepositions/conjunctions/articles) lowercase
    unless they open the title.
    """
    text = str(text)

    def cap_word(w, force_upper):
        if w.isupper():
            return w
        if any(ch.isupper() for ch in w[1:]):
            return w
        if not force_upper and w.lower() in _TITLE_MINOR:
            return w.lower()
        return w[:1].upper() + w[1:]

    m = re.match(r"^((?:Table|Figure)\s+\d+[a-z]?\s*[:\-–]\s*)(.*)$", text, re.IGNORECASE)
    prefix, rest = m.groups() if m else ("", text)
    words = re.split(r"(\s+)", rest)
    out = []
    force = True
    for w in words:
        if w.isspace():
            out.append(w)
            continue
        out.append(cap_word(w, force))
        force = False
    return prefix + "".join(out)
