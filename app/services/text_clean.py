import re

_MOJIBAKE = [
    ("â€œ", "\u201c"),
    ("â€", "\u201d"),
    ("â€™", "\u2019"),
    ("â€“", "\u2013"),
    ("â€”", "\u2014"),
    ("Ã¡", "á"),
    ("Ã©", "é"),
    ("Ã³", "ó"),
    ("Ãº", "ú"),
    ("Ã±", "ñ"),
    ("Ã¼", "ü"),
    ("Ã¨", "è"),
    ("Ã¯", "ï"),
    ("Ã¶", "ö"),
    ("Ã»", "û"),
    ("Ã¤", "ä"),
    ("Ã´", "ô"),
    ("Ã¢", "â"),
    ("Ãª", "ê"),
    ("Ã§", "ç"),
    ("Ã„", "Ä"),
    ("Ã‰", "É"),
    ("Ãœ", "Ü"),
    ("Ã‘", "Ñ"),
    ("Ã“", "Ó"),
    ("Ãš", "Ú"),
    ("Ã€", "À"),
    ("Ã†", "Æ"),
    ("Ã", "Í"),
    ("Â", ""),
]

_CLEAN_RE = re.compile(r"[^A-Za-z0-9ÁÉÍÓÚÜÑáéíóúüñ .;:,/\-()&'%+#_°º]")

_WORD_FIXES = {
    "canon": "cañón",
    "sandvick": "Sandvik",
    "epiroc": "Epiroc",
    "resemin": "Resemin",
    "caterpillar": "Caterpillar",
    "motoconformadora": "Motoniveladora",
    "retroexcavadora": "Retroexcavadora",
    "camion": "camión",
    "polvorera": "Polvorera",
    "hilux": "Hilux",
}

_CATEGORY_FIXES = {
    "BARRENACION": "BARRENACIÓN",
    "VEHICULOS LIGEROS": "VEHÍCULOS LIGEROS",
    "ACARREO": "ACARREO",
    "ANCLAJE": "ANCLAJE",
    "SCOOP TRAM": "SCOOP TRAM",
    "RETROEXCAVADORAS": "RETROEXCAVADORAS",
    "VEHICULO": "VEHÍCULO",
    "PERFORACION": "PERFORACIÓN",
    "EXPLOSIVOS": "EXPLOSIVOS",
    "MANTENIMIENTO": "MANTENIMIENTO",
    "TALLER": "TALLER",
}

_CONNECTORS = {
    "de", "del", "la", "las", "los", "el", "y", "e", "o", "a", "con",
    "por", "para", "en", "su", "sus", "un", "una", "al", "ni", "lo",
}


def caps(value):
    s = clean_text(value)
    if not s:
        return s
    index = -1
    for i, ch in enumerate(s):
        if ch.isalpha():
            index = i
            break
    if index >= 0 and s[index].islower():
        s = s[:index] + s[index].upper() + s[index + 1:]
    return s


def clean_text(value):
    if value is None:
        return None
    if not isinstance(value, str):
        value = str(value)
    s = value
    for old, new in _MOJIBAKE:
        if old in s:
            s = s.replace(old, new)
    s = _CLEAN_RE.sub(" ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def clean_eco(value):
    s = clean_text(value)
    if s is None:
        return None
    return "".join(s.upper().split())


def clean_category(value):
    s = clean_text(value)
    if s is None:
        return None
    s = s.upper()
    return _CATEGORY_FIXES.get(s, s)


def clean_status(value):
    s = clean_text(value)
    if s is None:
        return None
    return s.upper()


def clean_unit(value):
    s = clean_text(value)
    if s is None:
        return None
    return s.upper()


def clean_supervisor(value):
    s = clean_text(value)
    if s is None:
        return None
    return s.upper()


def title_case(value):
    s = clean_text(value)
    if s is None:
        return None
    words = []
    word_index = 0
    for raw in re.split(r"(\s+)", s):
        if raw.isspace() or raw == "":
            words.append(raw)
            continue
        word = _WORD_FIXES.get(raw.lower(), raw)
        if any(ch.isdigit() for ch in word):
            word = word.upper()
        elif len(word) == 1:
            word = word.upper()
        elif word_index > 0 and word.lower() in _CONNECTORS:
            word = word.lower()
        else:
            word = word[:1].upper() + word[1:].lower()
        words.append(word)
        word_index += 1
    return "".join(words)


clean_name = title_case