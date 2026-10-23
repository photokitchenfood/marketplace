"""Brand canonicalization for the extract / ingest scripts.

Reads ../brand-aliases.json (the same single source of truth index.html uses) and maps
every brand value to its canonical spelling, so new extractions can't reintroduce
duplicate variants. A name that differs from a known one only by case, accents or
apostrophes is mapped automatically; unknown brands pass through unchanged.
"""
import json
import re
import unicodedata
from pathlib import Path

GENERIC_BRAND = "! Generic"
_PATH = Path(__file__).resolve().parent.parent / "brand-aliases.json"
_aliases = _keys = None


def fold(text):
    t = unicodedata.normalize("NFKD", text)
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = re.sub(r"['‘’ʼ`]", "", t).lower()
    return re.sub(r"\s+", " ", t).strip()


def _load():
    global _aliases, _keys
    if _aliases is None:
        d = json.loads(_PATH.read_text(encoding="utf-8"))
        _aliases = dict(d.get("aliases", {}))
        _keys = {fold(c): c for c in d.get("canonical", [])}
        _keys.update({fold(v): c for v, c in _aliases.items()})
    return _aliases, _keys


def canon_brand(name):
    t = re.sub(r"\s+", " ", str(name)).strip()
    if not t or t == GENERIC_BRAND:
        return t
    aliases, keys = _load()
    return aliases.get(t) or keys.get(fold(t)) or t


def canon_brands(names):
    out = []
    for n in names or []:
        c = canon_brand(n)
        if c and c not in out:
            out.append(c)
    return out
