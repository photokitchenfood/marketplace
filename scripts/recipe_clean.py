"""Keeps Recipe / Caption / Procedure content in its own section.

Two bugs this prevents (both seen across the 2021-2025 decks):
  1. The Caption text box carries the whole recipe (Serving Time / Yield / Ingredients /
     Procedure ...) between the real caption and the footer, duplicating the Recipe fields.
  2. Heading rows ("Ingredients", "Procedure") end up as the first item of the Recipe lists.

clean_item(item) fixes both in place and returns a list of what changed. Used by
extract_recipe_deck.py and ingest_photo_folder.py on every new card.
"""
import re
import unicodedata

HEAD_P = re.compile(r"^\s*(procedures?|directions?|method|instructions?|preparation)\s*:?\s*$", re.I)
HEAD_I = re.compile(r"^\s*(ingredients?)\s*:?\s*$", re.I)
META = re.compile(r"^\s*(serving time|yield|servings?)\s*:", re.I)


def _nz(s):
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c)).lower().replace("’", "'")
    s = re.sub(r"^\s*(?:[-•*▪●◦·]+|\d+\s*[.)])\s*", "", s)
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def _strip_pref(n):
    return re.sub(r"^(optional|tips?|notes?|serving suggestion)\s+", "", n)


def strip_heading(text, pattern):
    """Drop leading heading line(s) matching pattern (blank lines skipped)."""
    lines = (text or "").split("\n")
    while lines and (not lines[0].strip() or pattern.match(lines[0])):
        lines.pop(0)
    return "\n".join(lines)


def remove_recipe_block(caption, recipe):
    """Caption with the duplicated recipe block removed (intro + footer kept), or None if no duplication."""
    if not recipe or not caption:
        return None
    cl = caption.split("\n")
    pool = set()
    for fld in ("ingredients", "procedure", "tips"):
        for l in (recipe.get(fld) or "").split("\n"):
            n = _nz(l)
            if n:
                pool.add(n)
                pool.add(_strip_pref(n))
    dup, heads = [], 0
    for i, l in enumerate(cl):
        n = _nz(l)
        if not n:
            continue
        if HEAD_I.match(l) or HEAD_P.match(l) or META.match(l):
            dup.append(i)
            heads += 1
        elif n in pool or _strip_pref(n) in pool:
            dup.append(i)
    if not dup or (heads == 0 and len(dup) < 3) or len(dup) < 2:
        return None
    a, b = min(dup), max(dup)
    before = "\n".join(cl[:a]).strip()
    after = "\n".join(cl[b + 1:]).strip()
    return "\n\n".join(x for x in (before, after) if x)


def clean_item(item):
    changes = []
    rec = item.get("recipe")
    if rec:
        new = remove_recipe_block(item.get("caption") or "", rec)
        if new is not None:
            item["caption"] = new
            changes.append("caption")
        for fld, pat in (("ingredients", HEAD_I), ("procedure", HEAD_P)):
            old = rec.get(fld) or ""
            cur = strip_heading(old, pat)
            if cur != old:
                rec[fld] = cur
                changes.append(fld)
    return changes
