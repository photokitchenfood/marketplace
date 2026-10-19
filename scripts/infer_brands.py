#!/usr/bin/env python3
"""
Infer the `brand` field for recipes/*.js items whose source deck has no bolded
brand text (the 2022 decks set brands in sentence case instead of bold).

Brands are read from the item's own text, not the deck:

  1. Known brands. Every brand string already in the catalog (other years) is
     matched, as a whole phrase, in the caption, ingredients and procedure.
     Longest match wins; a match inside a longer match is dropped.
  2. Unknown brands. In the ingredients list, a line that starts with a quantity
     and then a Capitalised phrase ("1 pack UnMeat Chick'N Patties",
     "6 Tender Juicy Jumbo Hotdogs, sliced") names a branded product. The
     phrase is trimmed to a known brand prefix when one exists, otherwise to
     its first two words.
  3. Brand-only layouts (OUT OF PACK / NON-FOOD / GROUP SHOT items titled with the
     brand itself, e.g. "Skippy", "Sensodyne"): the layout name, matched or
     trimmed the same way against the caption.

Only items whose `brand` is empty are touched, and only the `brand` key changes.

Usage:
    python3 scripts/infer_brands.py --evaluate 2023 2024     # score against real brands
    python3 scripts/infer_brands.py --apply 22               # fill empty brands in recipes/22*.js
    python3 scripts/infer_brands.py --apply 22 --dry-run     # show what would change
"""

import argparse
import collections
import json
import re
import sys
from pathlib import Path

RECIPES = Path(__file__).resolve().parent.parent / "recipes"
NO_BRAND_TAG = "! Generic"

# Capitalised words that start an ingredient phrase but are not brands.
_GENERIC = None


def generic_words():
    """Lowercase food/ingredient words learned from the catalog's own ingredient text:
    a word that shows up in lowercase far more than Capitalised is not a brand."""
    global _GENERIC
    if _GENERIC is None:
        low, cap = collections.Counter(), collections.Counter()
        for f in RECIPES.glob("*.js"):
            s = f.read_text(encoding="utf-8")
            for m in re.finditer(r'"ingredients": "((?:[^"\\]|\\.)*)"', s):
                for line in re.split(r"\\n|\\u000b", m.group(1)):
                    toks = re.findall(r"[A-Za-zé’'\-]+", line)
                    for i, tk in enumerate(toks):
                        (cap if tk[:1].isupper() else low)[tk.lower()] += 1
        _GENERIC = {w for w, c in low.items() if c >= 4 and c >= 3 * cap.get(w, 0)}
    return _GENERIC


NOT_BRANDS = {
    "salt", "pepper", "oil", "water", "sugar", "butter", "egg", "eggs", "milk", "cheese", "rice", "flour",
    "chicken", "beef", "pork", "fish", "shrimp", "salmon", "tuna", "garlic", "onion", "tomato", "tomatoes",
    "lemon", "lime", "orange", "apple", "banana", "mango", "strawberry", "strawberries", "chocolate", "cream",
    "korean", "japanese", "italian", "chinese", "thai", "mexican", "greek", "french", "american", "filipino",
    "spanish", "vietnamese", "indian", "australian", "canadian", "german", "singaporean", "hawaiian",
    "parmesan", "parmigiano", "mozzarella", "cheddar", "gouda", "brie", "feta", "ricotta", "gruyere", "dijon",
    "worcestershire", "tabasco", "sriracha", "bbq", "worcester", "extra", "fresh", "dried", "ground", "whole",
    "the", "a", "an", "and", "or", "for", "to", "of", "with", "in", "on", "optional", "topping", "garnish",
    "sauce", "mix", "add", "serve", "serving", "yield", "tips", "tags", "ingredients", "procedure", "caption",
    "i", "if", "it", "is", "this", "that", "these", "those", "our", "your", "you", "we", "make", "try",
    "order", "join", "head", "find", "get", "love", "enjoy", "available", "perfect", "great", "good",
    "monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday", "mom", "dad", "christmas",
    "lamb", "bangus", "tofu", "noodles", "pasta", "bread", "ham", "bacon", "sausage", "hotdog", "hotdogs", "patties",
    "indent", "grated", "marinated", "boneless", "ready-to-eat", "locally",
    "halloween", "easter", "valentine", "valentines", "mother", "mothers", "father", "fathers",
}
MODIFIERS = {
    "medium", "medium-sized", "large", "small", "jumbo", "mini", "extra", "mild", "sweet", "spicy", "plain", "unsalted",
    "salted", "whole", "pure", "organic", "natural", "pitted", "sliced", "diced", "chopped", "fresh", "frozen", "canned",
    "dried", "instant", "ready", "cooked", "raw", "lean", "low", "fat", "free", "light", "dark", "white", "black", "red",
    "green", "yellow", "brown", "golden", "toasted", "smoked", "boneless", "skinless", "thick", "thin", "long", "short",
    "grated", "shredded", "crushed", "minced", "ripe", "firm", "soft", "hard", "unsweetened", "sweetened", "all-purpose",
    "full", "half", "double", "single", "baby", "wild", "sea", "kosher", "roasted", "fried", "baked", "steamed", "premium",
}
UNITS = (r"cups?|tbsps?\.?|tsps?\.?|tablespoons?|teaspoons?|g|kg|mg|ml|l|oz|lbs?\.?|pcs?\.?|pieces?|packs?|packets?|"
         r"cans?|bars?|bottles?|sheets?|slices?|strips?|cloves?|boxes?|jars?|tubs?|blocks?|sticks?|sachets?|"
         r"bunch(?:es)?|heads?|stalks?|scoops?|dashes?|pinch(?:es)?|sprigs?|rolls?|loaf|loaves|bags?|tetra|pints?|"
         r"pouch(?:es)?|cartons?|links?|fillets?|steaks?|cubes?|drops?|handfuls?")
QTY = re.compile(rf"^\s*(?:[\d¼½¾⅓⅔⅛⅜⅝⅞/.\-–]+\s*)+(?:\([^)]*\)\s*)?(?:(?:{UNITS})\b\.?\s*)*(?:\([^)]*\)\s*)?", re.I)
CAP_TOKEN = re.compile(r"^[A-Z0-9][A-Za-z0-9’'&.\-éñ!]*$")
CONNECTORS = {"&", "and", "of", "de", "la", "le", "n’", "n'", "y", "the"}


def norm(text):
    t = re.sub(r"[^a-z0-9]+", " ", text.lower().replace("’", "'").replace("é", "e"))
    return re.sub(r"\band\b", " ", t).replace("  ", " ").strip() if " and " in f" {t} " else t.strip()


def known_brands(exclude_prefixes=()):
    """Counter of brand strings in the catalog (sentence-case dictionary source)."""
    found = collections.Counter()
    for f in RECIPES.glob("*.js"):
        if f.name.startswith(tuple(exclude_prefixes)):
            continue
        s = f.read_text(encoding="utf-8")
        for m in re.finditer(r'"brand": (\[[^\]]*\])', s):
            try:
                for b in json.loads(m.group(1)):
                    b = b.strip()
                    if not b or b == NO_BRAND_TAG or re.match(r"^\d", b) or len(b.split()) > 4:
                        continue
                    found[b.replace("'", "’")] += 1
            except ValueError:
                pass
    # drop entries that are really generic words / ingredients
    return collections.Counter({b: c for b, c in found.items()
                                if norm(b) not in NOT_BRANDS and len(norm(b)) >= 3})


def find_known(text, known_norm, known=None, cased=True):
    """Whole-phrase matches of known brands in text; longest first, no overlaps.
    With cased=True a match must start with a capital letter in the original text
    ("Ensure Gold" counts, "to ensure that" doesn't), and a one-word brand must
    appear in the catalog at least twice (single-use one-worders are mostly noise)."""
    words = re.findall(r"[A-Za-z0-9é’'&.!\-ñ]+", text)
    nw = [norm(w) for w in words]
    hits = []
    for key, display in known_norm:
        parts = key.split()
        if not parts:
            continue
        if cased and len(parts) == 1 and known is not None and known.get(display, 0) < 2:
            continue
        n = len(parts)
        for i in range(len(words) - n + 1):
            seg = [x for x in nw[i:i + n]]
            if " ".join(x for x in seg if x) == key or norm(" ".join(words[i:i + n])) == key:
                if cased and not words[i][:1].isupper() and not words[i][:1].isdigit():
                    continue
                hits.append((i, i + n, display))
    hits.sort(key=lambda h: (-(h[1] - h[0]), h[0]))
    taken, out = [], []
    for a, b, d in hits:
        if any(not (b <= x or a >= y) for x, y in taken):
            continue
        taken.append((a, b)); out.append((a, d))
    return [d for a, d in sorted(out)]


def cap_phrase(rest):
    """Leading run of Capitalised tokens (connectors allowed inside), cut at comma/lowercase."""
    rest = re.split(r",|\(|;|\bto taste\b", rest)[0].strip()
    toks, out = rest.split(), []
    for i, tk in enumerate(toks):
        if CAP_TOKEN.match(tk) or (out and tk.lower() in CONNECTORS and i + 1 < len(toks) and CAP_TOKEN.match(toks[i + 1])):
            out.append(tk)
        else:
            break
    while out and out[-1].lower() in CONNECTORS:
        out.pop()
    return out


def trim_to_known(tokens, known_norm_set):
    """Longest known-brand prefix of the phrase, else first two tokens."""
    for n in range(min(len(tokens), 4), 0, -1):
        if norm(" ".join(tokens[:n])) in known_norm_set:
            return " ".join(tokens[:n])
    gen = generic_words()
    out = [tokens[0]]
    for tk in tokens[1:2]:                # unknown brand: at most two words, stopping at a food/modifier word
        if tk.lower() in gen or tk.lower() in MODIFIERS or tk.lower() in NOT_BRANDS:
            break
        out.append(tk)
    return " ".join(out)


def ingredient_candidates(ingredients, known_norm_set):
    out = []
    for line in re.split(r"[\n\x0b]+", ingredients or ""):
        line = line.strip().lstrip("•-– ")
        if not re.match(r"^[\d¼½¾⅓⅔⅛⅜⅝⅞]", line):
            continue                      # branded products carry a quantity; bare "Salt and pepper" lines don't
        rest = QTY.sub("", line, count=1).strip()
        # a number that belongs to the brand ("19 Crimes") is kept when a Capitalised word follows it
        m = re.match(r"^(\d{1,3})\s+([A-Z][A-Za-z’']+)", rest)
        toks = cap_phrase(rest)
        first = toks[0].lower() if toks else ""
        if not toks or first in NOT_BRANDS or first in MODIFIERS or (first in generic_words() and not (m and toks[0].isdigit())):
            continue
        while len(toks) > 1 and toks[0].lower() in MODIFIERS:
            toks = toks[1:]
        out.append(trim_to_known(toks, known_norm_set))
    return out


def tidy(brands, known):
    """Trim product-line suffixes to the most common known brand prefix, title-case ALL CAPS,
    and drop a brand that contains another selected brand as a whole-word prefix or suffix."""
    out = []
    for b in brands:
        if b.isupper() and len(b) > 4:
            b = b.title()
        words = b.split()
        cands = [" ".join(words[:n]) for n in range(1, len(words))]
        cands = [c for c in cands if known.get(c, 0) >= 3]
        if cands and known.get(b, 0) < max(known[c] for c in cands):
            b = max(cands, key=lambda c: (known[c], -len(c)))
        out.append(b)
    final = []
    for b in out:
        nb = norm(b)
        if any(o is not b and nb != norm(o) and (nb.startswith(norm(o) + " ") or nb.endswith(" " + norm(o))) for o in out):
            continue
        if nb not in [norm(x) for x in final]:
            final.append(b)
    return final


def infer(item, known, known_norm, known_norm_set):
    rec = item.get("recipe") or {}
    caption = item.get("caption", "")
    ingredients = rec.get("ingredients", "")
    procedure = rec.get("procedure", "")
    brands, seen = [], set()

    def add(b):
        b = b.strip().replace("'", "’")
        if not b or norm(b) in NOT_BRANDS or norm(b) in seen or len(norm(b)) < 3:
            return
        seen.add(norm(b)); brands.append(b)

    for b in find_known(" . ".join([caption, ingredients, procedure]), known_norm, known):
        add(b)
    for b in ingredient_candidates(ingredients, known_norm_set):
        # skip when an already-found known brand covers it
        if not any(norm(b) in norm(x) or norm(x) in norm(b) for x in brands):
            add(b)
    if not brands and item.get("category") in ("OUT OF PACK", "NON-FOOD", "GROUP SHOT"):
        name = re.sub(r"\s*\(group shot\)\s*$", "", item.get("layoutName", ""), flags=re.I).strip()
        name = re.sub(r"\s+[A-D]$", "", name)                 # "Sensodyne A" / "Sensodyne B" photo variants
        toks = name.split()
        if name and toks and norm(name) not in NOT_BRANDS and toks[0].lower() not in NOT_BRANDS and 1 <= len(toks) <= 9:
            hit = find_known(name, known_norm, known, cased=False)
            if hit:
                add(hit[0])
            elif len(toks) <= 3 and all(CAP_TOKEN.match(t) or t.lower() in CONNECTORS for t in toks):
                add(name)                                       # the layout is titled with the brand itself
            else:
                first = cap_phrase(name)
                if first and first[0].lower() not in NOT_BRANDS and first[0].lower() not in generic_words():
                    add(trim_to_known(first, known_norm_set))
    return tidy(brands, known)


def prep(known):
    kn = sorted(((norm(b), b) for b in known), key=lambda kv: -len(kv[0]))
    return kn, {k for k, _ in kn}


def load_items(path):
    s = Path(path).read_text(encoding="utf-8")
    return s, json.loads(s[s.index("= [") + 2:s.index("\n];") + 2])


def match(pred, truth):
    """A predicted brand matches a real one if equal, or one is a whole-word prefix of the other."""
    a, b = norm(pred), norm(truth)
    return a == b or a.startswith(b + " ") or b.startswith(a + " ")


def evaluate(years):
    for y in years:
        yy = y[-2:]
        known = known_brands(exclude_prefixes=(yy, "22"))
        kn, kns = prep(known)
        tp = fp = fn = items = exact = 0
        misses, extras = [], []
        for f in sorted(RECIPES.glob(f"{yy}*.js")):
            _, arr = load_items(f)
            for it in arr:
                truth = [b for b in it.get("brand", []) if b != NO_BRAND_TAG]
                if not truth:
                    continue
                items += 1
                pred = infer({**it, "brand": []}, known, kn, kns)
                m_t = [t for t in truth if any(match(p, t) for p in pred)]
                m_p = [p for p in pred if any(match(p, t) for t in truth)]
                tp += len(m_t); fn += len(truth) - len(m_t); fp += len(pred) - len(m_p)
                if len(m_t) == len(truth) and len(m_p) == len(pred):
                    exact += 1
                misses += [(f.stem, it["layoutCode"], t) for t in truth if t not in m_t]
                extras += [(f.stem, it["layoutCode"], p) for p in pred if p not in m_p]
        rec = tp / (tp + fn) if tp + fn else 0
        pre = tp / (tp + fp) if tp + fp else 0
        print(f"{y}: {items} items with real brands | recall {rec:.0%} | precision {pre:.0%} | exact item match {exact/items:.0%}")
        print("   missed (sample):", [m[2] for m in misses[:12]])
        print("   extra  (sample):", [e[2] for e in extras[:12]])


def apply(prefix, dry):
    known = known_brands(exclude_prefixes=())
    kn, kns = prep(known)
    changed = 0
    for f in sorted(RECIPES.glob(f"{prefix}*.js")):
        s, arr = load_items(f)
        n = 0
        for it in arr:
            if it.get("brand"):
                continue
            b = infer(it, known, kn, kns)
            if b:
                n += 1
                if dry:
                    print(f"  {f.stem:34} {it['layoutCode']:7} {it['layoutName'][:34]:34} -> {b}")
                else:
                    it["brand"] = b
        if n and not dry:
            # rewrite only the brand arrays: patch the JSON text of each changed item in place
            sys.path.insert(0, str(Path(__file__).resolve().parent))
            import extract_recipe_deck as x
            label = re.search(r"// Shoot: (.*)", s).group(1)
            u = re.search(r'\.dropboxUrl = (".*?");', s)
            prof = re.search(r"// Template profile: (.*)", s)
            for it in arr:
                it["_profile"] = prof.group(1) if prof else "TMP 2023-2026"
            x.write_group(f, label, arr, json.loads(u.group(1)) if u else None)
        changed += n
        print(f"{f.stem}: {n} item(s) filled")
    print("items filled:", changed)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--evaluate", nargs="+", metavar="YEAR")
    ap.add_argument("--apply", metavar="YY", help="file prefix, e.g. 22")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    if a.evaluate:
        evaluate(a.evaluate)
    elif a.apply:
        apply(a.apply, a.dry_run)
    else:
        ap.print_help()
