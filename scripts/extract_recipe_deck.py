#!/usr/bin/env python3
"""
Extract Recipe & Captions Deck exports into per-campaign JS data files
for the TMP Recipes Catalog, mirroring the campaigns/ folder pattern used
by the existing PhotoKitchen Shoot Catalog (marketplace repo).

Two source decks are read:

  --recipe-deck   The month's "Recipe & Captions Deck" export. Source of
                   truth for Layout Name, layout code, campaign banner,
                   photo, caption, and recipe text.

  --preprod-deck  The month's "Pre-Prod Deck" export. The Recipe & Captions
                   Deck carries no category label anywhere (confirmed by
                   diagnose_deck.py), so category (RECIPE / OUT OF PACK /
                   OUT OF PACK W/ FOOD STYLING / NON-FOOD / GROUP SHOT) is
                   looked up from the Pre-Prod deck by matching layout code.

All fields are read by position ("drop zone"), not by placeholder name or
alt text, matching the existing PPM-to-Recipe-deck automation approach —
these decks are manually assembled from a Pre-Prod deck template, so the
only stable signal is where a box sits on the slide.

Several deck pairs can be processed in one run (e.g. March + April, when a
campaign like April IG spans both). Every run:

  1. Checks each pair's filenames share a YYMM- prefix (asks if not).
  2. DRY RUN: lists every distinct campaign banner across all decks — its
     text, fill color, slide count, and which deck(s) it appears in — plus
     cross-deck checks. Nothing is written. --dry-run stops here.
  3. Asks which content month (YYMM) each campaign is filed under. There
     is no default: a deck's month is when it was shot, and campaigns in
     it often target other months (Mother's Day shot in April -> May).
  4. Writes one <YYMM>-<Campaign-Name>.js per campaign + month, pulling
     that campaign's items from whichever deck(s) they're in.

Campaign names are normalized for the output `campaign` field and for
grouping: "Highlight" / "Instagram" -> "IG", case and apostrophes ignored.
Slides are correlated between the two decks by layout code, never by
banner text, so the decks' inconsistent labels can't mismatch slides. If a
layout code is missing from the PPM deck (slides left with the template's
"IG00"), the category falls back to matching the slide title; every such
match is listed in the dry run. Hidden slides are skipped in both decks.

Usage:
    python3 scripts/extract_recipe_deck.py \\
        --preprod-deck "pptx/2603-PreProd_TMP Mar 2026 Pre-Prod Deck (...).pptx" \\
        --recipe-deck  "pptx/2603-RecipeCaptions_TMP Mar 2026 Recipes & Captions Deck.pptx" \\
        --preprod-deck "pptx/2604-PreProd_TMP Apr 2026 Pre-Prod Deck (...).pptx" \\
        --recipe-deck  "pptx/2604-RecipeCaptions_TMP Apr 2026 Recipes & Captions Deck (...).pptx" \\
        [--dry-run] [--out-dir recipes]
"""

import argparse
import base64
import io
import json
import re
import sys
from datetime import date
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from PIL import Image

THUMB_MAX_DIM = 500
THUMB_JPEG_QUALITY = 85

# The Recipe & Captions deck's "Download Link/s" slide (slide 2 in every
# deck checked so far) lists one bulleted, hyperlinked line per campaign
# (and sometimes per stray asset). Scanning a few slides past that covers
# decks where it shifts by one or two.
DROPBOX_LINK_SCAN_SLIDES = 10

KNOWN_CATEGORIES = [
    "RECIPE",
    "OUT OF PACK W/ FOOD STYLING",  # check before "OUT OF PACK" (substring)
    "OUT OF PACK",
    "NON-FOOD",
    "GROUP SHOT",
]

# Bolded phrases matching any of these (case-insensitive, matched after
# stripping punctuation/symbols) are dropped from the `brand` field — they
# refer to the shoot company itself, not a product brand, even though
# they're sometimes bolded in captions ("...at The Marketplace! 🛒").
# Add more terms here as needed; no extraction-logic changes required.
EXCLUDED_BRAND_TERMS = [
    "The Marketplace",
    "TMP",
]

# Bolded field labels that show up as the first run of the Caption/Procedure
# text boxes themselves ("Caption:", "Procedure:") — structural, not brands.
BOLD_STRUCTURAL_LABELS = {"caption", "procedure"}
# ...which can also sit in the same bold run as a brand right after it, with
# no non-bold run in between ("Caption:" + "Cetaphil Baby Gentle Wash").
_LEADING_STRUCTURAL_LABEL = re.compile(
    r"^(?:" + "|".join(BOLD_STRUCTURAL_LABELS) + r")\s*:[\s\x0b]*", flags=re.IGNORECASE
)

# Drop zones, in inches from slide top-left, with tolerance. Derived from
# diagnose_deck.py / scan_overview.py runs against the Sep 2026 decks.
# Both decks use the same slide size (10.0 x 5.62 in) and the same
# Pre-Prod-mirrored positions for layout code / campaign banner.
ZONE_LAYOUT_CODE = dict(left=(0.10, 0.65), top=(0.40, 0.90))
ZONE_CAMPAIGN_BANNER = dict(left=(0.95, 1.85), top=(-0.05, 0.15))
ZONE_ITEM_PHOTO = dict(left=(0.10, 0.65), top=(0.90, 1.65))
# Pre-Prod deck only: the on-canvas "selected" category badge. The full
# legend of category options also lives on every Pre-Prod slide but is
# parked off-canvas at left=-1.19in — any positive-left zone naturally
# excludes it.
ZONE_CATEGORY_BADGE = dict(left=(6.30, 7.30), top=(0.40, 0.75))


def in_zone(shape, zone):
    if shape.left is None or shape.top is None:
        return False
    left = Emu(shape.left).inches
    top = Emu(shape.top).inches
    return zone["left"][0] <= left <= zone["left"][1] and zone["top"][0] <= top <= zone["top"][1]


def is_hidden(slide):
    # "Hide slide" in PowerPoint / Google Slides sets show="0" on the slide
    return slide._element.get("show") == "0"


def normalize_title(text):
    return re.sub(r"[^a-z0-9]+", " ", text.lower().replace("’", "'")).strip()


def shape_text(shape):
    if shape.has_text_frame:
        return shape.text_frame.text
    return ""


def placeholder_idx_type(shape):
    if not getattr(shape, "is_placeholder", False):
        return None, None
    try:
        pf = shape.placeholder_format
        return pf.idx, pf.type
    except Exception:
        return None, None


def find_placeholder(slide, idx=None, ptype=None, prefer_longest_text=True):
    """Return the placeholder shape matching idx/ptype. When several
    shapes share the same idx (seen on complex recipes where a filled
    text box was pasted over an empty leftover placeholder), prefer the
    one with the most actual text."""
    candidates = []
    for shape in slide.shapes:
        shape_idx, shape_ptype = placeholder_idx_type(shape)
        if idx is not None and shape_idx != idx:
            continue
        if ptype is not None and shape_ptype != ptype:
            continue
        candidates.append(shape)
    if not candidates:
        return None
    if prefer_longest_text:
        candidates.sort(key=lambda s: len(shape_text(s).strip()), reverse=True)
    return candidates[0]


def find_shape_in_zone(slide, zone, shape_types=None):
    matches = []
    for shape in slide.shapes:
        if shape_types is not None and shape.shape_type not in shape_types:
            continue
        if in_zone(shape, zone):
            matches.append(shape)
    return matches


def strip_label(text, label_pattern):
    """Remove a leading field-label line like 'Caption:' or 'Ingredients:'
    (plus any Google-Slides soft-break \\x0b) from extracted text."""
    text = text.lstrip("\ufeff")
    m = re.match(label_pattern + r"\s*[\x0b\n]*", text, flags=re.IGNORECASE)
    if m:
        text = text[m.end():]
    return text.strip()


def normalize_campaign(text):
    """Canonical campaign name. The PPM and Recipe & Captions decks label
    the same IG content inconsistently ("MARCH HIGHLIGHT" vs "MARCH
    INSTAGRAM", "JULY IG" vs "JULY HIGHLIGHT"), and differ in case and
    apostrophes ("MOTHERS’ DAY" vs "MOTHERS DAY") — all map to one name
    using "IG"."""
    text = re.sub(r"['‘’]", "", text.upper())
    text = re.sub(r"\b(?:INSTAGRAM|HIGHLIGHTS?)\b", "IG", text)
    return " ".join(text.split())


def find_dropbox_links(prs):
    """Hyperlinked text runs pointing to a dropbox.com URL, scanned from
    the deck's first few slides (its "Download Link/s" slide, one bulleted
    line per campaign/asset, hyperlinked on the label rather than shown as
    a plain URL). Skips any line explicitly labeled "(Video)"/"(Videos)"
    — this catalog only wants the photo folder."""
    links = []
    for slide in list(prs.slides)[:DROPBOX_LINK_SCAN_SLIDES]:
        if is_hidden(slide):
            continue
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                label = para.text.strip()
                if not label:
                    continue
                url = None
                for run in para.runs:
                    try:
                        addr = run.hyperlink.address
                    except Exception:
                        addr = None
                    if addr and "dropbox.com" in addr.lower():
                        url = addr
                        break
                if url and not _is_video_dropbox_label(label):
                    links.append({"label": label, "url": url})
    return links


# Words stripped before comparing a Dropbox link's label against a
# campaign name — sit around the meaningful part of the label ("TMP
# Summer Campaign") but never appear on the campaign banner ("SUMMER").
_DROPBOX_LABEL_FILLER_WORDS = {"tmp", "campaign"}


def _strip_trailing_parenthetical(text):
    # "TMP April IG (Batch 2)" -> "TMP April IG"; also drops "(Photos)".
    return re.sub(r"\s*\([^)]*\)\s*$", "", text).strip()


def _is_video_dropbox_label(label):
    m = re.search(r"\(([^)]*)\)\s*$", label)
    annotation = m.group(1) if m else ""
    return bool(re.search(r"\bvideos?\b", annotation, flags=re.IGNORECASE))


def _dropbox_match_tokens(text):
    # "&" and "and" are used interchangeably between the campaign banner
    # and the Dropbox label ("Health & Wellness" vs "Health And Wellness")
    # — normalize_title would otherwise drop "&" as punctuation, so spell
    # it out first to make both forms produce the same "and" token.
    text = text.replace("&", " and ")
    return set(normalize_title(text).split()) - _DROPBOX_LABEL_FILLER_WORDS


def _campaign_dropbox_tokens(campaign):
    # `campaign` is already normalize_campaign()'d (e.g. "SUMMER CAMPAIGN").
    return _dropbox_match_tokens(campaign)


def _label_dropbox_tokens(label):
    text = re.sub(r"^\s*TMP\s+", "", label, flags=re.IGNORECASE)
    text = _strip_trailing_parenthetical(text)
    return _dropbox_match_tokens(normalize_campaign(text))


def _label_base_and_annotation(label):
    """('health wellness', 'Batch 2') for 'TMP Health & Wellness (Batch
    2)': the label with "TMP"/trailing parenthetical removed (lowercased,
    for comparing whether two labels are otherwise identical), plus the
    raw parenthetical text ('' if none)."""
    text = re.sub(r"^\s*TMP\s+", "", label, flags=re.IGNORECASE)
    m = re.search(r"\(([^)]*)\)\s*$", text)
    annotation = m.group(1).strip() if m else ""
    base = _strip_trailing_parenthetical(text).strip().lower()
    return base, annotation


_BATCH_SUFFIX_RE = re.compile(r"^batch\s*(\d+)$", re.IGNORECASE)


def _resolve_duplicate_batch_links(found):
    """`found` is {url: label} for several links that all matched the same
    campaign. If they're the same label with only a trailing "(Batch N)"
    (or no suffix at all) distinguishing them, they're duplicates of the
    same folder rather than a genuine ambiguity — prefer the unsuffixed
    label, else the lowest-numbered batch. Returns (url, label), or None
    if the labels differ in some other way (a real ambiguity)."""
    parsed = [(url, label, *_label_base_and_annotation(label)) for url, label in found.items()]
    if len({base for _, _, base, _ in parsed}) != 1:
        return None

    unsuffixed = [p for p in parsed if not p[3]]
    if unsuffixed:
        url, label, _, _ = unsuffixed[0]
        return url, label

    numbered = []
    for url, label, _, annotation in parsed:
        m = _BATCH_SUFFIX_RE.match(annotation)
        if not m:
            return None
        numbered.append((int(m.group(1)), url, label))
    numbered.sort(key=lambda x: x[0])
    _, url, label = numbered[0]
    return url, label


def assign_dropbox_urls(campaigns, pairs):
    """Match each campaign's Dropbox photo-folder link (from every Recipe
    & Captions deck in this run) by comparing its label's meaningful words
    against the campaign name, and store it as campaigns[c]['dropboxUrl'].
    Never guesses: a campaign with zero or multiple matching links is left
    without one, and reported back as a warning to print/review instead.

    Two campaigns can coincidentally normalize to the same words when a
    deck's Download Link/s slide labels its link by shoot month rather
    than content month (seen: a July shoot's only link read "TMP July
    IG" while its items were filed as content month "AUGUST IG", and a
    separate June-shot "JULY IG" campaign's own link collided with it).
    To avoid that cross-contamination, each campaign is matched first
    against links from only the deck(s) that contributed its items —
    the full pool across every deck in the run is tried only if that
    comes up empty."""
    warnings = []
    links_by_deck = {pair["yymm"]: pair["dropbox_links"] for pair in pairs}
    all_links = [link for pair in pairs for link in pair["dropbox_links"]]
    for campaign, entry in campaigns.items():
        if not entry["items"]:
            continue  # PPM-only campaign; no Recipe & Captions items to attach a link to
        campaign_tokens = _campaign_dropbox_tokens(campaign)
        if not campaign_tokens:
            continue  # "(no banner text)" campaign; already flagged elsewhere
        item_decks = {it["_deck"] for it in entry["items"]}
        local_links = [link for deck in item_decks for link in links_by_deck.get(deck, [])]
        found = {}
        for pool in (local_links, all_links):
            found = {
                link["url"]: link["label"] for link in pool
                if _label_dropbox_tokens(link["label"]) == campaign_tokens
            }
            if found:
                break  # prefer the campaign's own deck(s); only widen the search if empty
        if len(found) == 1:
            entry["dropboxUrl"] = next(iter(found))
        elif len(found) == 0:
            warnings.append(
                f"no Dropbox link found for campaign {campaign!r} in the first "
                f"{DROPBOX_LINK_SCAN_SLIDES} slide(s) of any Recipe & Captions deck in this run "
                f"— check the deck manually"
            )
        else:
            resolved = _resolve_duplicate_batch_links(found)
            if resolved is not None:
                entry["dropboxUrl"], _ = resolved
            else:
                labels = ", ".join(repr(lbl) for lbl in found.values())
                warnings.append(
                    f"ambiguous Dropbox links for campaign {campaign!r} ({labels}) — "
                    f"leaving dropboxUrl out; check the deck manually"
                )
    return warnings


def banner_fill_color(shape):
    """Campaign banner fill as 'RRGGBB', or 'theme:<name>' for theme
    colors (seen on Mar 2026's APRIL INSTAGRAM banner). None if the
    banner has no solid fill."""
    try:
        if shape.fill.type != MSO_FILL_TYPE.SOLID:
            return None
        color = shape.fill.fore_color
        if color.type == MSO_COLOR_TYPE.RGB:
            return str(color.rgb)
        if color.type == MSO_COLOR_TYPE.SCHEME:
            return f"theme:{color.theme_color.name}"
    except Exception:
        pass
    return None


def clean_multiline(text):
    # collapse Google Slides' \x0b soft line breaks to real newlines
    return text.replace("\x0b", "\n").strip()


# ---------------------------------------------------------------------
# Brand detection: bolded runs in the Caption / Procedure text boxes
# ---------------------------------------------------------------------

_BOLD_PHRASE_TRIM_CHARS = " \t\n\r.,;:!?\"'()[]{}<>*_~`®™•·-–—"


def extract_bold_phrases(text_frame):
    """Return the raw text of each maximal run of adjacent bolded runs
    within a paragraph. Bold formatting sometimes splits a single word or
    phrase across multiple consecutive runs (e.g. 'Mondial ' + 'Real Thai'
    + ' Rice Paper', all bold, back-to-back) — those merge into one phrase.
    A non-bold run breaks the run of adjacent bold runs."""
    phrases = []
    for para in text_frame.paragraphs:
        current = []
        for run in para.runs:
            if run.font.bold:
                current.append(run.text)
            elif current:
                phrases.append("".join(current))
                current = []
        if current:
            phrases.append("".join(current))
    return phrases


def clean_bold_phrase(text):
    return text.strip(_BOLD_PHRASE_TRIM_CHARS)


def _normalize_for_brand_match(text):
    # lowercase and collapse everything but letters/digits to spaces, so
    # variants like "The Marketplace®" or trailing punctuation/emoji still
    # match the plain "The Marketplace" entry in EXCLUDED_BRAND_TERMS.
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def is_excluded_brand(phrase):
    normalized = _normalize_for_brand_match(phrase)
    if not normalized:
        return True
    for term in EXCLUDED_BRAND_TERMS:
        term_normalized = _normalize_for_brand_match(term)
        if term_normalized and term_normalized in normalized:
            return True
    return False


def extract_brands(caption_ph, procedure_ph):
    """Brand names for one item: bolded phrases from the Caption and
    Procedure text boxes, minus field labels and EXCLUDED_BRAND_TERMS,
    deduplicated case-insensitively (first-seen casing kept)."""
    raw_phrases = []
    for ph in (caption_ph, procedure_ph):
        if ph is not None and ph.has_text_frame:
            raw_phrases.extend(extract_bold_phrases(ph.text_frame))

    brands = []
    seen = set()
    for raw in raw_phrases:
        phrase = clean_bold_phrase(_LEADING_STRUCTURAL_LABEL.sub("", clean_bold_phrase(raw)))
        if not phrase:
            continue
        if phrase.lower() in BOLD_STRUCTURAL_LABELS:
            continue
        if is_excluded_brand(phrase):
            continue
        key = phrase.lower()
        if key in seen:
            continue
        seen.add(key)
        brands.append(phrase)
    return brands


def print_brand_summary(items):
    counts = {"0": 0, "1": 0, "2+": 0}
    unique_brands = {}
    for item in items:
        n = len(item.get("brand") or [])
        counts["0" if n == 0 else "1" if n == 1 else "2+"] += 1
        for b in item.get("brand") or []:
            unique_brands.setdefault(b.lower(), b)

    print(f"\nBrand extraction summary ({len(items)} slide(s)):")
    print(f"  0 brands:  {counts['0']}")
    print(f"  1 brand:   {counts['1']}")
    print(f"  2+ brands: {counts['2+']}")
    if unique_brands:
        print(f"\n  {len(unique_brands)} distinct brand string(s) found (eyeball for typos/near-dupes):")
        for b in sorted(unique_brands.values(), key=str.lower):
            print(f"    - {b!r}")
    else:
        print("  no brand strings found")


# ---------------------------------------------------------------------
# Pre-Prod deck: layout code -> category lookup
# ---------------------------------------------------------------------

def build_category_lookup(preprod_path):
    """Returns (lookup by layout code, lookup by slide title, warnings).

    The title lookup is only a fallback for Recipe & Captions items whose
    layout code isn't in the PPM deck — seen when PPM slides were never
    given their real code and still carry the template's "IG00"."""
    prs = Presentation(preprod_path)
    by_code = {}   # layout code -> [(slide number, category)]
    by_title = {}  # normalized title -> {categories}
    warnings = []

    for i, slide in enumerate(prs.slides):
        if is_hidden(slide):
            continue
        code_shapes = find_shape_in_zone(slide, ZONE_LAYOUT_CODE, {MSO_SHAPE_TYPE.AUTO_SHAPE})
        if not code_shapes:
            continue  # not an item slide (front matter / divider / logistics)
        layout_code = shape_text(code_shapes[0]).strip()
        if not layout_code:
            continue

        badge_shapes = find_shape_in_zone(slide, ZONE_CATEGORY_BADGE, {MSO_SHAPE_TYPE.AUTO_SHAPE})
        category = None
        for shape in badge_shapes:
            text = shape_text(shape).strip().upper()
            for known in KNOWN_CATEGORIES:
                if text == known:
                    category = known
                    break
            if category:
                break

        if category is None:
            warnings.append(
                f"[preprod slide {i + 1}] layout_code={layout_code!r}: "
                f"no recognized category badge found in zone "
                f"(badge texts seen: {[shape_text(s).strip() for s in badge_shapes]})"
            )
            continue

        by_code.setdefault(layout_code, []).append((i + 1, category))
        title_ph = find_placeholder(slide, idx=0, ptype=PP_PLACEHOLDER.TITLE)
        title = normalize_title(shape_text(title_ph)) if title_ph else ""
        if title:
            by_title.setdefault(title, set()).add(category)

    lookup = {}
    for layout_code, hits in by_code.items():
        categories = {category for _, category in hits}
        if len(categories) > 1:
            warnings.append(
                f"layout_code={layout_code!r} is on {len(hits)} PPM slides with different categories "
                f"(unfilled template code?) — not used for matching; those items fall back to title matching"
            )
            continue
        lookup[layout_code] = hits[0][1]

    # a title shared by slides with different categories can't decide anything
    title_lookup = {t: next(iter(c)) for t, c in by_title.items() if len(c) == 1}
    return lookup, title_lookup, warnings


def scan_preprod_banners(preprod_path):
    """Campaign banner on each Pre-Prod item slide, keyed to its layout
    code. Only used to list campaigns in the dry run and cross-check them
    against the Recipe & Captions deck — never to correlate slides."""
    prs = Presentation(preprod_path)
    rows = []
    for slide in prs.slides:
        if is_hidden(slide):
            continue
        code_shapes = find_shape_in_zone(slide, ZONE_LAYOUT_CODE, {MSO_SHAPE_TYPE.AUTO_SHAPE})
        if not code_shapes:
            continue
        layout_code = shape_text(code_shapes[0]).strip()
        if not layout_code:
            continue
        banner_shapes = find_shape_in_zone(slide, ZONE_CAMPAIGN_BANNER, {MSO_SHAPE_TYPE.AUTO_SHAPE})
        campaign_raw = shape_text(banner_shapes[0]).strip() if banner_shapes else ""
        rows.append({
            "layoutCode": layout_code,
            "campaign": normalize_campaign(campaign_raw),
            "campaignRaw": campaign_raw,
            "color": banner_fill_color(banner_shapes[0]) if banner_shapes else None,
        })
    return rows


# ---------------------------------------------------------------------
# Recipe & Captions deck: the real extraction
# ---------------------------------------------------------------------

def is_divider_slide(slide):
    shapes = list(slide.shapes)
    if len(shapes) != 1:
        return False
    idx, ptype = placeholder_idx_type(shapes[0])
    return ptype == PP_PLACEHOLDER.TITLE and shape_text(shapes[0]).strip() != ""


def get_item_photo_placeholder(slide):
    for shape in slide.shapes:
        idx, ptype = placeholder_idx_type(shape)
        if ptype == PP_PLACEHOLDER.PICTURE and in_zone(shape, ZONE_ITEM_PHOTO):
            return shape
    return None


def slide_has_recipe_text(slide):
    procedure = find_placeholder(slide, idx=1, ptype=PP_PLACEHOLDER.BODY)
    ingredients = find_placeholder(slide, idx=3, ptype=PP_PLACEHOLDER.BODY)
    if procedure is None or ingredients is None:
        return False
    proc_text = clean_multiline(shape_text(procedure))
    ing_text = clean_multiline(shape_text(ingredients))
    return proc_text.lower().startswith("procedure") and "ingredient" in ing_text.lower()


def parse_serving_block(raw_text):
    text = clean_multiline(raw_text)
    result = {"servingTime": "", "yield": "", "tips": ""}
    patterns = {
        "servingTime": r"Serving Time:\s*(.*)",
        "yield": r"Yield:\s*(.*)",
        "tips": r"Tips\s*\(Optional\):\s*(.*)",
    }
    for line in text.split("\n"):
        line = line.strip()
        for key, pat in patterns.items():
            m = re.match(pat, line, flags=re.IGNORECASE)
            if m:
                result[key] = m.group(1).strip()
    return result


def extract_photo(shape):
    image = shape.image
    img = Image.open(io.BytesIO(image.blob))
    img = img.convert("RGB")
    w, h = img.size
    scale = THUMB_MAX_DIM / max(w, h)
    if scale < 1:
        img = img.resize((max(1, round(w * scale)), max(1, round(h * scale))), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=THUMB_JPEG_QUALITY)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def yymm_from_filename(path):
    """'2604' from 'pptx/2604-RecipeCaptions_TMP Apr 2026 ....pptx'. Shoots
    are monthly, so the source filenames' YYMM prefix is the deck's date —
    there is no day component. None if the name has no valid prefix."""
    m = re.match(r"(\d{2})(\d{2})-", Path(path).name)
    if not m or not 1 <= int(m.group(2)) <= 12:
        return None
    return m.group(1) + m.group(2)


def resolve_deck_yymm(recipe_path, preprod_path):
    """The deck month from the two source filenames. When they disagree
    (or one lacks a YYMM- prefix), print both and ask — never pick one
    silently. Returns None if it can't be resolved; nothing is written."""
    recipe_yymm = yymm_from_filename(recipe_path)
    preprod_yymm = yymm_from_filename(preprod_path)
    if recipe_yymm and recipe_yymm == preprod_yymm:
        return recipe_yymm

    problem = "have different YYMM prefixes" if recipe_yymm and preprod_yymm else "don't both have a YYMM- prefix"
    print(f"The source deck filenames {problem}:")
    print(f"  Recipe & Captions deck: {recipe_yymm or '(none)':6}  {Path(recipe_path).name}")
    print(f"  PPM / Pre-Prod deck:    {preprod_yymm or '(none)':6}  {Path(preprod_path).name}")
    print("No output files have been written.")
    if not sys.stdin.isatty():
        print("Rename the decks so both start with the same YYMM-, or re-run in an interactive terminal.")
        return None
    try:
        return _ask("Which YYMM should the output use?", None, _parse_yymm)
    except (EOFError, KeyboardInterrupt):
        print("\nCancelled — no files written.")
        return None


def extract_items(recipe_path, category_lookup, warnings, title_lookup=None):
    prs = Presentation(recipe_path)
    slides = list(prs.slides)
    items = []
    current_shoot_type = None

    i = 0
    while i < len(slides):
        slide = slides[i]

        if is_hidden(slide):
            if get_item_photo_placeholder(slide) is not None:
                code_shapes = find_shape_in_zone(slide, ZONE_LAYOUT_CODE, {MSO_SHAPE_TYPE.AUTO_SHAPE})
                title_ph = find_placeholder(slide, idx=0, ptype=PP_PLACEHOLDER.TITLE)
                warnings.append(
                    f"[recipe slide {i + 1}] hidden slide skipped: "
                    f"{shape_text(code_shapes[0]).strip() if code_shapes else '?'} "
                    f"({shape_text(title_ph).strip() if title_ph else ''!r})"
                )
            i += 1
            continue

        if is_divider_slide(slide):
            current_shoot_type = shape_text(list(slide.shapes)[0]).strip()
            i += 1
            continue

        photo_ph = get_item_photo_placeholder(slide)
        if photo_ph is None:
            i += 1
            continue  # front matter / overview slide, not an item

        code_shapes = find_shape_in_zone(slide, ZONE_LAYOUT_CODE, {MSO_SHAPE_TYPE.AUTO_SHAPE})
        banner_shapes = find_shape_in_zone(slide, ZONE_CAMPAIGN_BANNER, {MSO_SHAPE_TYPE.AUTO_SHAPE})
        title_ph = find_placeholder(slide, idx=0, ptype=PP_PLACEHOLDER.TITLE)
        caption_ph = find_placeholder(slide, idx=1, ptype=PP_PLACEHOLDER.BODY)

        layout_code = shape_text(code_shapes[0]).strip() if code_shapes else ""
        campaign_raw = shape_text(banner_shapes[0]).strip() if banner_shapes else ""
        banner_color = banner_fill_color(banner_shapes[0]) if banner_shapes else None
        layout_name = shape_text(title_ph).strip() if title_ph else ""
        caption_raw = clean_multiline(shape_text(caption_ph)) if caption_ph else ""
        caption = strip_label(caption_raw, r"Caption:")

        if not layout_code:
            warnings.append(f"[recipe slide {i + 1}] no layout code found in zone; skipping item")
            i += 1
            continue

        try:
            photo_b64 = extract_photo(photo_ph)
        except Exception as e:
            warnings.append(f"[recipe slide {i + 1}] layout_code={layout_code!r}: photo extraction failed: {e}")
            photo_b64 = None

        item = {
            "layoutCode": layout_code,
            "layoutName": layout_name,
            "category": None,
            "campaign": normalize_campaign(campaign_raw),
            "shootType": current_shoot_type,
            "caption": caption,
            "brand": [],
            "recipe": None,
            "photo": {
                "filename": f"{layout_code}.jpg",
                "dropbox_path": "",
                "image_data": photo_b64,
            },
            "needsPhotoSwap": True,
            # internal (underscore keys are never written to output):
            # used for the dry-run campaign listing
            "_campaignRaw": campaign_raw,
            "_bannerColor": banner_color,
            "_slides": 1,
        }

        paired_next = (
            i + 1 < len(slides)
            and not is_hidden(slides[i + 1])
            and slide_has_recipe_text(slides[i + 1])
        )
        procedure_ph = None

        if paired_next:
            recipe_slide = slides[i + 1]
            next_title = find_placeholder(recipe_slide, idx=0, ptype=PP_PLACEHOLDER.TITLE)
            next_title_text = shape_text(next_title).strip() if next_title else ""
            if next_title_text and next_title_text != layout_name:
                warnings.append(
                    f"[recipe slide {i + 2}] title {next_title_text!r} != "
                    f"photo slide title {layout_name!r} (proceeding anyway)"
                )

            item["category"] = "RECIPE"

            serving_ph = find_placeholder(recipe_slide, idx=2, ptype=PP_PLACEHOLDER.BODY)
            ingredients_ph = find_placeholder(recipe_slide, idx=3, ptype=PP_PLACEHOLDER.BODY)
            procedure_ph = find_placeholder(recipe_slide, idx=1, ptype=PP_PLACEHOLDER.BODY)

            serving = parse_serving_block(shape_text(serving_ph)) if serving_ph else {
                "servingTime": "", "yield": "", "tips": ""
            }
            ingredients = strip_label(
                clean_multiline(shape_text(ingredients_ph)) if ingredients_ph else "", r"Ingredients:"
            )
            procedure = strip_label(
                clean_multiline(shape_text(procedure_ph)) if procedure_ph else "", r"Procedure:"
            )

            item["recipe"] = {
                "servingTime": serving["servingTime"],
                "yield": serving["yield"],
                "tips": serving["tips"],
                "ingredients": ingredients,
                "procedure": procedure,
            }
            item["_slides"] = 2
            i += 2
        else:
            category = category_lookup.get(layout_code)
            if category is None and title_lookup:
                category = title_lookup.get(normalize_title(layout_name))
                if category is not None:
                    # text match, not layout code: listed in the dry run for review
                    item["_categoryByTitle"] = True
            if category is None:
                warnings.append(
                    f"[recipe slide {i + 1}] layout_code={layout_code!r} ({layout_name!r}): "
                    f"no matching layout code or slide title in Pre-Prod deck; "
                    f"leaving category as 'UNKNOWN'"
                )
                category = "UNKNOWN"
            item["category"] = category
            i += 1

        item["brand"] = extract_brands(caption_ph, procedure_ph)
        items.append(item)

    return items, prs


# ---------------------------------------------------------------------
# Output: one JS file per campaign + content month
# ---------------------------------------------------------------------

def slugify(text):
    text = re.sub(r"[()/]", "", text)
    text = re.sub(r"[^A-Za-z0-9]+", "-", text).strip("-")
    return text


def write_js_file(out_path, shoot_type_label, items, generated_date, dropbox_url=None):
    var_name = re.sub(r"[^A-Z0-9]+", "_", shoot_type_label.upper()).strip("_") + "_DATA"
    lines = [
        f"// Shoot: {shoot_type_label}",
        f"// Generated: {generated_date.isoformat()}",
        f"// Items: {len(items)}",
        "",
        f"const {var_name} = [",
    ]
    blocks = []
    for item in items:
        payload = {k: v for k, v in item.items() if k != "shootType" and not k.startswith("_")}
        block = json.dumps(payload, indent=2, ensure_ascii=False)
        blocks.append("\n".join("  " + line for line in block.split("\n")))
    lines.append(",\n".join(blocks))
    lines.append("];")
    if dropbox_url:
        # Attached to the array itself (not a wrapper object) so existing
        # consumers that treat window.__recipesData as a plain items array
        # (e.g. index.html's data.forEach(...)) keep working unchanged.
        lines.append(f"{var_name}.dropboxUrl = {json.dumps(dropbox_url)};")
    lines.append("")
    lines.append("if (typeof window !== 'undefined') window.__recipesData = " + var_name + ";")
    out_path.write_text("\n".join(lines), encoding="utf-8")


# ---------------------------------------------------------------------
# Campaigns: dry-run listing, then map each campaign to a content month
# ---------------------------------------------------------------------

def load_deck_pair(yymm, recipe_path, preprod_path):
    """Read one PPM + Recipe & Captions pair into memory. Writes nothing."""
    warnings = []
    print(f"[{yymm}] Reading category lookup from Pre-Prod deck: {preprod_path}")
    category_lookup, title_lookup, lookup_warnings = build_category_lookup(preprod_path)
    warnings.extend(f"[{yymm}] {w}" for w in lookup_warnings)
    print(f"[{yymm}]   -> {len(category_lookup)} layout codes mapped to categories")

    print(f"[{yymm}] Reading items from Recipe & Captions deck: {recipe_path}")
    item_warnings = []
    items, prs = extract_items(recipe_path, category_lookup, item_warnings, title_lookup)
    warnings.extend(f"[{yymm}] {w}" for w in item_warnings)
    for item in items:
        item["_deck"] = yymm
    print(f"[{yymm}]   -> {len(items)} items extracted")

    dropbox_links = find_dropbox_links(prs)
    print(f"[{yymm}]   -> {len(dropbox_links)} Dropbox link(s) found in the first "
          f"{DROPBOX_LINK_SCAN_SLIDES} slide(s)")

    return {
        "yymm": yymm,
        "items": items,
        "preprod_banners": scan_preprod_banners(preprod_path),
        "dropbox_links": dropbox_links,
        "warnings": warnings,
    }


def collect_campaigns(pairs):
    """{campaign: {"rows": {(yymm, deck, raw text, color): counts}, "items": [...]}}
    in first-seen order, PPM before Recipe & Captions within each pair."""
    campaigns = {}

    def row(campaign, key):
        entry = campaigns.setdefault(campaign, {"rows": {}, "items": []})
        return entry, entry["rows"].setdefault(key, {"slides": 0, "items": 0})

    for pair in pairs:
        for b in pair["preprod_banners"]:
            _, counts = row(b["campaign"], (pair["yymm"], "PPM", b["campaignRaw"], b["color"]))
            counts["slides"] += 1
        for item in pair["items"]:
            entry, counts = row(item["campaign"], (pair["yymm"], "R&C", item["_campaignRaw"], item["_bannerColor"]))
            counts["slides"] += item["_slides"]
            counts["items"] += 1
            entry["items"].append(item)
    return campaigns


def correlation_checks(pairs):
    """Sanity checks on how the two decks of each pair line up. Slides are
    correlated by layout code (read by position), never by banner text —
    these checks only report where the banner text disagrees."""
    notes = []
    for pair in pairs:
        yymm = pair["yymm"]
        ppm_by_code = {}
        for b in pair["preprod_banners"]:
            ppm_by_code.setdefault(b["layoutCode"], b)

        relabeled = 0
        for item in pair["items"]:
            ppm = ppm_by_code.get(item["layoutCode"])
            if ppm is None:
                continue
            if ppm["campaign"] != item["campaign"]:
                notes.append(
                    f"[{yymm}] {item['layoutCode']}: PPM banner {ppm['campaignRaw']!r} vs "
                    f"Recipe & Captions banner {item['_campaignRaw']!r} — filed under "
                    f"the Recipe & Captions campaign ({item['campaign']!r})"
                )
            elif ppm["campaignRaw"].upper() != item["_campaignRaw"].upper():
                relabeled += 1
        if relabeled:
            notes.append(
                f"[{yymm}] {relabeled} layout code(s) have differently-worded banners across "
                f"the two decks that normalize to the same campaign (IG/Highlight/Instagram, apostrophes)"
            )

        by_title = [it for it in pair["items"] if it.get("_categoryByTitle")]
        if by_title:
            notes.append(
                f"[{yymm}] {len(by_title)} item(s) got their category by SLIDE TITLE, not layout code "
                f"(code missing from the PPM deck) — check these:"
            )
            for it in by_title:
                notes.append(f"    {it['layoutCode']:5} {it['layoutName']!r} -> {it['category']}")

        seen = {}
        for item in pair["items"]:
            seen.setdefault(item["layoutCode"], []).append(item["layoutName"])
        for code, names in seen.items():
            if len(names) > 1:
                notes.append(
                    f"[{yymm}] layout code {code!r} is used on {len(names)} different photo slides "
                    f"in the Recipe & Captions deck: {names} — they'll share {code}.jpg"
                )
    return notes


def print_dry_run(campaigns, notes):
    print("\n" + "=" * 70)
    print("DRY RUN — campaigns found (no output files have been written)")
    print("=" * 70)
    mappable = [c for c, e in campaigns.items() if e["items"]]
    for n, campaign in enumerate(mappable, 1):
        entry = campaigns[campaign]
        print(f"\n  [{n}] {campaign or '(no banner text)'}")
        for (yymm, deck, raw, color), counts in entry["rows"].items():
            size = (f"{counts['items']} item(s) / {counts['slides']} slide(s)" if deck == "R&C"
                    else f"{counts['slides']} slide(s)")
            print(f"        {yymm} {deck:3}  {raw or '(no banner text)'!r:34} fill={color or 'none':16} {size}")
        print(f"        dropboxUrl: {entry.get('dropboxUrl') or '(not found — see warnings below)'}")

    ppm_only = [c for c, e in campaigns.items() if not e["items"]]
    if ppm_only:
        print("\n  PPM deck only — no Recipe & Captions items, so nothing to output:")
        for campaign in ppm_only:
            for (yymm, deck, raw, color), counts in campaigns[campaign]["rows"].items():
                print(f"        {yymm} {deck:3}  {raw or '(no banner text)'!r:34} fill={color or 'none':16} "
                      f"{counts['slides']} slide(s)")

    print("\n  Correlation: PPM <-> Recipe & Captions slides are matched by layout code")
    print("  (read from its fixed position), never by banner text. Only when a code is")
    print("  missing from the PPM deck is the slide title tried instead (listed below).")
    for note in notes:
        print(f"  {note}" if note.startswith(" ") else f"  - {note}")
    return mappable


def _ask(prompt, default=None, parse=None):
    while True:
        suffix = f" [{default}]" if default else ""
        raw = input(f"{prompt}{suffix}: ").strip()
        if not raw and default:
            raw = default
        if not raw:
            continue
        try:
            return parse(raw) if parse else raw
        except ValueError as e:
            print(f"    {e}")


def _parse_yymm(raw):
    raw = raw.strip()
    m = re.fullmatch(r"(\d{2})(\d{2})", raw) or re.fullmatch(r"20(\d{2})-(\d{1,2})", raw)
    if not m:
        raise ValueError("enter the month as YYMM (e.g. 2603) or YYYY-MM (e.g. 2026-03)")
    yy, mm = m.group(1), int(m.group(2))
    if not 1 <= mm <= 12:
        raise ValueError(f"{raw!r} is not a real month")
    return f"{yy}{mm:02d}"


def _default_label(campaign):
    # "MARCH IG" -> "March IG"; short all-caps tokens (IG, &) stay as-is
    return " ".join(w if len(w) <= 2 else w.capitalize() for w in campaign.split())


def prompt_campaign_mapping(campaigns, mappable, out_dir):
    """Ask the target content month (no default) and file label for every
    campaign. Returns [(filename, label, items)] or None if cancelled."""
    print("\nMap each campaign to the content month it should be filed under.")
    print("There is no default: shoot month and content month often differ.")

    files = {}  # filename -> {"label": ..., "items": [...], "dropbox_urls": {url, ...}}
    for n, campaign in enumerate(mappable, 1):
        entry = campaigns[campaign]
        decks = sorted({it["_deck"] for it in entry["items"]})
        print(f"\n[{n}] {campaign or '(no banner text)'} — {len(entry['items'])} item(s) "
              f"from {', '.join(decks)}")
        yymm = _ask("    content month (YYMM)", None, _parse_yymm)
        label = _ask("    file label", _default_label(campaign) or "Unsorted")
        filename = f"{yymm}-{slugify(label)}.js"
        file_entry = files.setdefault(filename, {"label": label, "items": [], "dropbox_urls": set()})
        file_entry["items"].extend(entry["items"])
        if entry.get("dropboxUrl"):
            file_entry["dropbox_urls"].add(entry["dropboxUrl"])

    print("\nPlanned output:")
    for filename, file_entry in files.items():
        label, file_items, dropbox_urls = file_entry["label"], file_entry["items"], file_entry["dropbox_urls"]
        exists = "  (EXISTS — will be overwritten)" if (out_dir / filename).exists() else ""
        sources = ", ".join(
            f"{c} x{sum(1 for it in file_items if it['campaign'] == c)}"
            for c in dict.fromkeys(it["campaign"] for it in file_items)
        )
        print(f"  {out_dir / filename}  label={label!r}  {len(file_items)} item(s)  [{sources}]{exists}")
        codes = [it["layoutCode"] for it in file_items]
        dupes = sorted({c for c in codes if codes.count(c) > 1})
        if dupes:
            print(f"      ! duplicate layout code(s) in this file: {dupes}")
        if len(dropbox_urls) > 1:
            print(f"      ! {len(dropbox_urls)} different Dropbox links among the campaigns merged "
                  f"into this file — leaving dropboxUrl out")
    if input("\nWrite these files? [y/N]: ").strip().lower() not in ("y", "yes"):
        return None
    return [
        (filename, fe["label"], fe["items"], next(iter(fe["dropbox_urls"])) if len(fe["dropbox_urls"]) == 1 else None)
        for filename, fe in files.items()
    ]


def write_group(out_path, label, group_items, dropbox_url=None):
    write_js_file(out_path, label, group_items, date.today(), dropbox_url)
    recipe_count = sum(1 for it in group_items if it["category"] == "RECIPE")
    other_count = len(group_items) - recipe_count
    print(f"  wrote {out_path}  ({len(group_items)} items: {recipe_count} RECIPE, {other_count} other)")


def print_warnings(warnings, items):
    if warnings:
        print(f"\n{len(warnings)} warning(s):")
        for w in warnings:
            print(f"  - {w}")

    unknown = [it for it in items if it["category"] == "UNKNOWN"]
    if unknown:
        print(f"\n{len(unknown)} item(s) left with category=UNKNOWN (no Pre-Prod match) — needs manual fix:")
        for it in unknown:
            print(f"  - {it['layoutCode']} ({it['layoutName']!r})")


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--recipe-deck", action="append", required=True,
                        help="Recipe & Captions Deck .pptx (repeat for each deck pair)")
    parser.add_argument("--preprod-deck", action="append", required=True,
                        help="PPM / Pre-Prod Deck .pptx (repeat, same order as --recipe-deck)")
    parser.add_argument("--out-dir", default="recipes", help="Output folder (default: recipes)")
    parser.add_argument("--dry-run", action="store_true",
                        help="List campaigns found and exit without prompting or writing")
    args = parser.parse_args()

    if len(args.recipe_deck) != len(args.preprod_deck):
        parser.error("pass one --preprod-deck for each --recipe-deck (same order)")

    out_dir = Path(args.out_dir)

    deck_months = []
    for recipe_path, preprod_path in zip(args.recipe_deck, args.preprod_deck):
        yymm = resolve_deck_yymm(recipe_path, preprod_path)
        if yymm is None:
            return 1
        deck_months.append(yymm)

    pairs = [
        load_deck_pair(yymm, recipe_path, preprod_path)
        for yymm, recipe_path, preprod_path in zip(deck_months, args.recipe_deck, args.preprod_deck)
    ]
    items = [it for pair in pairs for it in pair["items"]]
    warnings = [w for pair in pairs for w in pair["warnings"]]

    campaigns = collect_campaigns(pairs)
    warnings.extend(assign_dropbox_urls(campaigns, pairs))
    mappable = print_dry_run(campaigns, correlation_checks(pairs))
    print_warnings(warnings, items)

    if args.dry_run:
        print("\nDry run only — no files written.")
        return 0
    if not mappable:
        print("\nNo Recipe & Captions items found — nothing to write.")
        return 1
    if not sys.stdin.isatty():
        print("\nRe-run in an interactive terminal to map each campaign to a month.")
        return 1

    try:
        if input("\nReview the list above. Continue to month mapping? [y/N]: ").strip().lower() not in ("y", "yes"):
            plan = None
        else:
            plan = prompt_campaign_mapping(campaigns, mappable, out_dir)
    except (EOFError, KeyboardInterrupt):
        plan = None
    if plan is None:
        print("\nCancelled — no files written.")
        return 1

    out_dir.mkdir(parents=True, exist_ok=True)
    print()
    for filename, label, file_items, dropbox_url in plan:
        write_group(out_dir / filename, label, file_items, dropbox_url)
    for filename, label, file_items, dropbox_url in plan:
        print(f"\n=== {filename} ===", end="")
        print_brand_summary(file_items)
    print("\nRun generate_recipes_index.py to add a recipes-index.js entry for each file.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
