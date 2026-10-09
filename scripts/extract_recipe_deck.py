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
                   OUT OF PACK W/ FOOD STYLING / NON-FOOD / GROUP SHOT / EAP) is
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
grouping: "Highlight" / "Instagram" -> "IG", the word "Campaign" is dropped,
case and apostrophes ignored.
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
from dataclasses import dataclass
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
    "EAP",
    "STOP MOTION",
]

# What the text-based rules produce. EAP / STOP MOTION only ever come from a
# PPM label, an overview SKU tag, or by hand, and existing ones are never overwritten.
INFERABLE_CATEGORIES = ("RECIPE", "OUT OF PACK", "NON-FOOD", "GROUP SHOT")

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


# ---------------------------------------------------------------------
# Template profiles: what a deck "looks like", and where to find things
# ---------------------------------------------------------------------

class DeckFormatError(Exception):
    """A deck doesn't match any known template profile."""


RECIPE_MARKER_RE = re.compile(
    r"^\s*﻿?(Serving Time|Yield|Ingredients?|Procedure)\s*:", re.IGNORECASE | re.MULTILINE
)
DIFFICULTY_CHIPS = {"BEGINNER", "INTERMEDIATE", "ADVANCED"}


@dataclass(frozen=True)
class TemplateProfile:
    """Everything position- or shape-dependent about one deck template.
    All element lookup (code, layout name, caption, recipe boxes, photo,
    banner, category badge) goes through here, so a new template (2021,
    2022, ...) is a new profile instance rather than new extraction logic.

    Zones are (min, max) inches from the slide's top-left."""
    name: str
    slide_size_in: tuple
    code_re: "re.Pattern"            # what a layout-code chip's text must look like
    zone_code: dict
    zone_banner: dict
    zone_photo: dict
    zone_badge: dict                 # Pre-Prod deck only
    min_share: float = 0.9           # share of item slides that must satisfy each fingerprint check
    size_tol_in: float = 0.05

    # -- element lookup ------------------------------------------------
    def _autoshapes_in(self, slide, zone):
        return find_shape_in_zone(slide, zone, {MSO_SHAPE_TYPE.AUTO_SHAPE})

    def codes_in_zone(self, slide):
        """(valid codes in z-order bottom->top, rejected texts). Decks are
        built by pasting over template slides, so chips can be stacked; a
        stray chip (e.g. a leftover '1') is rejected by the code pattern."""
        valid, rejected = [], []
        for shape in self._autoshapes_in(slide, self.zone_code):
            text = shape_text(shape).strip()
            if self.code_re.fullmatch(text):
                valid.append(text)
            elif text:
                rejected.append(text)
        return valid, rejected

    def code_of(self, slide):
        """(layout code, rejected texts). When valid chips are stacked, the
        topmost (last in z-order) is the one visible on the slide."""
        valid, rejected = self.codes_in_zone(slide)
        return (valid[-1] if valid else ""), rejected

    def banner_of(self, slide):
        """Topmost campaign banner with real text. Stacked banners are
        common (an old one left under the visible one), and untouched
        template placeholders ('[INSERT CAMPAIGN TITLE ...]') are skipped."""
        shapes = self._autoshapes_in(slide, self.zone_banner)
        real = [s for s in shapes if shape_text(s).strip() and not shape_text(s).strip().startswith("[")]
        pool = real or shapes
        return pool[-1] if pool else None

    def title_of(self, slide):
        return find_placeholder(slide, idx=0, ptype=PP_PLACEHOLDER.TITLE)

    def caption_of(self, slide):
        return find_placeholder(slide, idx=1, ptype=PP_PLACEHOLDER.BODY)

    def photo_of(self, slide):
        for shape in slide.shapes:
            idx, ptype = placeholder_idx_type(shape)
            if ptype == PP_PLACEHOLDER.PICTURE and in_zone(shape, self.zone_photo):
                return shape
        # some slides carry the photo as a plain picture instead of a placeholder
        pictures = [sh for sh in slide.shapes
                    if sh.shape_type == MSO_SHAPE_TYPE.PICTURE and in_zone(sh, self.zone_photo)
                    and sh.width and sh.width >= Emu(int(2 * 914400))]
        if pictures:
            return max(pictures, key=lambda sh: sh.width * sh.height)
        return None

    def badge_texts(self, slide):
        return [shape_text(s).strip() for s in self._autoshapes_in(slide, self.zone_badge)]

    def recipe_markers(self, slide):
        """Distinct recipe field labels (serving time / yield / ingredients /
        procedure) found anywhere on the slide — by content, not by
        placeholder index, reading order or position."""
        found = set()
        for shape in slide.shapes:
            if shape.has_text_frame:
                for m in RECIPE_MARKER_RE.finditer(clean_multiline(shape.text_frame.text)):
                    found.add(m.group(1).lower().rstrip("s"))
        return found

    def has_recipe_content(self, slide):
        return len(self.recipe_markers(slide)) >= 2

    def recipe_boxes(self, slide):
        """{'serving', 'ingredients', 'procedure'} text boxes of a recipe
        slide, located by what they say. When several boxes qualify (a
        filled box pasted over an empty leftover placeholder) the one with
        the most text wins. Falls back to the template's placeholder
        indices if no box announces itself."""
        wanted = {
            "serving": (r"\s*﻿?Serving Time", 2),
            "ingredients": (r"\s*﻿?Ingredient", 3),
            "procedure": (r"\s*﻿?Procedure", 1),
        }
        boxes = {}
        for key, (pattern, idx) in wanted.items():
            matches = [s for s in slide.shapes
                       if s.has_text_frame and re.match(pattern, clean_multiline(s.text_frame.text), re.IGNORECASE)]
            if matches:
                matches.sort(key=lambda s: len(shape_text(s).strip()), reverse=True)
                boxes[key] = matches[0]
            else:
                boxes[key] = find_placeholder(slide, idx=idx, ptype=PP_PLACEHOLDER.BODY)
        return boxes

    # -- fingerprint ---------------------------------------------------
    def _item_and_recipe_slides(self, prs):
        items, recipes = [], []
        for slide in prs.slides:
            if is_hidden(slide):
                continue
            if self.photo_of(slide) is not None:
                items.append(slide)
            elif self.has_recipe_content(slide):
                recipes.append(slide)
        return items, recipes

    def mismatches(self, prs, kind="recipe"):
        """Why this deck is NOT this template (empty list = it matches).
        kind='recipe' checks a Recipe & Captions deck, 'ppm' a Pre-Prod deck."""
        fails = []
        w, h = Emu(prs.slide_width).inches, Emu(prs.slide_height).inches
        ew, eh = self.slide_size_in
        if abs(w - ew) > self.size_tol_in or abs(h - eh) > self.size_tol_in:
            fails.append(f"slide size is {w:.2f} x {h:.2f} in, expected {ew} x {eh} in")

        if kind == "ppm":
            slides = [s for s in prs.slides if not is_hidden(s) and self.code_of(s)[0]]
            if not slides:
                fails.append("no slide has a layout-code chip matching "
                             f"{self.code_re.pattern!r} in the expected position")
                return fails
            with_badge = [s for s in slides
                          if any(t.upper() in KNOWN_CATEGORIES for t in self.badge_texts(s))]
            if len(with_badge) < self.min_share * len(slides):
                fails.append(f"category badge in the expected position on only {len(with_badge)}/{len(slides)} "
                             f"layout slides")
            return fails

        items, recipes = self._item_and_recipe_slides(prs)
        if not items:
            fails.append("no item slides (no picture placeholder in the expected photo position)")
            return fails

        def banner_ok(s):
            b = self.banner_of(s)
            return b is not None and shape_text(b).strip() != ""

        def title_ok(s):
            t = self.title_of(s)
            return t is not None and shape_text(t).strip() != ""

        checks = [
            (f"layout-code chip matching {self.code_re.pattern!r}", lambda s: bool(self.code_of(s)[0])),
            ("campaign banner text", banner_ok),
            ("title placeholder", title_ok),
            ("caption text box starting with 'Caption:'",
             lambda s: self.caption_of(s) is not None
             and shape_text(self.caption_of(s)).lstrip("\ufeff \n\x0b").lower().startswith("caption")),
            ("NO Pre-Prod category badge (a deck with one is a Pre-Prod deck)",
             lambda s: not any(t.upper() in KNOWN_CATEGORIES for t in self.badge_texts(s))),
        ]
        for label, ok in checks:
            n = sum(1 for s in items if ok(s))
            if n < self.min_share * len(items):
                fails.append(f"{label} found on only {n}/{len(items)} item slides")
        if recipes:
            n = sum(1 for s in recipes
                    if DIFFICULTY_CHIPS <= {shape_text(sh).strip().upper() for sh in s.shapes})
            if n < self.min_share * len(recipes):
                fails.append(f"BEGINNER/INTERMEDIATE/ADVANCED chips on only {n}/{len(recipes)} recipe slides")
        return fails


PROFILE_2023_2026 = TemplateProfile(
    name="TMP 2023-2026",
    slide_size_in=(10.0, 5.625),
    code_re=re.compile(r"[A-Za-z]{1,6}\d{1,3}"),
    zone_code=dict(left=(0.10, 0.65), top=(0.40, 0.90)),
    zone_banner=dict(left=(0.95, 1.85), top=(-0.05, 0.15)),
    zone_photo=dict(left=(0.10, 0.65), top=(0.90, 1.65)),
    # The full legend of category options also lives on every Pre-Prod slide
    # but is parked off-canvas at left=-1.19in — any positive-left zone
    # naturally excludes it.
    zone_badge=dict(left=(6.30, 7.30), top=(0.40, 0.75)),
)

# Add 2021 / 2022 profiles here (bare-number codes, category printed on each slide, ...).
PROFILES = [PROFILE_2023_2026]


def detect_profile(prs, label, kind="recipe"):
    """The first profile this deck matches. If none match, raises
    DeckFormatError naming the deck and, per profile, what didn't match —
    nothing is guessed."""
    reasons = []
    for profile in PROFILES:
        fails = profile.mismatches(prs, kind)
        if not fails:
            return profile
        reasons.append((profile.name, fails))
    kind_name = "Pre-Prod" if kind == "ppm" else "Recipe & Captions"
    lines = [f"{label}: matches no known template profile ({kind_name} deck)."]
    for name, fails in reasons:
        lines.append(f"  profile {name!r} did not match:")
        lines.extend(f"    - {f}" for f in fails)
    raise DeckFormatError("\n".join(lines))


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
    using "IG". The word "CAMPAIGN" is dropped ("EASTER CAMPAIGN" ->
    "EASTER") to match the catalog's shortened display names."""
    text = re.sub(r"['‘’]", "", text.upper())
    text = re.sub(r"\b(?:INSTAGRAM|HIGHLIGHTS?)\b", "IG", text)
    stripped = " ".join(re.sub(r"\bCAMPAIGNS?\b", " ", text).split())
    # A banner that is only the word "CAMPAIGN" keeps it rather than going blank.
    return stripped or " ".join(text.split())


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
_DROPBOX_LABEL_FILLER_WORDS = {"tmp", "campaign", "collaboration"}


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
# Pre-Prod deck (optional): layout code -> category lookup
# ---------------------------------------------------------------------

def build_category_lookup(prs, profile):
    """Returns (lookup by layout code, lookup by slide title, warnings).

    The title lookup is only a fallback for Recipe & Captions items whose
    layout code isn't in the PPM deck — seen when PPM slides were never
    given their real code and still carry the template's "IG00"."""
    by_code = {}   # layout code -> [(slide number, category)]
    by_title = {}  # normalized title -> {categories}
    warnings = []

    for i, slide in enumerate(prs.slides):
        if is_hidden(slide):
            continue
        layout_code, _ = profile.code_of(slide)
        if not layout_code:
            continue  # not an item slide (front matter / divider / logistics)

        category = None
        for text in profile.badge_texts(slide):
            if text.upper() in KNOWN_CATEGORIES:
                category = text.upper()
                break

        if category is None:
            warnings.append(
                f"[preprod slide {i + 1}] layout_code={layout_code!r}: "
                f"no recognized category badge found in zone "
                f"(badge texts seen: {profile.badge_texts(slide)})"
            )
            continue

        by_code.setdefault(layout_code, []).append((i + 1, category))
        title_ph = profile.title_of(slide)
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


def scan_preprod_banners(prs, profile):
    """Campaign banner on each Pre-Prod item slide, keyed to its layout
    code. Only used to list campaigns in the dry run and cross-check them
    against the Recipe & Captions deck — never to correlate slides."""
    rows = []
    for slide in prs.slides:
        if is_hidden(slide):
            continue
        layout_code, _ = profile.code_of(slide)
        if not layout_code:
            continue
        banner = profile.banner_of(slide)
        campaign_raw = shape_text(banner).strip() if banner is not None else ""
        rows.append({
            "layoutCode": layout_code,
            "campaign": normalize_campaign(campaign_raw),
            "campaignRaw": campaign_raw,
            "color": banner_fill_color(banner) if banner is not None else None,
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
    """The deck month from the source filename(s). When a PPM deck is given
    and they disagree (or one lacks a YYMM- prefix), print both and ask —
    never pick one silently. Returns None if it can't be resolved; nothing
    is written. Without a PPM deck only the Recipe & Captions name matters."""
    recipe_yymm = yymm_from_filename(recipe_path)
    preprod_yymm = yymm_from_filename(preprod_path) if preprod_path else None
    if recipe_yymm and (preprod_path is None or recipe_yymm == preprod_yymm):
        return recipe_yymm

    if preprod_path is None:
        problem = "doesn't have a YYMM- prefix"
    else:
        problem = "have different YYMM prefixes" if recipe_yymm and preprod_yymm else "don't both have a YYMM- prefix"
    print(f"The source deck filename(s) {problem}:")
    print(f"  Recipe & Captions deck: {recipe_yymm or '(none)':6}  {Path(recipe_path).name}")
    if preprod_path:
        print(f"  PPM / Pre-Prod deck:    {preprod_yymm or '(none)':6}  {Path(preprod_path).name}")
    print("No output files have been written.")
    if not sys.stdin.isatty():
        print("Rename the deck(s) so they start with the same YYMM-, or re-run in an interactive terminal.")
        return None
    try:
        return _ask("Which YYMM should the output use?", None, _parse_yymm)
    except (EOFError, KeyboardInterrupt):
        print("\nCancelled — no files written.")
        return None


# ---------------------------------------------------------------------
# Overview slide: "(N LAYOUTS)" / "(N VIDEOS)" per campaign
# ---------------------------------------------------------------------

OVERVIEW_SCAN_SLIDES = 10
_OVERVIEW_COUNT_RE = re.compile(r"\(\s*(\d+)\s*(layouts?|videos?)\s*\)\s*$", re.IGNORECASE)
_OVERVIEW_HEADING_RE = re.compile(r"^(?P<name>.+?)\s*\((?P<paren>[^()]*)\)\s*$")


def parse_overview(prs):
    """Headings from the deck's overview slide, in order: {'name', 'n',
    'unit' ('layouts'|'videos'|None), 'sub'}. 'sub' headings are photo-style
    sub-groups ("Lifestyle Group Photo (1 layout)") under the campaign
    heading above them; the same sub-heading also covers recipe layouts, so
    it's only ever a Group Shot clue. Best effort: [] if the deck's overview
    carries no counts (2025 decks mostly don't)."""
    for slide in list(prs.slides)[:OVERVIEW_SCAN_SLIDES]:
        if is_hidden(slide):
            continue
        heads = []
        has_count = False
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for line in re.split(r"[\n\x0b]", shape.text_frame.text):
                line = line.strip()
                if not line:
                    continue
                if line[0].isdigit():  # SKU line: belongs to the heading above it
                    if heads:
                        heads[-1]["skus"].append(line)
                    continue
                m = _OVERVIEW_HEADING_RE.match(line)
                if not m:
                    continue
                cm = _OVERVIEW_COUNT_RE.search(line)
                has_count = has_count or bool(cm)
                heads.append({
                    "name": m.group("name").strip(),
                    "n": int(cm.group(1)) if cm else None,
                    "unit": ("videos" if cm.group(2).lower().startswith("video") else "layouts") if cm else None,
                    "sub": "photo" in m.group("name").lower(),
                    "skus": [],
                })
        if has_count:
            return heads
    return []


_OVERVIEW_TAG_RE = re.compile(r"\(([^()]*)\)\s*$")


def overview_sku_tag(sku_line):
    """Category named by a trailing parenthetical on an overview SKU line:
    "(EAP)" -> EAP, "(Stop Motion Video)" -> STOP MOTION. A plain "(Video)"
    names no category."""
    m = _OVERVIEW_TAG_RE.search(sku_line)
    if not m:
        return None
    tag = m.group(1).strip().lower()
    if "stop motion" in tag or "stop-motion" in tag:
        return "STOP MOTION"
    if re.search(r"\beap\b", tag):
        return "EAP"
    return None


def summarize_overview(items, heads):
    """One line per overview heading saying what the cross-check did."""
    if not heads:
        return ["overview: no '(N layouts/videos)' counts found — cross-check skipped"]
    codes = {}
    for it in items:
        c = codes.setdefault(it["campaign"], [])
        if it["layoutCode"] not in c:
            c.append(it["layoutCode"])
    lines, parent, totals, subtot = [], None, {}, {}
    for h in heads:
        if h["sub"]:
            if parent is not None and h["n"] is not None:
                subtot[parent] = subtot.get(parent, 0) + h["n"]
        else:
            parent = normalize_campaign(h["name"])
            if h["n"] is not None:
                totals[parent] = totals.get(parent, 0) + h["n"]
    for campaign, n in totals.items():
        have = len(codes.get(campaign, []))
        lines.append(f"overview: {campaign!r} lists {n}, deck has {have} distinct layout code(s) -> "
                     + ("match" if n == have else "MISMATCH"))
    for campaign, n in subtot.items():
        have = len(codes.get(campaign, []))
        lines.append(f"overview: sub-groups under {campaign!r} list {n}, deck has {have} -> "
                     + ("match, group clue applied" if n == have else "no match, group clue skipped"))
    return lines


def apply_overview(items, heads, warnings):
    """Cross-check the overview's per-campaign counts against the distinct
    layout codes extracted, and tag each layout with the overview group it
    falls in (item['_ovUnit'] = 'videos'|'layouts', item['_ovGroup'] = a
    sub-heading name). Counts are only used when they add up to the number
    of layouts found; a mismatch is warned about, never guessed around."""
    if not heads:
        return
    codes_by_campaign = {}
    for it in items:
        codes = codes_by_campaign.setdefault(it["campaign"], [])
        if it["layoutCode"] not in codes:
            codes.append(it["layoutCode"])

    parent = None
    segs, subs = {}, {}
    for h in heads:
        if h["sub"]:
            if parent is not None and h["n"] is not None:
                subs.setdefault(parent, []).append(h)
        else:
            parent = normalize_campaign(h["name"])
            if h["n"] is not None:
                segs.setdefault(parent, []).append(h)

    def assign(campaign, groups, setter):
        codes = codes_by_campaign[campaign]
        pos = 0
        for g in groups:
            for k, code in enumerate(codes[pos:pos + g["n"]]):
                for it in items:
                    if it["campaign"] == campaign and it["layoutCode"] == code:
                        setter(it, g, k)
            pos += g["n"]

    for campaign, groups in segs.items():
        codes = codes_by_campaign.get(campaign)
        if codes is None:
            continue
        total = sum(g["n"] for g in groups)
        if total != len(codes):
            listing = " + ".join(f"{g['n']} {g['unit']}" for g in groups)
            warnings.append(
                f"overview lists {listing} for {campaign!r} but {len(codes)} distinct layout code(s) "
                f"were extracted ({', '.join(codes)})"
            )
            continue
        def tag(it, g, k):
            it["_ovUnit"] = g["unit"]
            # SKU lines pair with layouts one-to-one only when their count matches the group's
            if len(g["skus"]) == g["n"]:
                t = overview_sku_tag(g["skus"][k])
                if t:
                    it["_ovTag"] = t
        assign(campaign, groups, tag)
    for campaign, groups in subs.items():
        codes = codes_by_campaign.get(campaign)
        if codes is None or sum(g["n"] for g in groups) != len(codes):
            continue  # best effort: skip silently
        assign(campaign, groups, lambda it, g, k: it.__setitem__("_ovGroup", g["name"]))


# ---------------------------------------------------------------------
# Category inference (used only when there is no Pre-Prod deck)
# ---------------------------------------------------------------------

# Deliberately small and conservative: a keyword hit gives "medium"
# confidence, no hit gives "low". Matched as whole words/phrases against the
# layout name + caption.
NON_FOOD_KEYWORDS = [
    "laundry", "fabric conditioner", "fabric softener", "detergent", "downy",
    "cleaning", "cleaner", "dishwashing", "shampoo", "body wash", "face wash",
    "gentle wash", "lotion", "skincare", "soap", "toothpaste", "diaper",
    "sanitizer", "personal care", "air freshener", "freshener", "automatic spray", "auto spray",
    "fragrance", "feminine", "mildew", "tote bag", "bathroom",
]
GROUP_SHOT_NAME_RE = re.compile(r"\bgroup\s+(?:shot|photo)\b", re.IGNORECASE)


def _keyword_hit(text, keywords):
    text = text.lower()
    for kw in keywords:
        if re.search(r"(?<![a-z0-9])" + re.escape(kw) + r"(?![a-z0-9])", text):
            return kw
    return None


def infer_category(item):
    """(category, confidence, reason) from the Recipe & Captions deck alone.
    RECIPE is decided by slide CONTENT (recipe text on the layout's slides);
    slide count is only a cross-check, and a disagreement forces "low"."""
    has_recipe_content = item["recipe"] is not None or item["_inlineRecipe"]
    n_slides = item["_slides"] + item["_extraSlides"]
    video = item.get("_ovUnit") == "videos" or "video" in (item.get("shootType") or "").lower()

    if item.get("_ovTag"):
        return item["_ovTag"], "medium", f"overview SKU line tagged ({item['_ovTag']})"

    def finish(category, confidence, reason):
        if video and confidence != "low":
            reason += "; in overview video group, review"
            confidence = "low"
        return category, confidence, reason

    if has_recipe_content:
        confidence, reason = "high", "recipe text on the layout's slides"
        if item["_inlineRecipe"] and item["recipe"] is None:
            confidence, reason = "low", "recipe text sits on the photo slide but there is no separate recipe slide"
        elif item["_pairMethod"] == "adjacent-nocode":
            confidence, reason = "medium", "recipe slide has no layout code; attached to the slide before it"
        elif item["_pairMethod"] == "adjacent-codemismatch":
            confidence, reason = "low", "recipe slide's code matches no photo slide; attached to the slide before it"
        elif item["_pairMethod"] == "adjacent-title":
            confidence, reason = "medium", "recipe slide's code differs from the photo slide; matched by adjacency + identical title"
        elif item["_pairDisagree"]:
            confidence, reason = "medium", "code-pairing and adjacency pointed at different recipe slides"
        if item["_extraSlides"]:
            confidence, reason = "low", f"{item['_extraSlides']} extra slide(s) share this layout code"
        return finish("RECIPE", confidence, reason)

    # non-recipe subtype
    cross = "low" if n_slides >= 2 else None  # content says no recipe, slide count says there is
    text = f"{item['layoutName']}\n{item['caption']}"
    if GROUP_SHOT_NAME_RE.search(item["layoutName"]):
        category, confidence, reason = "GROUP SHOT", "low", "'group shot' in layout name"
    elif "group" in (item.get("_ovGroup") or "").lower():
        category, confidence, reason = "GROUP SHOT", "low", f"overview heading {item['_ovGroup']!r}"
    else:
        kw = _keyword_hit(text, NON_FOOD_KEYWORDS)
        if kw:
            category, confidence, reason = "NON-FOOD", "medium", f"keyword {kw!r}"
        else:
            category, confidence, reason = "OUT OF PACK", "low", "default for single-slide food layouts"
    if cross:
        confidence, reason = "low", reason + "; but slide count says recipe"
    return finish(category, confidence, reason)


def extract_items(prs, profile, ppm=None, warnings=None, with_photos=True):
    """Items from a Recipe & Captions deck.

    ppm: (category_lookup, title_lookup) from the Pre-Prod deck, or None to
    infer categories from this deck alone.

    Recipe slides are paired to photo slides by layout code first,
    adjacency only as a fallback; disagreements are warned about."""
    if warnings is None:
        warnings = []
    slides = list(prs.slides)
    photo_recs, recipe_recs, other_recs = [], [], []
    current_shoot_type = None

    for i, slide in enumerate(slides):
        if is_hidden(slide):
            if profile.photo_of(slide) is not None:
                code, _ = profile.code_of(slide)
                title_ph = profile.title_of(slide)
                warnings.append(
                    f"[recipe slide {i + 1}] hidden slide skipped: "
                    f"{code or '?'} "
                    f"({shape_text(title_ph).strip() if title_ph else ''!r})"
                )
            continue
        if is_divider_slide(slide):
            current_shoot_type = shape_text(list(slide.shapes)[0]).strip()
            continue

        photo_ph = profile.photo_of(slide)
        code, rejected = profile.code_of(slide)
        valid_codes, _ = profile.codes_in_zone(slide)
        if len(set(valid_codes)) > 1:
            warnings.append(
                f"[recipe slide {i + 1}] stacked layout-code chips {valid_codes}; "
                f"using the topmost ({valid_codes[-1]!r})"
            )
        if photo_ph is not None:
            t = profile.title_of(slide)
            photo_recs.append({"i": i, "slide": slide, "photo": photo_ph, "code": code,
                               "rejected": rejected, "shoot": current_shoot_type,
                               "title": normalize_title(shape_text(t)) if t else ""})
        elif profile.has_recipe_content(slide):
            t = profile.title_of(slide)
            recipe_recs.append({"i": i, "slide": slide, "code": code, "claimed": False,
                                "title": normalize_title(shape_text(t)) if t else ""})
        elif code:
            other_recs.append({"i": i, "code": code})
            cap = profile.caption_of(slide)
            if cap is not None and shape_text(cap).lstrip("\ufeff \n\x0b").lower().startswith("caption"):
                warnings.append(
                    f"[recipe slide {i + 1}] has layout code {code!r} and a caption but no photo was found "
                    f"in the photo position — not extracted as an item"
                )

    # ---- pair recipe slides to photo slides: layout code first, adjacency as fallback
    photo_codes = {r["code"] for r in photo_recs if r["code"]}
    by_index = {r["i"]: r for r in recipe_recs}
    for rec in photo_recs:
        rec.update(recipe=None, pair_method=None, pair_disagree=False)
    # pass 0: the directly-following recipe slide with the same title *and* code
    # is the unambiguous case; one with the same title but a different code is
    # a copy-pasted code typo (seen in Mar/Jun 2026 decks) and is accepted with a warning.
    for rec in photo_recs:
        adj = by_index.get(rec["i"] + 1)
        if adj is None or adj["claimed"] or not rec["title"] or adj["title"] != rec["title"]:
            continue
        adj["claimed"] = True
        if adj["code"] == rec["code"]:
            rec.update(recipe=adj, pair_method="code")
        else:
            rec.update(recipe=adj, pair_method="adjacent-title", pair_disagree=True)
            warnings.append(
                f"[recipe slide {adj['i'] + 1}] code {adj['code'] or 'none'!r} differs from photo slide "
                f"{rec['code']!r} but the title matches ({rec['title']!r}) — paired by adjacency + title"
            )
    for rec in photo_recs:  # pass 1: same layout code
        if not rec["code"] or rec["recipe"] is not None:
            continue
        same = [r for r in recipe_recs if r["code"] == rec["code"] and not r["claimed"]]
        if not same:
            continue
        same.sort(key=lambda r: (abs(r["i"] - rec["i"]), r["i"]))
        chosen = same[0]
        chosen["claimed"] = True
        rec.update(recipe=chosen, pair_method="code")
        adj = by_index.get(rec["i"] + 1)
        if adj is not None and adj is not chosen and adj["code"] != rec["code"]:
            rec["pair_disagree"] = True
            warnings.append(
                f"[recipe slide {rec['i'] + 1}] layout {rec['code']!r}: recipe slide by code is slide "
                f"{chosen['i'] + 1}, but slide {adj['i'] + 1} (code {adj['code'] or 'none'!r}) directly follows "
                f"the photo slide — pairing by code"
            )
        elif chosen["i"] != rec["i"] + 1:
            warnings.append(
                f"[recipe slide {rec['i'] + 1}] layout {rec['code']!r}: recipe slide {chosen['i'] + 1} "
                f"is not adjacent to the photo slide — paired by layout code"
            )
    unrecipe_codes = {r["code"] for r in photo_recs if r["recipe"] is None}
    for rec in photo_recs:  # pass 2: adjacency fallback
        if rec["recipe"] is not None:
            continue
        adj = by_index.get(rec["i"] + 1)
        if adj is None or adj["claimed"]:
            continue
        if not adj["code"]:
            adj["claimed"] = True
            rec.update(recipe=adj, pair_method="adjacent-nocode")
        elif adj["code"] not in unrecipe_codes:
            # its code belongs to a layout that already has its recipe (or to none)
            adj["claimed"] = True
            rec.update(recipe=adj, pair_method="adjacent-codemismatch")
            warnings.append(
                f"[recipe slide {adj['i'] + 1}] code {adj['code']!r} is not this layout's ({rec['code']!r}) and "
                f"is not wanted by any other photo slide; paired with the photo slide before it by adjacency"
            )
    for r in recipe_recs:
        if not r["claimed"]:
            warnings.append(
                f"[recipe slide {r['i'] + 1}] recipe slide (code {r['code'] or 'none'!r}) "
                f"not paired with any photo slide; ignored"
            )

    # ---- build items
    items = []
    for rec in photo_recs:
        i, slide, layout_code = rec["i"], rec["slide"], rec["code"]
        banner = profile.banner_of(slide)
        title_ph = profile.title_of(slide)
        caption_ph = profile.caption_of(slide)

        campaign_raw = shape_text(banner).strip() if banner is not None else ""
        banner_color = banner_fill_color(banner) if banner is not None else None
        layout_name = shape_text(title_ph).strip() if title_ph else ""
        caption_raw = clean_multiline(shape_text(caption_ph)) if caption_ph else ""
        caption = strip_label(caption_raw, r"Caption:")

        if not layout_code:
            seen = f" (chip text rejected: {rec['rejected']})" if rec["rejected"] else ""
            warnings.append(f"[recipe slide {i + 1}] no layout code found in zone; skipping item{seen}")
            continue

        photo_b64 = None
        if with_photos:
            try:
                photo_b64 = extract_photo(rec["photo"])
            except Exception as e:
                warnings.append(f"[recipe slide {i + 1}] layout_code={layout_code!r}: photo extraction failed: {e}")

        item = {
            "layoutCode": layout_code,
            "layoutName": layout_name,
            "category": None,
            "categorySource": None,
            "categoryConfidence": None,
            "campaign": normalize_campaign(campaign_raw),
            "shootType": rec["shoot"],
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
            # used for the dry-run campaign listing and category inference
            "_campaignRaw": campaign_raw,
            "_bannerColor": banner_color,
            "_slides": 1,
            "_slideNo": i + 1,
            "_inlineRecipe": profile.has_recipe_content(slide),
            "_pairMethod": rec["pair_method"],
            "_pairDisagree": rec["pair_disagree"],
            "_extraSlides": sum(1 for o in other_recs if o["code"] == layout_code),
        }

        procedure_ph = None
        recipe_rec = rec["recipe"]
        if recipe_rec is not None:
            recipe_slide = recipe_rec["slide"]
            next_title = profile.title_of(recipe_slide)
            next_title_text = shape_text(next_title).strip() if next_title else ""
            if next_title_text and next_title_text != layout_name:
                warnings.append(
                    f"[recipe slide {recipe_rec['i'] + 1}] title {next_title_text!r} != "
                    f"photo slide title {layout_name!r} (proceeding anyway)"
                )

            boxes = profile.recipe_boxes(recipe_slide)
            serving_ph, ingredients_ph, procedure_ph = boxes["serving"], boxes["ingredients"], boxes["procedure"]
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

        item["brand"] = extract_brands(caption_ph, procedure_ph)
        items.append(item)

    apply_overview(items, parse_overview(prs), warnings)

    # ---- categories
    for item in items:
        if ppm is not None:
            category_lookup, title_lookup = ppm
            code = item["layoutCode"]
            if item["recipe"] is not None:
                # a Pre-Prod EAP slide with recipe slides keeps its recipe, but files as EAP
                item["category"] = "EAP" if category_lookup.get(code) == "EAP" else "RECIPE"
                item["categoryConfidence"] = "high"
            else:
                category = category_lookup.get(code)
                confidence = "high"
                if category is None and title_lookup:
                    category = title_lookup.get(normalize_title(item["layoutName"]))
                    if category is not None:
                        # text match, not layout code: listed in the dry run for review
                        item["_categoryByTitle"] = True
                        confidence = "medium"
                if category is None:
                    warnings.append(
                        f"[recipe slide {item['_slideNo']}] layout_code={code!r} ({item['layoutName']!r}): "
                        f"no matching layout code or slide title in Pre-Prod deck; "
                        f"leaving category as 'UNKNOWN'"
                    )
                    category, confidence = "UNKNOWN", "low"
                item["category"] = category
                item["categoryConfidence"] = confidence
            item["categorySource"] = "ppm"
        else:
            category, confidence, reason = infer_category(item)
            item["category"], item["categoryConfidence"] = category, confidence
            item["categorySource"] = "inferred"
            item["_categoryReason"] = reason
            if item["_extraSlides"] or (item["_inlineRecipe"] and item["recipe"] is None):
                warnings.append(
                    f"[recipe slide {item['_slideNo']}] layout {item['layoutCode']!r}: recipe content and "
                    f"slide count disagree ({reason}) — confidence set to low"
                )

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

def validate_decks(specs):
    """specs: [(recipe_path, preprod_path|None)]. Opens every deck and
    matches it to a template profile BEFORE anything is extracted or
    written. Returns ([(recipe_prs, recipe_profile, preprod_prs|None)], errors).
    A deck that matches no profile is an error naming the deck and what
    didn't match — never guessed at."""
    loaded, errors = [], []
    for recipe_path, preprod_path in specs:
        try:
            rprs = Presentation(recipe_path)
            profile = detect_profile(rprs, Path(recipe_path).name, "recipe")
            pprs = None
            if preprod_path:
                pprs = Presentation(preprod_path)
                detect_profile(pprs, Path(preprod_path).name, "ppm")
            loaded.append((rprs, profile, pprs))
        except DeckFormatError as e:
            errors.append(str(e))
            loaded.append(None)
    return loaded, errors


def load_deck_pair(yymm, recipe_path, preprod_path, rprs, profile, pprs):
    """Read one Recipe & Captions deck (+ optional PPM deck) into memory.
    Writes nothing."""
    warnings = []
    mode = "ppm" if pprs is not None else "inferred"
    print(f"[{yymm}] {Path(recipe_path).name}")
    print(f"[{yymm}]   template profile: {profile.name}   category mode: {mode}"
          + ("" if pprs is not None else "   (no Pre-Prod deck)"))

    ppm, banners = None, []
    if pprs is not None:
        print(f"[{yymm}]   reading category lookup from Pre-Prod deck: {Path(preprod_path).name}")
        category_lookup, title_lookup, lookup_warnings = build_category_lookup(pprs, profile)
        warnings.extend(f"[{yymm}] {w}" for w in lookup_warnings)
        print(f"[{yymm}]   -> {len(category_lookup)} layout codes mapped to categories")
        ppm = (category_lookup, title_lookup)
        banners = scan_preprod_banners(pprs, profile)

    item_warnings = []
    items, prs = extract_items(rprs, profile, ppm, item_warnings)
    warnings.extend(f"[{yymm}] {w}" for w in item_warnings)
    for item in items:
        item["_deck"] = yymm
    print(f"[{yymm}]   -> {len(items)} items extracted")

    for line in summarize_overview(items, parse_overview(prs)):
        print(f"[{yymm}]   {line}")

    dropbox_links = find_dropbox_links(prs)
    print(f"[{yymm}]   -> {len(dropbox_links)} Dropbox link(s) found in the first "
          f"{DROPBOX_LINK_SCAN_SLIDES} slide(s)")

    return {
        "yymm": yymm,
        "items": items,
        "preprod_banners": banners,
        "dropbox_links": dropbox_links,
        "warnings": warnings,
        "mode": mode,
        "ppm": ppm,
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


# ---------------------------------------------------------------------
# Re-runs: never overwrite a category someone already settled
# ---------------------------------------------------------------------

def load_existing_rows(path):
    """Items array of an existing recipes/*.js data file, or None."""
    if not path.exists():
        return None
    text = path.read_text(encoding="utf-8")
    start = text.find("= [")
    end = text.find("\n];", start)
    if start < 0 or end < 0:
        return None
    try:
        return json.loads(text[start + 2:end + 2])
    except ValueError:
        return None


def apply_protection(items, path):
    """Keep categories already in an existing data file. Only rows whose
    categorySource is "inferred" (or that don't exist yet) may change:
      - "manual"                              -> protected
      - no categorySource (written before the field existed) -> protected as
        "ppm"-equivalent, or as "manual" if the category is outside what
        inference can produce (EAP, STOP MOTION, ...)
      - "ppm"                                 -> protected
    A protected row keeps its category/source/confidence; if this run would
    have produced a different category, the conflict is returned (and never
    applied). A legacy "UNKNOWN" row is a placeholder, not a decision, so it
    isn't protected. Returns (conflict messages, number of protected rows)."""
    rows = load_existing_rows(path)
    if not rows:
        return [], 0
    by_key = {(r.get("layoutCode"), r.get("layoutName")): r for r in rows}
    by_code = {}
    for r in rows:
        by_code.setdefault(r.get("layoutCode"), []).append(r)

    conflicts, protected = [], 0
    for it in items:
        row = by_key.get((it["layoutCode"], it["layoutName"]))
        if row is None and len(by_code.get(it["layoutCode"], [])) == 1:
            row = by_code[it["layoutCode"]][0]
        if row is None:
            continue
        old_cat, old_src = row.get("category"), row.get("categorySource")
        if old_src == "inferred" or old_cat in (None, "UNKNOWN"):
            continue
        if old_src is None:
            old_src = "ppm" if old_cat in INFERABLE_CATEGORIES + ("OUT OF PACK W/ FOOD STYLING",) else "manual"
        protected += 1
        if it["category"] != old_cat:
            conflicts.append(
                f"{path.name} {it['layoutCode']} ({it['layoutName']!r}): kept {old_cat!r} [{old_src}]; "
                f"this run gives {it['category']!r} [{it['categorySource']}, {it['categoryConfidence']}]"
            )
        it["category"] = old_cat
        it["categorySource"] = old_src
        it["categoryConfidence"] = row.get("categoryConfidence") or "high"
    return conflicts, protected


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
        conflicts, n_protected = apply_protection(file_items, out_dir / filename)
        if n_protected:
            print(f"      {n_protected} existing row(s) keep their category (manual / ppm / legacy rows are protected)")
        for c in conflicts:
            print(f"      ! CONFLICT, existing value kept: {c}")
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


_PAIR_LABELS = {
    "code": "code", "adjacent-title": "adjacency + title (tiebreaker)",
    "adjacent-nocode": "adjacency (recipe slide has no code)",
    "adjacent-codemismatch": "adjacency (code matches nothing)", None: "-",
}


def print_layout_table(items):
    print("\nLayouts:")
    print(f"  {'deck':5} {'code':6} {'layout name':40} {'category':28} {'source':9} {'conf':7} recipe slide paired by")
    for it in items:
        print(f"  {it['_deck']:5} {it['layoutCode']:6} {it['layoutName'][:39]!r:40} {it['category']:28} "
              f"{it['categorySource']:9} {it['categoryConfidence']:7} {_PAIR_LABELS[it['_pairMethod']]}")


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

    review = [it for it in items if it["categorySource"] == "inferred" and it["categoryConfidence"] != "high"]
    if review:
        print(f"\n{len(review)} inferred categorie(s) below 'high' confidence — cross-check these:")
        for it in review:
            print(f"  - [{it['_deck']}] {it['layoutCode']:6} {it['layoutName']!r:42} -> {it['category']:12} "
                  f"({it['categoryConfidence']}: {it.get('_categoryReason', '')})")


# ---------------------------------------------------------------------
# Backtest: PPM labels vs. what inference says for the same deck
# ---------------------------------------------------------------------

def ppm_truth(item, ppm):
    category_lookup, title_lookup = ppm
    return category_lookup.get(item["layoutCode"]) or title_lookup.get(normalize_title(item["layoutName"]))


def _pct(n, d):
    return f"{n}/{d} = {100 * n / d:.0f}%" if d else "n/a"


def print_backtest(rows, label):
    """rows: [{'deck','code','name','truth','pred','conf','reason'}] for layouts
    whose PPM label is known. EAP / STOP MOTION labels can't be inferred, so
    they're reported separately and never counted as inference errors."""
    print("\n" + "=" * 78)
    print(f"BACKTEST — {label}: {len(rows)} layout(s) with a PPM label")
    print("=" * 78)
    if not rows:
        return
    inferable = rows  # EAP / STOP MOTION count too: the overview's SKU tags can produce them
    skipped = []

    rec_ok = sum(1 for r in inferable if (r["truth"] == "RECIPE") == (r["pred"] == "RECIPE"))
    print(f"\nRecipe vs non-recipe accuracy: {_pct(rec_ok, len(inferable))}")

    print("\nPer subtype (PPM label -> inferred), exact category match:")
    cats = ["RECIPE", "OUT OF PACK", "OUT OF PACK W/ FOOD STYLING", "NON-FOOD", "GROUP SHOT", "EAP", "STOP MOTION"]
    for truth in cats:
        sub = [r for r in inferable if r["truth"] == truth]
        if not sub:
            continue
        ok = sum(1 for r in sub if r["pred"] == truth)
        dist = {}
        for r in sub:
            dist[r["pred"]] = dist.get(r["pred"], 0) + 1
        print(f"  {truth:30} {_pct(ok, len(sub)):12} predicted as: {dist}")
    print("\nPer predicted subtype (precision):")
    for pred in cats:
        sub = [r for r in inferable if r["pred"] == pred]
        if sub:
            print(f"  {pred:30} {_pct(sum(1 for r in sub if r['truth'] == pred), len(sub))}")

    print("\nAccuracy per confidence level (exact category):")
    for conf in ("high", "medium", "low"):
        sub = [r for r in inferable if r["conf"] == conf]
        print(f"  {conf:7} {_pct(sum(1 for r in sub if r['pred'] == r['truth']), len(sub))}")

    wrong = [r for r in rows if r["pred"] != r["truth"]]
    print(f"\nEvery disagreement ({len(wrong)}):")
    for r in wrong:
        note = "  [label not inferable]" if r in skipped else ""
        print(f"  [{r['deck']}] {r['code']:6} {r['name']!r:40} PPM={r['truth']:28} inferred={r['pred']:12} "
              f"({r['conf']}: {r['reason']}){note}")
    if skipped:
        print(f"\n{len(skipped)} layout(s) carry a PPM label inference never produces (EAP etc.); excluded from accuracy.")


def run_backtest(loaded_specs, tune_through):
    """For every deck with a Pre-Prod deck: extract with PPM labels and again
    without, compare. Decks with YYMM <= tune_through are the tuning set,
    later ones the validation set."""
    all_rows = {}
    for (recipe_path, preprod_path), loaded in loaded_specs:
        if loaded is None:
            continue
        rprs, profile, pprs = loaded
        yymm = yymm_from_filename(recipe_path) or "????"
        if pprs is None:
            print(f"[{yymm}] no Pre-Prod deck; nothing to compare against")
            continue
        print(f"[{yymm}] template profile: {profile.name}   compared: ppm vs inferred")
        lookup, titles, _ = build_category_lookup(pprs, profile)
        ppm = (lookup, titles)
        with_ppm, _ = extract_items(rprs, profile, ppm, [], with_photos=False)
        inferred, _ = extract_items(rprs, profile, None, [], with_photos=False)
        rows = []
        for a, b in zip(with_ppm, inferred):
            truth = ppm_truth(a, ppm)
            if truth is None:
                continue
            rows.append({"deck": yymm, "code": a["layoutCode"], "name": a["layoutName"], "truth": truth,
                         "pred": b["category"], "conf": b["categoryConfidence"], "reason": b["_categoryReason"]})
        all_rows[yymm] = rows
        print_backtest(rows, f"{yymm}")
    tune = [r for y, rs in all_rows.items() if tune_through and y <= tune_through for r in rs]
    valid = [r for y, rs in all_rows.items() if tune_through and y > tune_through for r in rs]
    if tune_through:
        print_backtest(tune, f"TUNING SET (decks through {tune_through})")
        print_backtest(valid, f"VALIDATION SET (decks after {tune_through})")
    else:
        print_backtest([r for rs in all_rows.values() for r in rs], "ALL DECKS")


def pair_decks(recipe_decks, preprod_decks):
    """[(recipe_path, preprod_path|None)] or None on error. PPM decks are
    matched to Recipe & Captions decks by YYMM- filename prefix when that is
    unambiguous; otherwise, with equal counts, by command-line order (the
    original behavior)."""
    if not preprod_decks:
        return [(r, None) for r in recipe_decks]
    r_months = [yymm_from_filename(r) for r in recipe_decks]
    p_months = [yymm_from_filename(p) for p in preprod_decks]
    if (None not in r_months and None not in p_months and len(set(r_months)) == len(r_months)
            and len(set(p_months)) == len(p_months) and set(p_months) <= set(r_months)):
        by_month = dict(zip(p_months, preprod_decks))
        return [(r, by_month.get(m)) for r, m in zip(recipe_decks, r_months)]
    if len(recipe_decks) == len(preprod_decks):
        return list(zip(recipe_decks, preprod_decks))
    print("Can't tell which Pre-Prod deck goes with which Recipe & Captions deck: pass the same number of "
          "each in the same order, or give every deck a unique YYMM- filename prefix.")
    return None


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--recipe-deck", action="append", required=True,
                        help="Recipe & Captions Deck .pptx (repeat for each deck)")
    parser.add_argument("--preprod-deck", action="append", default=[],
                        help="PPM / Pre-Prod Deck .pptx (optional; repeat, matched to --recipe-deck by YYMM- "
                             "prefix or order). Decks without one get categories inferred.")
    parser.add_argument("--out-dir", default="recipes", help="Output folder (default: recipes)")
    parser.add_argument("--dry-run", action="store_true",
                        help="List campaigns found and exit without prompting or writing")
    parser.add_argument("--backtest", action="store_true",
                        help="Compare inferred categories against PPM labels for every deck pair; writes nothing")
    parser.add_argument("--tune-through", metavar="YYMM",
                        help="With --backtest: decks up to this month are the tuning set, later ones the validation set")
    args = parser.parse_args()

    specs = pair_decks(args.recipe_deck, args.preprod_deck)
    if specs is None:
        return 2
    out_dir = Path(args.out_dir)

    loaded, errors = validate_decks(specs)
    if errors:
        print("Template fingerprint check FAILED — nothing was extracted or written:\n")
        print("\n\n".join(errors))
        print("\nIf this is a new/old deck format, it needs its own TemplateProfile before it can be used.")
        return 2

    if args.backtest:
        run_backtest(list(zip(specs, loaded)), args.tune_through)
        return 0

    deck_months = []
    for recipe_path, preprod_path in specs:
        yymm = resolve_deck_yymm(recipe_path, preprod_path)
        if yymm is None:
            return 1
        deck_months.append(yymm)

    pairs = [
        load_deck_pair(yymm, recipe_path, preprod_path, rprs, profile, pprs)
        for yymm, (recipe_path, preprod_path), (rprs, profile, pprs) in zip(deck_months, specs, loaded)
    ]
    items = [it for pair in pairs for it in pair["items"]]
    warnings = [w for pair in pairs for w in pair["warnings"]]

    campaigns = collect_campaigns(pairs)
    warnings.extend(assign_dropbox_urls(campaigns, pairs))
    mappable = print_dry_run(campaigns, correlation_checks(pairs))
    print_layout_table(items)
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
