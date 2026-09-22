#!/usr/bin/env python3
"""
Extract Recipe & Captions Deck exports into per-shoot-type JS data files
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

Usage:
    python3 extract_recipe_deck.py \\
        --recipe-deck "TMP Sep 2026 Recipe & Captions Deck (Sept IG, Shop & Collect).pptx" \\
        --preprod-deck "TMP Sep 2026 Pre-Prod Deck (Sep IG, Shop & Collect).pptx" \\
        --out-dir campaigns
"""

import argparse
import base64
import io
import json
import re
import sys
from datetime import date, datetime

from pptx import Presentation
from pptx.util import Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from PIL import Image

THUMB_MAX_DIM = 500
THUMB_JPEG_QUALITY = 85

KNOWN_CATEGORIES = [
    "RECIPE",
    "OUT OF PACK W/ FOOD STYLING",  # check before "OUT OF PACK" (substring)
    "OUT OF PACK",
    "NON-FOOD",
    "GROUP SHOT",
]

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


def clean_multiline(text):
    # collapse Google Slides' \x0b soft line breaks to real newlines
    return text.replace("\x0b", "\n").strip()


# ---------------------------------------------------------------------
# Pre-Prod deck: layout code -> category lookup
# ---------------------------------------------------------------------

def build_category_lookup(preprod_path):
    prs = Presentation(preprod_path)
    lookup = {}
    warnings = []

    for i, slide in enumerate(prs.slides):
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

        if layout_code in lookup and lookup[layout_code] != category:
            warnings.append(
                f"[preprod slide {i + 1}] layout_code={layout_code!r} "
                f"category conflict: {lookup[layout_code]!r} vs {category!r} (keeping first)"
            )
            continue

        lookup[layout_code] = category

    return lookup, warnings


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


def extract_deck_shoot_date(prs):
    """Best-effort parse of 'Shoot Date/s: Sept 4, 2026' off the cover slide.

    Some months' Recipe & Captions deck cover slide is rewritten and drops
    the Shoot Date line (it only survives when the cover was copy-pasted
    from the Pre-Prod deck) — callers should fall back to the Pre-Prod
    deck's cover slide when this returns None.
    """
    if len(prs.slides) == 0:
        return None
    for shape in prs.slides[0].shapes:
        text = shape_text(shape)
        m = re.search(r"Shoot Date/s:\s*([A-Za-z]+)\.?\s+(\d{1,2}),?\s+(\d{4})", text)
        if m:
            month_str, day_str, year_str = m.groups()
            # normalize informal abbreviations ("Sept") that don't match
            # strptime's %b/%B before falling back to trying both
            month_str_norm = month_str[:3]
            raw = f"{month_str_norm} {day_str} {year_str}"
            for fmt in ("%b %d %Y", "%B %d %Y"):
                try:
                    return datetime.strptime(raw, fmt).date()
                except ValueError:
                    continue
            for fmt in ("%b %d %Y", "%B %d %Y"):
                try:
                    return datetime.strptime(f"{month_str} {day_str} {year_str}", fmt).date()
                except ValueError:
                    continue
    return None


def extract_items(recipe_path, category_lookup, warnings):
    prs = Presentation(recipe_path)
    slides = list(prs.slides)
    items = []
    current_shoot_type = None

    i = 0
    while i < len(slides):
        slide = slides[i]

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
        campaign = shape_text(banner_shapes[0]).strip() if banner_shapes else ""
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
            "campaign": campaign,
            "shootType": current_shoot_type,
            "caption": caption,
            "recipe": None,
            "photo": {
                "filename": f"{layout_code}.jpg",
                "dropbox_path": "",
                "image_data": photo_b64,
            },
            "needsPhotoSwap": True,
        }

        paired_next = i + 1 < len(slides) and slide_has_recipe_text(slides[i + 1])

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
            i += 2
        else:
            category = category_lookup.get(layout_code)
            if category is None:
                warnings.append(
                    f"[recipe slide {i + 1}] layout_code={layout_code!r} ({layout_name!r}): "
                    f"no matching entry in Pre-Prod deck category lookup; "
                    f"leaving category as 'UNKNOWN'"
                )
                category = "UNKNOWN"
            item["category"] = category
            i += 1

        items.append(item)

    return items, prs


# ---------------------------------------------------------------------
# Output: one JS file per shoot type, mirroring campaigns/ conventions
# ---------------------------------------------------------------------

def slugify(text):
    text = re.sub(r"[()/]", "", text)
    text = re.sub(r"[^A-Za-z0-9]+", "-", text).strip("-")
    return text


def write_js_file(out_path, shoot_type_label, items, generated_date):
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
        payload = {k: v for k, v in item.items() if k != "shootType"}
        block = json.dumps(payload, indent=2, ensure_ascii=False)
        blocks.append("\n".join("  " + line for line in block.split("\n")))
    lines.append(",\n".join(blocks))
    lines.append("];")
    lines.append("")
    lines.append("if (typeof window !== 'undefined') window.__recipesData = " + var_name + ";")
    out_path.write_text("\n".join(lines), encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--recipe-deck", required=True, help="Path to the Recipe & Captions Deck .pptx")
    parser.add_argument("--preprod-deck", required=True, help="Path to the Pre-Prod Deck .pptx (category lookup only)")
    parser.add_argument("--out-dir", default="campaigns", help="Output folder (default: campaigns)")
    args = parser.parse_args()

    from pathlib import Path
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    warnings = []

    print(f"Reading category lookup from Pre-Prod deck: {args.preprod_deck}")
    category_lookup, lookup_warnings = build_category_lookup(args.preprod_deck)
    warnings.extend(lookup_warnings)
    print(f"  -> {len(category_lookup)} layout codes mapped to categories")

    print(f"Reading items from Recipe & Captions deck: {args.recipe_deck}")
    items, prs = extract_items(args.recipe_deck, category_lookup, warnings)
    print(f"  -> {len(items)} items extracted")

    shoot_date = extract_deck_shoot_date(prs)
    if shoot_date is None:
        shoot_date = extract_deck_shoot_date(Presentation(args.preprod_deck))
    if shoot_date is None:
        shoot_date = date.today()
    date_prefix = shoot_date.strftime("%y%m%d")

    groups = {}
    for item in items:
        groups.setdefault(item["shootType"] or "Unsorted", []).append(item)

    print()
    for shoot_type, group_items in groups.items():
        slug = slugify(shoot_type)
        filename = f"{date_prefix}-{slug}.js"
        out_path = out_dir / filename
        write_js_file(out_path, shoot_type, group_items, date.today())
        recipe_count = sum(1 for it in group_items if it["category"] == "RECIPE")
        other_count = len(group_items) - recipe_count
        print(f"  wrote {out_path}  ({len(group_items)} items: {recipe_count} RECIPE, {other_count} other)")

    if warnings:
        print(f"\n{len(warnings)} warning(s):")
        for w in warnings:
            print(f"  - {w}")

    unknown = [it for it in items if it["category"] == "UNKNOWN"]
    if unknown:
        print(f"\n{len(unknown)} item(s) left with category=UNKNOWN (no Pre-Prod match) — needs manual fix:")
        for it in unknown:
            print(f"  - {it['layoutCode']} ({it['layoutName']!r})")


if __name__ == "__main__":
    sys.exit(main())
