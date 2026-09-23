#!/usr/bin/env python3
"""
Backfill the `brand` field into already-extracted recipes/*.js files
without touching any other field — including manual edits made later via
the Recipes Catalog's own editor (e.g. a fixed typo in a procedure step).

extract_recipe_deck.py's normal run fully re-derives every field from the
source decks, which would silently overwrite any such manual edit. This
script instead: re-reads only the Caption/Procedure text boxes from the
source Recipe & Captions deck to compute `brand` per layoutCode (see
extract_recipe_deck.py's extract_brands()), then patches just the `brand`
key onto the matching item already sitting in the target JS file(s).
Idempotent — safe to re-run.

Usage:
    python3 scripts/backfill_brand.py \\
        --recipe-deck "pptx/2609-RecipeCaptions_TMP Sep 2026 (Sept IG, Shop & Collect).pptx" \\
        --js-files recipes/260904-September-IG.js recipes/260904-Shop-Collect-Photos.js recipes/260904-Shop-Collect-Videos.js
"""

import argparse
import json
import re
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from extract_recipe_deck import extract_items, print_brand_summary, write_js_file  # noqa: E402

ITEM_KEY_ORDER = [
    "layoutCode", "layoutName", "category", "campaign",
    "caption", "brand", "recipe", "photo", "needsPhotoSwap",
]


def reorder_item(item):
    ordered = {k: item[k] for k in ITEM_KEY_ORDER if k in item}
    for k, v in item.items():
        if k not in ordered:
            ordered[k] = v
    return ordered


def load_js_items(path):
    text = Path(path).read_text(encoding="utf-8")
    data_m = re.search(r"const \w+_DATA = (\[[\s\S]*\]);", text)
    if not data_m:
        raise ValueError(f"could not parse data array out of {path}")
    label_m = re.search(r"^// Shoot:\s*(.+)$", text, flags=re.MULTILINE)
    shoot_label = label_m.group(1).strip() if label_m else None
    return json.loads(data_m.group(1)), shoot_label


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--recipe-deck", required=True, help="Path to that month's Recipe & Captions Deck .pptx")
    parser.add_argument("--js-files", nargs="+", required=True, help="recipes/*.js files produced from that deck")
    args = parser.parse_args()

    print(f"Reading brand data from Recipe & Captions deck: {args.recipe_deck}")
    fresh_items, _prs = extract_items(args.recipe_deck, {}, [])
    brand_by_code = {}
    for it in fresh_items:
        code = it["layoutCode"]
        if code in brand_by_code:
            print(f"  warning: duplicate layoutCode {code!r} in deck; keeping first brand match")
            continue
        brand_by_code[code] = it["brand"]
    print(f"  -> brand data ready for {len(brand_by_code)} layout code(s)")

    all_items_for_summary = []
    for js_path in args.js_files:
        items, shoot_label = load_js_items(js_path)
        updated = 0
        missing = []
        for item in items:
            code = item.get("layoutCode")
            if code in brand_by_code:
                item["brand"] = brand_by_code[code]
                updated += 1
            else:
                item.setdefault("brand", [])
                missing.append(code)
            all_items_for_summary.append(item)

        items = [reorder_item(it) for it in items]
        write_js_file(Path(js_path), shoot_label or "Unsorted", items, date.today())
        note = f" (no deck match for: {missing})" if missing else ""
        print(f"  patched {js_path}: {updated}/{len(items)} item(s) matched by layoutCode{note}")

    print_brand_summary(all_items_for_summary)


if __name__ == "__main__":
    main()
