#!/usr/bin/env python3
"""
Ingest a YYMM-prefixed photo folder (no PPM / Recipe & Captions deck) into
recipes/<YYMM-Campaign-Name>.js, in the same shape as extract_recipe_deck.py.

Per-photo inference (category, layout name, brands, needsReview) is made by
looking at the photos and stored in scripts/ingest-plans/<YYMM>.json, one entry
per photo in card order. This script resizes the photos, embeds them as base64
(like every other shoot) and writes the data file. Originals are never copied.

Layout names ending in " *" are best guesses for recipe photos (title unknown).
The Dropbox link comes only from dropbox-links.json (folder prefix -> URL).

    python3 scripts/ingest_photo_folder.py "pptx/2001 January IG" [--dry-run]
    then: python3 scripts/generate_recipes_index.py
"""
import argparse, base64, json, re, subprocess, sys, tempfile
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("folder")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--width", type=int, default=1200)
    a = ap.parse_args()

    folder = Path(a.folder)
    m = re.match(r"^(\d{4})\s+(.+)$", folder.name)
    if not m:
        sys.exit("folder name must be 'YYMM Campaign Name'")
    prefix, campaign = m.group(1), m.group(2).strip()
    slug = re.sub(r"[^A-Za-z0-9]+", "-", campaign).strip("-")
    out = ROOT / "recipes" / f"{prefix}-{slug}.js"
    code = "IG" if re.search(r"\bIG\b", campaign) else re.sub(r"[^A-Z]", "", campaign.upper())[:2] or "P"
    varname = "_".join(slug.upper().split("-")) + "_DATA"

    plan = json.loads((ROOT / "scripts" / "ingest-plans" / f"{prefix}.json").read_text(encoding="utf-8"))
    links = json.loads((ROOT / "dropbox-links.json").read_text(encoding="utf-8"))
    dropbox = links.get(prefix, "")

    items = []
    with tempfile.TemporaryDirectory() as tmp:
        for i, p in enumerate(plan, 1):
            src = folder / p["src"]
            if not src.exists():
                sys.exit(f"missing photo: {src}")
            fname = f"{code}{i}.jpg"
            dst = Path(tmp) / fname
            subprocess.run(["sips", "-s", "format", "jpeg", "-s", "formatOptions", "80",
                            "-Z", str(a.width), str(src), "--out", str(dst)],
                           check=True, capture_output=True)
            data = dst.read_bytes()
            items.append({
                "layoutCode": f"{code}{i}",
                "layoutName": p["layoutName"],
                "category": p["category"],
                "categorySource": "inferred",
                "categoryConfidence": "low" if p["needsReview"] else "medium",
                "campaign": campaign,
                "caption": "",
                "brand": p["brand"],
                "recipe": None,
                "photo": {"filename": fname, "dropbox_path": "", "image_data": base64.b64encode(data).decode()},
                "needsPhotoSwap": False,
                "needsReview": p["needsReview"],
                "_bytes": len(data),
            })

    print(f"{out.relative_to(ROOT)}  ({len(items)} cards)")
    for it in items:
        print(f"  {it['layoutCode']:4} {it['category']:30} review={str(it['needsReview']):5} {it['photo']['filename']} ({it['_bytes']//1024} KB)  {it['layoutName']}")
        del it["_bytes"]
    flagged = sum(i["needsReview"] for i in items)
    print(f"summary: {len(items)} cards, {flagged} needsReview, dropbox link: {'set' if dropbox else 'MISSING for ' + prefix}")
    if a.dry_run:
        return

    header = (f"// Shoot: {campaign}\n// Generated: {date.today().isoformat()}\n// Items: {len(items)}\n"
              f"// Template profile: photo folder ingest (no deck; recipe/caption empty)\n\n")
    body = f"const {varname} = {json.dumps(items, indent=2, ensure_ascii=False)};\n"
    if dropbox:
        body += f"{varname}.dropboxUrl = {json.dumps(dropbox)};\n"
    body += f"\nif (typeof window !== 'undefined') window.__recipesData = {varname};\n"
    out.write_text(header + body, encoding="utf-8")


if __name__ == "__main__":
    main()
