#!/usr/bin/env python3
"""
Regenerate recipes-index.js by scanning recipes/*.js, mirroring how
campaigns-index.js indexes campaigns/*.js.

Each recipes/*.js file (written by extract_recipe_deck.py) starts with:
    // Shoot: <label>
    // Generated: <date>
    // Items: <count>

This script reads those header comments plus the YYMMDD filename prefix
to build recipes-index.js entries of the form:
    { id, label, type, year, month, file, count }

"type" is inferred as "video" when the shoot label mentions video,
otherwise "photo" — the only medium signal available at the shoot level.

Usage:
    python3 generate_recipes_index.py [--recipes-dir recipes] [--out recipes-index.js]
"""

import argparse
import re
from datetime import date
from pathlib import Path


def parse_shoot_file(path):
    text = path.read_text(encoding="utf-8")
    label_m = re.search(r"^// Shoot:\s*(.+)$", text, flags=re.MULTILINE)
    items_m = re.search(r"^// Items:\s*(\d+)$", text, flags=re.MULTILINE)
    date_m = re.match(r"^(\d{2})(\d{2})(\d{2})-", path.name)

    if not (label_m and items_m and date_m):
        return None

    label = label_m.group(1).strip()
    count = int(items_m.group(1))
    yy, mm, _dd = date_m.groups()
    year = 2000 + int(yy)
    month = int(mm)

    type_ = "video" if "video" in label.lower() else "photo"

    return {
        "id": path.stem,
        "label": label,
        "type": type_,
        "year": year,
        "month": month,
        "file": f"recipes/{path.name}",
        "count": count,
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--recipes-dir", default="recipes")
    parser.add_argument("--out", default="recipes-index.js")
    args = parser.parse_args()

    recipes_dir = Path(args.recipes_dir)
    entries = []
    skipped = []
    for path in sorted(recipes_dir.glob("*.js")):
        entry = parse_shoot_file(path)
        if entry is None:
            skipped.append(path.name)
            continue
        entries.append(entry)

    # newest shoot first, mirroring campaigns-index.js ordering
    entries.sort(key=lambda e: e["id"], reverse=True)

    lines = [
        "// PhotoKitchen Recipes Index",
        "// Auto-generated — do not edit manually",
        f"// Last updated: {date.today().isoformat()}",
        "",
        "const RECIPES_INDEX = [",
    ]
    blocks = []
    for entry in entries:
        block_lines = ["  {"]
        for i, (k, v) in enumerate(entry.items()):
            comma = "," if i < len(entry) - 1 else ""
            v_repr = f'"{v}"' if isinstance(v, str) else str(v)
            block_lines.append(f'    "{k}": {v_repr}{comma}')
        block_lines.append("  }")
        blocks.append("\n".join(block_lines))
    lines.append(",\n".join(blocks))
    lines.append("];")
    lines.append("")
    lines.append("if (typeof window !== 'undefined') window.__recipesIndex = RECIPES_INDEX;")

    Path(args.out).write_text("\n".join(lines), encoding="utf-8")

    print(f"Wrote {args.out} ({len(entries)} shoot(s))")
    for e in entries:
        print(f"  - {e['id']}: {e['label']} ({e['type']}, {e['count']} items)")
    if skipped:
        print(f"\nSkipped {len(skipped)} file(s) missing header comments: {skipped}")


if __name__ == "__main__":
    main()
