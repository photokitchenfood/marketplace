#!/usr/bin/env python3
"""
One-off: apply the Group Shot naming rule to already-extracted recipes/*.js.
Edits only the layoutName (and, for 3 April IG layouts, the category) text in
place, so photos and manual edits are untouched. Writes
scripts/group-shot-rename-log.json (old -> new) and can revert with --revert.

    python3 scripts/apply_group_shot_renames.py [--revert]
"""
import json, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LOG = ROOT / "scripts" / "group-shot-rename-log.json"
# (file, code, old name, old category, new name, new category)
PLAN = ROOT / "scripts" / "group-shot-rename-plan.json"


def entry_head(code, name, cat):
    return f'"layoutCode": "{code}",\n    "layoutName": {json.dumps(name, ensure_ascii=False)},\n    "category": "{cat}"'


def run(revert):
    plan = json.loads(LOG.read_text() if revert else PLAN.read_text())
    done = []
    for r in plan:
        path = ROOT / "recipes" / f"{r['file']}.js"
        text = path.read_text(encoding="utf-8")
        a = entry_head(r["code"], r["old"], r["oldCategory"])
        b = entry_head(r["code"], r["new"], r["newCategory"])
        if revert:
            a, b = b, a
        if text.count(a) != 1:
            raise SystemExit(f"{r['file']} {r['code']}: expected exactly 1 match, found {text.count(a)}")
        path.write_text(text.replace(a, b), encoding="utf-8")
        done.append(r)
    if not revert:
        LOG.write_text(json.dumps(done, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"{'reverted' if revert else 'applied'} {len(done)} layout(s)")


run("--revert" in sys.argv)
