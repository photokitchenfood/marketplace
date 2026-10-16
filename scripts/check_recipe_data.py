#!/usr/bin/env python3
"""Regression check for /recipes data. Run from repo root: python3 scripts/check_recipe_data.py
Exits 1 on any problem. Checks index ids/files, global const name collisions,
duplicate items inside/between shoots, and missing/orphan files."""
import json, re, sys, collections, pathlib
root = pathlib.Path(__file__).resolve().parent.parent
idx_txt = (root/"recipes-index.js").read_text()
index = json.loads(re.search(r"RECIPES_INDEX\s*=\s*(\[.*?\n\]);", idx_txt, re.S).group(1))
bad = 0
def report(title, rows):
    global bad
    if rows:
        bad += len(rows); print(f"\n[{title}] {len(rows)}")
        for r in rows: print("  ", r)
ids = collections.Counter(e["id"] for e in index)
report("duplicate index ids", [k for k,v in ids.items() if v>1])
files = collections.Counter(e["file"] for e in index)
report("duplicate index files", [k for k,v in files.items() if v>1])
on_disk = {f"recipes/{p.name}" for p in (root/"recipes").glob("*.js")}
report("index file missing on disk", sorted(set(files)-on_disk))
report("recipe file not in index", sorted(on_disk-set(files)))
varmap = collections.defaultdict(list)
items_by_key = collections.defaultdict(list)
for e in index:
    p = root/e["file"]
    if not p.exists(): continue
    t = p.read_text()
    m = re.search(r"^const (\w+)\s*=", t, re.M)
    if m: varmap[m.group(1)].append(e["id"])
    body = re.search(r"=\s*(\[.*?\n\]);", t, re.S)
    data = json.loads(body.group(1))
    if len(data) != e["count"]: report("count mismatch", [f'{e["id"]}: index={e["count"]} file={len(data)}'])
    seen = collections.Counter(d.get("layoutCode")+"|"+d.get("layoutName","") for d in data)
    report(f'dup layoutCode+name inside {e["id"]}', [k for k,v in seen.items() if v>1])
    for d in data:
        items_by_key[(d.get("layoutName","").strip().lower(), d.get("campaign","").strip().lower())].append(e["id"])
report("global const name collisions (2nd file fails to execute)", [f"{k}: {v}" for k,v in varmap.items() if len(v)>1])
report("same layoutName+campaign in >1 file", [f"{k}: {sorted(set(v))}" for k,v in items_by_key.items() if len(set(v))>1])
print("\nOK" if not bad else f"\n{bad} problem(s)"); sys.exit(1 if bad else 0)
