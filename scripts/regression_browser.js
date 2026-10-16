// Browser regression for the catalog's folder/filter/dup bugs. No install needed:
//   1. serve the repo (python3 -m http.server), open index.html, sign in
//   2. paste this whole file into the DevTools console (it returns a PASS/FAIL report)
(async () => {
  const wait = ms => new Promise(r => setTimeout(r, ms));
  const fails = [], log = [];
  const check = (ok, msg) => { (ok ? log : fails).push((ok ? "ok   " : "FAIL ") + msg); };
  await primeAllData();

  // 1. every shoot's loaded data is its own file's data (count matches index, uids unique, no shared objects)
  const all = recipesIndex.flatMap(e => shootDataCache[e.id] || []);
  recipesIndex.forEach(e => { if((shootDataCache[e.id]||[]).length !== e.count) fails.push(`FAIL ${e.id}: cached ${(shootDataCache[e.id]||[]).length} != index ${e.count}`); });
  check(new Set(all).size === all.length, `no shared item objects across shoots (${all.length} items)`);
  check(new Set(all.map(i => i._uid)).size === all.length, "every item has a unique _uid");

  // 2. Folders: search a term, click EVERY visible shoot leaf, rendered cards must be that shoot's own items
  const s = document.getElementById("shootTreeSearch");
  for(const term of ["Halloween", "Christmas", "IG"]){
    s.value = term; s.dispatchEvent(new Event("input"));
    for(const li of [...document.querySelectorAll("li.shootLi")].filter(l => l.style.display !== "none")){
      li.querySelector(".node").click(); await wait(150);
      const entry = recipesIndex.find(e => e.id === currentShoot.id);
      const cards = [...document.querySelectorAll("#catalogGrid .card")].map(c => c.dataset.uid);
      const own = cards.every(u => u.startsWith(entry.id + "#"));
      if(!(own && cards.length === entry.count)) fails.push(`FAIL folder "${term}" -> ${entry.id}: ${cards.length}/${entry.count} cards, own=${own}`);
    }
  }
  s.value = ""; s.dispatchEvent(new Event("input"));

  // 3. All Photos: unfiltered, filtered, cleared — rendered count == unique count == expected
  document.querySelector('.mode-btn[data-tab="shuffler"]').click(); await wait(300);
  const q = document.getElementById("shufflerSearch");
  const rendered = () => [...document.querySelectorAll("#shufflerGrid .card")].map(c => c.dataset.uid);
  const total = recipesIndex.reduce((n, e) => n + e.count, 0);
  let r = rendered(); check(r.length === total && new Set(r).size === total, `All Photos unfiltered: ${r.length} rendered / ${new Set(r).size} unique / ${total} expected`);
  for(const term of ["Gochujang", "Halloween", "chocolate"]){
    q.value = term; q.dispatchEvent(new Event("input")); r = rendered();
    const expected = all.filter(i => matchesFilters(i, term)).length;
    check(r.length === new Set(r).size && r.length === expected, `search "${term}": ${r.length} rendered, ${new Set(r).size} unique, ${expected} expected`);
  }
  document.getElementById("resetFiltersBtn").click(); await wait(200); r = rendered();
  check(r.length === total && new Set(r).size === total, `after Reset: ${r.length} rendered (expected ${total})`);

  const report = [...log, ...fails].join("\n") + `\n\n${fails.length ? "FAILED: " + fails.length : "ALL PASSED"}`;
  console.log(report); return report;
})();
