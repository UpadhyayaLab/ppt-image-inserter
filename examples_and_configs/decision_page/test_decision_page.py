"""Check a decision page before publishing, with the page's database replaced by a stand-in (sessionStorage) preloaded
from the db export (spec `db_export`), in headless Edge:
1. resolution: every item with a saved answer (applied list, or a pick or change in the export) shows as Resolved;
2. save test: on the first open decision, pick, change and reload (the item moves to Resolved with the change text),
   reopen; on deck slide 2, save a slide edit, reload, remove it; the page opens on All slides;
3. screenshots at desktop (1400 px) and phone (400 px) width, one tall shot per open decision
   (review/item_<code>.png) and one per 3 deck slides (review/all_<n>.png).
Fails (non-zero exit) on any lost save or wrong state; prints the time per step.

Usage: python test_decision_page.py <spec.yaml> <test dir>   (also run by build_decision_page.py --test-dir)
"""
import html
import json
import re
import subprocess
import sys
import time
from pathlib import Path

import yaml
from PIL import Image

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import build_decision_page as bp  # noqa: E402

EDGE = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"

STUB = r"""
(() => {
  const KEY = 'stubdb_' + location.pathname, PRE = __PRE__, ACT = __ACT__, ERRS = [];
  window.addEventListener('error', e => ERRS.push(e.message + ' (line ' + e.lineno + ')'));
  const errs = () => 'errors: ' + (ERRS.join(' | ') || 'none');
  if (!sessionStorage.getItem(KEY + '_pre')) {
    Object.keys(PRE).forEach(c => sessionStorage.setItem(KEY + '|' + c, JSON.stringify(PRE[c])));
    sessionStorage.setItem(KEY + '_pre', '1');
  }
  const load = c => JSON.parse(sessionStorage.getItem(KEY + '|' + c) || '{}');
  const subs = {};
  const snap = c => { const d = load(c); return { docs: Object.keys(d).map(id => ({ id, data: () => d[id] })), metadata: { fromCache: false } }; };
  const db = { collection: c => ({
    doc: id => ({ set: async v => { const d = load(c); d[id] = v; sessionStorage.setItem(KEY + '|' + c, JSON.stringify(d));
      setTimeout(() => (subs[c] || []).forEach(f => f(snap(c))), 0); } }),
    onSnapshot: f => { (subs[c] = subs[c] || []).push(f); setTimeout(() => f(snap(c)), 0); return () => {}; } }) };
  window.claude = { use: async n => (n === 'db' ? db : null) };
  const wait = ms => new Promise(r => setTimeout(r, ms));
  const q = s => document.querySelector(s);
  window.addEventListener('load', async () => {
    await wait(300);
    const out = document.createElement('pre'); out.id = 'log'; document.body.prepend(out);
    const log = JSON.parse(sessionStorage.getItem(KEY + '_log') || '[]'), phase = sessionStorage.getItem(KEY + '_phase') || '';
    const item = c => q('.item[data-item="' + c + '"]'), card = c => item(c).querySelector('.card');
    const st = c => (card(c).classList.contains('pick') ? 'green' : card(c).classList.contains('change') ? 'amber' : 'gray') +
      ' | card ' + (card(c).hidden ? 'hidden' : 'shown') + ' | in ' + item(c).parentNode.id + ' | line ' +
      (item(c).querySelector('.resline').hidden ? 'hidden' : item(c).querySelector('.resline').textContent.replace(/\s+/g, ' ').trim());
    const deck2 = () => q('.dslide[id="s-2"]');
    if (ACT === 'resolution') { out.textContent = 'LOG\n' + errs() + '\nresolved: ' + document.body.dataset.resolved + '\nEND'; return; }
    if (ACT && ACT.view) { window.jumpTo(ACT.view); return; }
    if (ACT && ACT.cycle && !phase) {
      const c = ACT.cycle;
      log.push('opens on: ' + (q('#view-all').hidden ? 'decisions' : 'all slides'));
      item(c).querySelectorAll('.ans')[1].click(); await wait(150); log.push('after pick: ' + st(c));
      q('#note-' + CSS.escape(c)).value = 'Test change text'; item(c).querySelector('.savechg').click(); await wait(150);
      log.push('after change: ' + st(c));
      const t = deck2().querySelector('textarea'); t.value += '\nTEST SLIDE EDIT'; t.dispatchEvent(new Event('input', { bubbles: true }));
      deck2().querySelector('.slsave').click(); await wait(150);
      log.push('before reload ' + errs());
      log.push('slide after save: '+ (deck2().classList.contains('changed') ? 'changed' : 'not changed') + ' | ' + q('#count').textContent);
      sessionStorage.setItem(KEY + '_log', JSON.stringify(log)); sessionStorage.setItem(KEY + '_phase', '1');
      await wait(200); location.reload(); return;
    }
    if (ACT && ACT.cycle) {
      const c = ACT.cycle;
      log.push('after reload: ' + st(c));
      log.push('stored decision: ' + JSON.stringify(load(ACT.coll)[c]));
      log.push('slide after reload: ' + (deck2().classList.contains('changed') ? 'changed' : 'not changed') + ' | box has edit ' +
        deck2().querySelector('textarea').value.includes('TEST SLIDE EDIT'));
      item(c).querySelector('.resbtn').click(); await wait(100); log.push('after reopen: ' + st(c));
      deck2().querySelector('.slclear').click(); await wait(150);
      log.push('slide after remove: ' + (deck2().classList.contains('changed') ? 'changed' : 'not changed') + ' | stored ' +
        JSON.stringify((load(ACT.slcoll)[deck2().dataset.slide] || {}).status));
      log.push('after reload ' + errs());
      out.textContent = 'LOG\n' + log.join('\n') + '\nEND';
    }
  });
})();
"""


def edge(src, png=None, size="1400,1000"):
    cmd = [EDGE, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=1",
           f"--window-size={size}", "--virtual-time-budget=8000"] + (["--dump-dom"] if not png else [f"--screenshot={png}"])
    r = subprocess.run(cmd + [Path(src).resolve().as_uri()], capture_output=True, timeout=180, text=True,
                       encoding="utf-8", errors="replace")
    if r.returncode:
        raise SystemExit(f"Edge failed on {src}")
    return r.stdout


def shoot(page, png, w, h):
    """Through an iframe: headless Edge will not open a window narrower than ~492 px nor paint a scrolled page."""
    frame = page.with_name(page.stem + f"_{w}_frame.html")
    frame.write_text(f'<body style="margin:0;overflow:hidden"><iframe src="{page.name}" style="width:{w}px;'
                     f'height:{h}px;border:0"></iframe></body>', encoding="utf-8")
    edge(frame, png, f"{max(w, 600)},{h}")
    Image.open(png).crop((0, 0, w, h)).save(png)
    return png


def logged(page):
    m = re.search(r"LOG\n(.*?)\nEND", html.unescape(edge(page)), re.S)
    if not m:
        raise SystemExit(f"{page}: the test produced no result")
    return m.group(1)


def run(spec_path, test_dir):
    t0 = time.perf_counter()
    spec = yaml.safe_load(Path(spec_path).read_text(encoding="utf-8"))
    exp = Path(spec["db_export"])
    pre = {c: {f.stem: json.loads(f.read_text(encoding="utf-8")) for f in (exp / c).glob("*.json")}
           for c in (spec.get("collection", "decisions"), spec.get("slide_collection", "slides"))}
    test_dir = Path(test_dir)
    (test_dir / "review").mkdir(parents=True, exist_ok=True)

    def page(name, act):
        p = test_dir / f"{name}.html"
        p.write_text(bp.build(spec_path, STUB.replace("__PRE__", json.dumps(pre, ensure_ascii=False))
                              .replace("__ACT__", json.dumps(act)))[0], encoding="utf-8")
        return p

    _, data = bp.build(spec_path)
    items = [it for it in data["items"] if "options" in it]
    saved = {c for c, d in pre[spec.get("collection", "decisions")].items()
             if (d.get("status") or ("pick" if d.get("choice") else "")) in ("pick", "change")}
    want = {it["code"] for it in items if it["resolved"] or it["code"] in saved}
    res = logged(page("test_resolution", "resolution"))
    if "errors: none" not in res:
        raise SystemExit(f"page script errors: {res}")
    got = set(res.split("resolved: ", 1)[1].split())
    print(f"resolution: {len(want)} items with a saved answer, {len(got)} shown as Resolved")
    if got != want:
        raise SystemExit(f"resolution FAILED: not resolved {sorted(want - got)}, resolved without an answer {sorted(got - want)}")
    t = time.perf_counter()
    first = next(it["code"] for it in items if not it["resolved"] and it["code"] not in saved and len(it["options"]) > 1)
    log = logged(page("test_cycle", {"cycle": first, "coll": spec.get("collection", "decisions"),
                                     "slcoll": spec.get("slide_collection", "slides")}))
    print(f"save test (decision {first}, slide 2):\n  " + log.replace("\n", "\n  "))
    expect = [r"before reload errors: none", r"after reload errors: none", r"opens on: all slides",r"after pick: green \| card shown \| in open \| line Resolved:",
              r"after change: amber \| card shown", r"slide after save: changed \| .*1 slide changed",
              r"after reload: amber \| card hidden \| in resolved \| line Resolved: Test change text \(your change\)",
              r'stored decision: .*"status":"change".*"note":"Test change text"',
              r"slide after reload: changed \| box has edit true", r"after reopen: amber \| card shown \| in resolved",
              r'slide after remove: not changed \| stored "cleared"']
    bad = [x for x in expect if not re.search(x, log)]
    if bad:
        raise SystemExit(f"save test FAILED: no match for {bad}")
    print(f"save test: PASS ({time.perf_counter() - t:.0f} s)")
    t = time.perf_counter()
    shots = []
    for name, act in (("all", None), ("decisions", {"view": "view-dec"}), ("resolved", {"view": "res-h"})):
        p = page(f"test_{name}", act)
        shots += [shoot(p, test_dir / f"test_{name}_desktop.png", 1400, 1000),
                  shoot(p, test_dir / f"test_{name}_phone.png", 400, 860)]
    for old in (test_dir / "review").glob("*"):
        old.unlink()
    for it in items:
        safe = re.sub(r"[^\w-]", "x", it["code"])
        shots.append(shoot(page(f"review/item_{safe}", {"view": "d-" + it["code"]}),
                           test_dir / "review" / f"item_{safe}.png", 1400, 2600))
    for k in range(0, len(data["deck"]), 3):
        n = data["deck"][k]["n"]
        shots.append(shoot(page(f"review/all_{n:02d}", {"view": f"s-{n}"}), test_dir / "review" / f"all_{n:02d}.png",
                           1400, 2600))
    print(f"screenshots: {len(shots)} in {test_dir} ({time.perf_counter() - t:.0f} s); total {time.perf_counter() - t0:.0f} s")


if __name__ == "__main__":
    run(sys.argv[1], sys.argv[2])
