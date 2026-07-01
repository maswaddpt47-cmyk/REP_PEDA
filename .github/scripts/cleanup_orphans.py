"""
Supprime les fiches, mémos et entrées slide-plans orphelins :
un atelier est orphelin si son PPTX n'existe plus dans assets/pptx/.
"""
import json, os
from pathlib import Path

ROOT = Path(__file__).parent.parent.parent
PPTX_DIR  = ROOT / "assets" / "pptx"
FICHES_DIR = ROOT / "assets" / "fiches"
MEMOS_DIR  = ROOT / "assets" / "memos"
PLANS_FILE = ROOT / "assets" / "slide-plans.json"

existing = {p.stem for p in PPTX_DIR.glob("*.pptx")}

removed = []

# Fiches orphelines
for f in sorted(FICHES_DIR.glob("*.pptx")):
    key = f.stem.split("-fiche-")[0]
    if key not in existing:
        f.unlink()
        removed.append(str(f.relative_to(ROOT)))

# Mémos orphelins
for f in sorted(MEMOS_DIR.glob("*.pptx")):
    key = f.stem.replace("-memo", "")
    if key not in existing:
        f.unlink()
        removed.append(str(f.relative_to(ROOT)))

# slide-plans.json
with open(PLANS_FILE) as fh:
    plans = json.load(fh)

before = set(plans.keys())
plans = {k: v for k, v in plans.items() if k in existing}
after = set(plans.keys())
orphan_keys = before - after

if orphan_keys:
    with open(PLANS_FILE, "w") as fh:
        json.dump(plans, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    for k in sorted(orphan_keys):
        removed.append(f"slide-plans.json/{k}")

if removed:
    print(f"{len(removed)} artefact(s) orphelin(s) supprimé(s) :")
    for r in removed:
        print(f"  - {r}")
else:
    print("Aucun orphelin détecté.")
