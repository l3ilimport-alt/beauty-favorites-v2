#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
סנכרון תוכן: קבצי הידע (knowledge/) → מסד הנתונים (products) — שלב חובה אחרי כל העשרה.

דוחף שדות תוכן בלבד דרך RPC מאובטח (sync_product_content):
  description (חדש מנצח) · brand · category · main_category · name_he (רק לכרטיס דק) · image_url (רק אם ריק)
לעולם לא נוגע ב: מחירים, מלאי, sku, name_en, active.

שימוש:
  python3 catalog/sync_content_to_db.py                  # כל המוצרים שבידע
  python3 catalog/sync_content_to_db.py 0123... 0456...  # ברקודים מסוימים בלבד
  python3 catalog/sync_content_to_db.py --force-names 0123...  # גם דריסת שם עברי לברקודים אלה

דרישות: backoffice/.sync-env עם CONTENT_SYNC_SECRET (מקומי, לא בריפו).
"""
import json, os, re, sys, glob, urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
KNOW = os.path.join(ROOT, "knowledge")
CATALOG = os.path.join(ROOT, "catalog")

def load_env():
    env = {}
    path = os.path.join(ROOT, "backoffice", ".sync-env")
    for line in open(path, encoding="utf-8"):
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip()
    cfg = json.load(open(os.path.join(CATALOG, "supabase.config.json"), encoding="utf-8"))
    env["URL"] = cfg.get("url") or cfg.get("SUPABASE_URL")
    env["ANON"] = cfg.get("anonKey") or cfg.get("anon_key") or cfg.get("SUPABASE_ANON_KEY")
    return env

def digits(v):
    return re.sub(r"\D", "", str(v or ""))

# מיפוי קטגוריה ראשית — אותה לוגיקה כמו build_catalog.ptype (ייבוא ישיר)
sys.path.insert(0, CATALOG)
from build_catalog import ptype, norm_brand, brand_from_name  # noqa: E402

def collect_rows(only_barcodes=None, force_names=None):
    rows, seen = [], set()
    for pj in glob.glob(os.path.join(KNOW, "*", "product.json")):
        try:
            p = json.load(open(pj, encoding="utf-8"))
        except Exception as e:
            print("skip", pj, e); continue
        bc = digits(p.get("barcode"))
        if not bc or bc in seen:
            continue
        if only_barcodes and bc not in only_barcodes:
            continue
        seen.add(bc)
        brand = norm_brand(p.get("brand"))
        if brand == "אחר":
            brand = brand_from_name(p.get("name_he") or "")
        row = {
            "barcode": bc,
            "description": (p.get("description") or "").strip(),
            "brand": ("" if brand == "אחר" else brand).strip(),
            "category": (p.get("category_refined") or "").strip(),
            "main_category": ptype(p),
            "name_he": (p.get("name_he") or "").strip(),
        }
        if force_names and bc in force_names:
            row["force_name"] = "1"
        rows.append(row)
    return rows

def push(env, rows, batch=200):
    total_updated = total_missing = 0
    for i in range(0, len(rows), batch):
        part = rows[i:i + batch]
        body = json.dumps({"p_secret": env["CONTENT_SYNC_SECRET"], "p_rows": part}).encode()
        req = urllib.request.Request(
            env["URL"] + "/rest/v1/rpc/sync_product_content", data=body, method="POST",
            headers={"apikey": env["ANON"], "Authorization": "Bearer " + env["ANON"],
                     "Content-Type": "application/json"})
        res = json.load(urllib.request.urlopen(req, timeout=60))
        total_updated += res.get("updated", 0)
        total_missing += res.get("missing", 0)
        print(f"  אצווה {i//batch+1}: עודכנו {res.get('updated',0)}, לא נמצאו במסד {res.get('missing',0)}")
    return total_updated, total_missing

def main():
    args = [a for a in sys.argv[1:]]
    force = set()
    if "--force-names" in args:
        idx = args.index("--force-names")
        force = set(digits(a) for a in args[idx + 1:])
        args = args[:idx]
    only = set(digits(a) for a in args) or None
    if force and not only:
        only = set(force)
    env = load_env()
    rows = collect_rows(only, force)
    if not rows:
        print("אין שורות לדחיפה."); return
    print(f"דוחף תוכן ל-{len(rows)} מוצרים…")
    updated, missing = push(env, rows)
    print(f"✅ סונכרן: {updated} עודכנו · {missing} ברקודים לא קיימים במסד (תקין למוצרים שטרם נקלטו)")

if __name__ == "__main__":
    main()
