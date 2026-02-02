#!/usr/bin/env python3
# ...existing code...
import csv
import re
import sys
from pathlib import Path
from difflib import SequenceMatcher
import pandas as pd

# Config
BASE = Path(__file__).resolve().parent
AREI_XLSX = BASE / 'AREI Codes_ F-codes en Vinçotte Gebruik (1).xlsx'
INPUT_CSVS = [BASE / 'extraction_results_fixed.csv', BASE / 'extraction_results_full.csv', BASE / 'extraction_results.csv']
OUTPUT_CSV = BASE / 'extraction_results_mapped.csv'

# Helpers

def norm_text(s):
    if s is None:
        return ''
    return re.sub(r"\s+"," ", str(s).strip()).lower()


def tokenize(s):
    # return words longer than 3 chars
    tokens = re.findall(r"[a-zA-Z0-9çéèàôû]+", s)
    return [t for t in tokens if len(t) > 3]


def fuzzy_score(a, b):
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


# 1) load AREI mapping
arei_items = []  # list of dicts: {'code':..., 'desc':..., 'tokens':...}
if AREI_XLSX.exists():
    try:
        # load all sheets and scan rows
        wb = pd.read_excel(AREI_XLSX, sheet_name=None, engine='openpyxl')
        for name, df in wb.items():
            # iterate rows and find columns that look like code/description
            for _, row in df.iterrows():
                # attempt to get code and description from first 3 columns
                code = None
                desc = None
                for col in row.index[:3]:
                    val = row.get(col)
                    if isinstance(val, str) and re.search(r'[A-Za-z0-9]{2,}', val):
                        # heuristic: code fields often look like 'F9' or '1306 - 1310'
                        if code is None and (re.search(r'^[A-Z0-9\-\s]+$', str(val))):
                            code = str(val).strip()
                            continue
                    # description heuristic
                # prefer column 1 as code, column 2 or 3 as desc
                try:
                    vals = list(row[:4])
                except Exception:
                    vals = list(row)
                # pick first non-null as code candidate and first long string as desc
                for v in vals[:2]:
                    if pd.notna(v):
                        if code is None and (isinstance(v, str) and len(v.strip())<40 and re.search(r'\d|[A-Z]', v)):
                            code = str(v).strip()
                # description pick third column if available else second
                if len(vals) >= 3 and pd.notna(vals[2]):
                    desc = str(vals[2]).strip()
                elif len(vals) >= 2 and pd.notna(vals[1]):
                    desc = str(vals[1]).strip()
                else:
                    # fallback: find first long text
                    for v in vals:
                        if pd.notna(v) and isinstance(v, str) and len(v.strip())>20:
                            desc = v.strip(); break

                if desc:
                    c = code if code else 'UNKNOWN'
                    d = desc
                    nd = norm_text(d)
                    arei_items.append({'code': c, 'desc': d, 'norm': nd, 'tokens': tokenize(nd)})
    except Exception as e:
        print('Error reading AREI Excel:', e, file=sys.stderr)
else:
    print('AREI excel not found at', AREI_XLSX, file=sys.stderr)

print(f'Loaded {len(arei_items)} AREI entries')

# 2) pick input CSV
input_csv = None
for c in INPUT_CSVS:
    if c.exists():
        input_csv = c
        break
if input_csv is None:
    print('No input extraction CSV found. Looked for:', INPUT_CSVS, file=sys.stderr)
    sys.exit(1)

# 3) read input
rows = []
with open(input_csv, newline='', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    for r in reader:
        rows.append(r)

print(f'Read {len(rows)} rows from {input_csv.name}')

# 4) function to split inbreuken into candidate phrases
# Use MULTILINE flag instead of inline (?m)
_split_re = re.compile(r'\n+|^\s*\d+\.|^\s*[ivxIVX]+\.|^\s*[A-Z]{2,}\s*$', re.MULTILINE)

def split_inbreuken(text):
    if not text:
        return []
    # normalize some separators
    txt = text.replace('\r','\n')
    pieces = [p.strip() for p in _split_re.split(txt) if p and len(p.strip())>10]
    # further split on semicolons and long lines
    out = []
    for p in pieces:
        for s in re.split(r'[;\\|\\u2022\\-\\*\\n]{1,}', p):
            s2 = s.strip()
            if s2 and len(s2)>6:
                out.append(s2)
    return out

# 5) matching
THRESHOLD_FUZZY = 0.55
results = []
count_with_inbreuken = 0
count_mapped = 0

for r in rows:
    inb = r.get('inbreuken') or ''
    readable = str(r.get('pdf_readable','')).strip().lower() in ('true','1','yes')
    mapped = []
    candidates = []
    if readable and inb.strip():
        count_with_inbreuken += 1
        candidates = split_inbreuken(inb)
        for cand in candidates:
            cand_norm = norm_text(cand)
            best = []
            for ai in arei_items:
                score = 0.0
                # strong substring check: any token of ai in cand
                matched_token = None
                for tok in ai['tokens']:
                    if tok and tok in cand_norm:
                        score = max(score, 0.95)
                        matched_token = tok
                        break
                # fuzzy fallback
                if score < 0.95:
                    s = fuzzy_score(cand_norm, ai['norm'])
                    if s>score:
                        score = s
                if score >= THRESHOLD_FUZZY:
                    best.append((score, ai['code'], ai['desc'], matched_token))
            # sort best
            best_sorted = sorted(best, key=lambda x: x[0], reverse=True)[:3]
            if best_sorted:
                count_mapped += 1
                mapped.append({'phrase': cand, 'matches': best_sorted})
    # flatten mapped into a string
    mapped_strs = []
    for m in mapped:
        for sc, code, desc, tok in m['matches']:
            mapped_strs.append(f"{code}|{sc:.2f}|{(tok or '')}|{desc[:80].replace(',',';')}")
    r['mapped_arei'] = ';'.join(mapped_strs)
    r['inbreuken_phrases'] = '|'.join(candidates)
    results.append(r)

# 6) write output
fieldnames = list(results[0].keys()) if results else []
with open(OUTPUT_CSV, 'w', newline='', encoding='utf-8') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames)
    writer.writeheader()
    for row in results:
        writer.writerow(row)

print('Wrote', OUTPUT_CSV)
print('Docs with inbreuken:', count_with_inbreuken, 'Mapped phrases:', count_mapped)

# show a few samples
for r in results[:10]:
    if r.get('mapped_arei'):
        print('\n', r['pdf_filename'])
        print('  inbreuken_phrases=', r.get('inbreuken_phrases'))
        print('  mapped_arei=', r.get('mapped_arei'))

print('\nDone')
