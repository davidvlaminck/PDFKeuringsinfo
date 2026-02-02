#!/usr/bin/env python3
# Generates a per-PDF review report from extraction_results_mapped.csv
import csv
import argparse
from pathlib import Path

BASE = Path(__file__).resolve().parent
DEFAULT_IN = BASE / 'extraction_results_mapped.csv'
DEFAULT_MD = BASE / 'inbreuken_review.md'
DEFAULT_SUM = BASE / 'inbreuken_review_summary.csv'

def parse_mapped_field(mapped_field):
    # mapped_field is a ';' separated list of entries of the form code|score|token|desc
    out = []
    if not mapped_field:
        return out
    parts = [p for p in mapped_field.split(';') if p.strip()]
    for p in parts:
        items = p.split('|')
        # tolerate missing parts
        code = items[0].strip() if len(items)>0 else ''
        score = items[1].strip() if len(items)>1 else ''
        token = items[2].strip() if len(items)>2 else ''
        desc = items[3].strip() if len(items)>3 else ''
        try:
            scoref = float(score) if score else 0.0
        except Exception:
            # fallback if code and score were run together (e.g. '2100 - 21500.95...')
            # attempt to extract trailing float
            import re
            m = re.search(r"([0-9]+\.[0-9]{2})", p)
            scoref = float(m.group(1)) if m else 0.0
        out.append({'code': code, 'score': scoref, 'token': token, 'desc': desc, 'raw': p})
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', '-i', default=str(DEFAULT_IN))
    ap.add_argument('--out-md', default=str(DEFAULT_MD))
    ap.add_argument('--out-csv', default=str(DEFAULT_SUM))
    ap.add_argument('--min-score', type=float, default=0.0)
    ap.add_argument('--show-first', type=int, default=20, help='print the first N lines of the markdown report after generation')
    args = ap.parse_args()

    inp = Path(args.input)
    if not inp.exists():
        print('Input not found:', inp)
        return 1

    rows = []
    with inp.open(newline='', encoding='utf-8') as f:
        r = csv.DictReader(f)
        for row in r:
            rows.append(row)

    md_lines = []
    summary_rows = []
    total_docs = len(rows)
    docs_with_inb = 0
    total_candidates = 0
    total_matches = 0

    md_lines.append('# Inbreuken Review Report')
    md_lines.append(f'Generated from `{inp.name}` — docs: {total_docs}\n')

    for row in rows:
        filename = row.get('pdf_filename') or row.get('pdf_name') or 'UNKNOWN'
        inb_phrases = (row.get('inbreuken_phrases') or '').split('|') if row.get('inbreuken_phrases') else []
        mapped = parse_mapped_field(row.get('mapped_arei') or '')
        md_lines.append(f'\n## {filename}\n')
        if not inb_phrases:
            md_lines.append('*No extracted inbreuken detected.*\n')
            summary_rows.append({'pdf': filename, 'candidates': 0, 'matches': 0})
            continue
        docs_with_inb += 1
        md_lines.append('### Extracted phrases:')
        for i,ph in enumerate(inb_phrases, start=1):
            phs = ph.strip()
            if not phs:
                continue
            md_lines.append(f'- Phrase {i}: {phs}')
            # find mapped entries for this phrase
            # the original mapping may not associate phrase->matches explicitly; try to filter mapped by raw occurrence
            # We'll show all mapped entries but only keep those above min-score
            total_candidates += 1
            # mapped entries are global; show those where phrase appears in raw or show top matches
            # better: show matches whose raw contains the phrase or show top 3 overall
            related = [m for m in mapped if m['score'] >= args.min_score]
            if related:
                md_lines.append('  - Suggested AREI matches:')
                for m in sorted(related, key=lambda x: x['score'], reverse=True)[:6]:
                    md_lines.append(f"    - {m['code']} (score={m['score']:.2f}) token='{m['token']}' — {m['desc']}")
                    total_matches += 1
            else:
                md_lines.append('  - (no matches above min-score)')
        summary_rows.append({'pdf': filename, 'candidates': len(inb_phrases), 'matches': len([m for m in mapped if m['score']>=args.min_score])})

    # write markdown
    out_md = Path(args.out_md)
    out_md.write_text('\n'.join(md_lines), encoding='utf-8')

    # write summary csv
    out_csv = Path(args.out_csv)
    with out_csv.open('w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=['pdf','candidates','matches'])
        w.writeheader()
        for s in summary_rows:
            w.writerow(s)

    print(f'Wrote markdown review: {out_md}')
    print(f'Wrote summary CSV: {out_csv}')
    print(f'Docs with inbreuken: {docs_with_inb}, total candidate phrases: {total_candidates}, total matches (>= {args.min_score}): {total_matches}')

    # show head of markdown if requested
    if args.show_first>0:
        txt = out_md.read_text(encoding='utf-8')
        lines = txt.splitlines()
        print('\n--- Review head ---\n')
        for line in lines[:args.show_first]:
            print(line)

    return 0

if __name__=='__main__':
    import sys
    sys.exit(main())
