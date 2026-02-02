#!/usr/bin/env python3
"""
Generate template metadata (template.json) for each cluster directory created by `run_cluster.py`.
This script reads `templates/clusters_mapping.csv`, computes pHash and logo centers for each cluster's rep
and writes `templates/cluster_XXX/template.json` containing:
  - rep: filename
  - logo_signature: {image_count, centers}
  - logo_phash: hex string or null
  - cluster_size

Usage:
  python make_templates.py --templates-dir templates

This is intended to be run from PyCharm as a main.
"""
import argparse
from pathlib import Path
import json
import sys

# ensure this script's directory is on sys.path so local imports work when run from PyCharm
base_dir = Path(__file__).resolve().parent
if str(base_dir) not in sys.path:
    sys.path.insert(0, str(base_dir))

import cluster_pdfs_by_template as cpt


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--templates-dir', default='templates', help='templates dir with cluster_XXX subfolders')
    ap.add_argument('--pdf-dir', default='pdfs', help='pdfs dir (relative to script)')
    args = ap.parse_args()

    base = Path(__file__).resolve().parent
    templates_dir = (base / args.templates_dir).resolve()
    pdf_dir = (base / args.pdf_dir).resolve()

    if not templates_dir.exists():
        print('templates-dir does not exist:', templates_dir)
        return 1

    mapping_csv = templates_dir / 'clusters_mapping.csv'
    if not mapping_csv.exists():
        print('Missing mapping CSV:', mapping_csv)
        return 2

    # gather cluster reps from mapping (first rep occurrence per cluster)
    reps = {}
    with mapping_csv.open(encoding='utf-8') as mf:
        header = mf.readline()
        for line in mf:
            parts = [x.strip() for x in line.split(',')[:4]]
            if len(parts) < 3:
                continue
            pdfname, cid, rep, size = parts
            if cid not in reps:
                reps[cid] = rep

    for cid, rep in reps.items():
        cluster_dir = templates_dir / cid
        if not cluster_dir.exists():
            cluster_dir.mkdir(parents=True, exist_ok=True)
        rep_path = pdf_dir / rep
        if not rep_path.exists():
            print('Warning: rep not found', rep_path)
            continue
        logo_sig = cpt.extract_first_page_logo_signature(rep_path)
        logo_phash = cpt.compute_first_page_phash(rep_path)
        # compute cluster size by counting files in cluster dir
        cluster_size = sum(1 for _ in cluster_dir.glob('*.pdf'))
        meta = {
            'rep': rep,
            'logo_signature': logo_sig,
            'logo_phash': logo_phash,
            'cluster_size': cluster_size
        }
        with (cluster_dir / 'template.json').open('w', encoding='utf-8') as fh:
            json.dump(meta, fh, ensure_ascii=False, indent=2)
        print('Wrote', cluster_dir / 'template.json')

    return 0


if __name__ == '__main__':
    raise SystemExit(main())
