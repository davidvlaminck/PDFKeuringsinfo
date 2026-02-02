#!/usr/bin/env python3
"""\nRun clustering of PDFs into template-like clusters.
This is a lightweight runnable main intended to be executed from PyCharm (set working dir to project root or this folder).

Creates `templates/cluster_XXX` folders with symlinks and writes `templates/clusters_mapping.csv`.

Usage:
  python run_cluster.py
  python run_cluster.py --pdf-dir pdfs --out-dir templates --min-sim 0.25

Defaults are tuned to the existing project layout.
"""
import argparse
from pathlib import Path
import csv
import shutil
import sys

# ensure this script's directory is on sys.path so local imports work when run from PyCharm
base_dir = Path(__file__).resolve().parent
if str(base_dir) not in sys.path:
    sys.path.insert(0, str(base_dir))

# import helpers from the clustering module in this folder
import cluster_pdfs_by_template as cpt


def build_shingles_map(pdf_files, shingle_size, max_pages=None):
    shingles_map = {}
    metadata = {}
    for p in pdf_files:
        txt = cpt.extract_pdf_text(p, max_pages=max_pages)
        nt = cpt.norm_text(txt)
        toks = cpt.tokens_from_text(nt)
        sh = cpt.shingles(toks, k=shingle_size)
        shingles_map[p] = sh
        metadata[p] = {'text_len': len(nt), 'tokens': len(toks)}
    return shingles_map, metadata


def create_clusters_and_mapping(clusters, out_dir, symlinks=True, no_symlinks=False):
    out_dir.mkdir(parents=True, exist_ok=True)
    mapping_csv = out_dir / 'clusters_mapping.csv'
    with mapping_csv.open('w', newline='', encoding='utf-8') as mf:
        writer = csv.writer(mf)
        writer.writerow(['pdf_filename', 'cluster_id', 'rep_pdf', 'cluster_size'])
        for i, cluster in enumerate(clusters, start=1):
            cid = f'cluster_{i:03d}'
            cluster_dir = out_dir / cid
            cluster_dir.mkdir(exist_ok=True)
            rep = cluster[0]
            for p in cluster:
                target = cluster_dir / p.name
                try:
                    use_symlink = symlinks and not no_symlinks
                    if use_symlink:
                        if target.exists():
                            target.unlink()
                        target.symlink_to(p)
                    else:
                        if not target.exists():
                            shutil.copy2(p, target)
                except Exception:
                    if not target.exists():
                        shutil.copy2(p, target)
                writer.writerow([p.name, cid, rep.name, len(cluster)])
    return mapping_csv


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--pdf-dir', default='pdfs', help='relative path to pdfs directory (inside this folder)')
    ap.add_argument('--out-dir', default='templates', help='output templates dir')
    ap.add_argument('--min-sim', type=float, default=0.25, help='Jaccard similarity threshold for clustering')
    ap.add_argument('--shingle-size', type=int, default=5, help='k for k-word shingles')
    ap.add_argument('--max-pages', type=int, default=None, help='max pages to extract per PDF (default all)')
    ap.add_argument('--symlinks', action='store_true', default=True, help='create symlinks instead of copies (default True)')
    ap.add_argument('--no-symlinks', action='store_true', help='disable symlinks; use copies instead')
    args = ap.parse_args()

    base = Path(__file__).resolve().parent
    pdf_dir = (base / args.pdf_dir).resolve()
    out_dir = (base / args.out_dir).resolve()

    if not pdf_dir.exists() or not pdf_dir.is_dir():
        print('pdf-dir does not exist:', pdf_dir)
        return 1

    pdf_files = sorted([p for p in pdf_dir.iterdir() if p.is_file() and p.suffix.lower() == '.pdf'])
    if not pdf_files:
        print('No PDF files found in', pdf_dir)
        return 0

    print(f'Found {len(pdf_files)} PDF(s). Extracting text and building shingles...')
    shingles_map, metadata = build_shingles_map(pdf_files, shingle_size=args.shingle_size, max_pages=args.max_pages)

    print('Clustering (min_sim=%s)...' % args.min_sim)
    clusters = cpt.cluster_files(pdf_files, shingles_map, min_sim=args.min_sim)
    print(f'Produced {len(clusters)} clusters')

    mapping_csv = create_clusters_and_mapping(clusters, out_dir, symlinks=args.symlinks, no_symlinks=args.no_symlinks)
    print('Wrote mapping to', mapping_csv)

    # print brief summary
    for i, c in enumerate(clusters, start=1):
        print(f'  cluster_{i:03d}: {len(c)} pdf(s)')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
