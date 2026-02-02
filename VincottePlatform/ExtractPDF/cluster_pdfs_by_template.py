#!/usr/bin/env python3
"""
Cluster PDFs by template-similarity using text shingles (5-word shingles) and greedy Jaccard clustering.
Creates a templates output dir with cluster_{:03d} subfolders containing symlinks to the original PDFs
and writes a mapping CSV (`clusters_mapping.csv`) in the output dir.

New features:
- --export-templates: for each cluster write a template.json with representative filename and a "logo signature"
  (first-page image count + normalized centers) that can be used for assignment.
- --assign-templates: read templates from the templates dir and assign each PDF to the best template based on
  Jaccard(shingles) with optional logo-match boost. Writes `assignment.csv` and creates `assigned/cluster_xxx` dirs.

Usage examples:
  build clusters + export templates:
    python cluster_pdfs_by_template.py --pdf-dir pdfs --out-dir templates --min-sim 0.25 --export-templates

  assign PDFs using existing templates (no reclustering):
    python cluster_pdfs_by_template.py --pdf-dir pdfs --out-dir templates --assign-templates --assign-threshold 0.20

Notes:
- Requires `pdfplumber` installed in the venv (already used in this repo).
- This is a heuristic quick clustering/assignment; tune thresholds for your data.
"""
import argparse
from pathlib import Path
import pdfplumber
import re
import shutil
import csv
import itertools
import json
import math
# add image hashing deps
try:
    from PIL import Image
    import imagehash
except Exception:
    Image = None
    imagehash = None


def norm_text(s: str) -> str:
    if not s:
        return ""
    s = s.lower()
    # replace common broken characters and multiple whitespace
    s = re.sub(r"[\r\n]+", " ", s)
    s = re.sub(r"[^a-z0-9\sçéèàôûäöüß-]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def tokens_from_text(s: str):
    if not s:
        return []
    return [t for t in s.split(" ") if len(t) > 2]


def shingles(tokens, k=5):
    if not tokens:
        return set()
    if len(tokens) < k:
        return {" ".join(tokens)}
    return {" ".join(tokens[i:i + k]) for i in range(len(tokens) - k + 1)}


def jaccard(a: set, b: set):
    if not a or not b:
        return 0.0
    inter = len(a & b)
    uni = len(a | b)
    return inter / uni if uni else 0.0


def extract_pdf_text(path: Path, max_pages=None):
    try:
        with pdfplumber.open(path) as pdf:
            pages_to_read = range(len(pdf.pages)) if max_pages is None else range(min(len(pdf.pages), max_pages))
            txt = []
            for i in pages_to_read:
                page = pdf.pages[i]
                t = page.extract_text()
                if t:
                    txt.append(t)
            return " \n " .join(txt)
    except Exception:
        # don't crash on unreadable PDFs
        return ""


def extract_first_page_logo_signature(path: Path):
    """Return a simple logo signature for the first page: image_count and list of normalized centers [(cx, cy), ...].
    If pdf has no pages or no images, returns {image_count:0, centers:[]}.
    """
    try:
        with pdfplumber.open(path) as pdf:
            if not pdf.pages:
                return {"image_count": 0, "centers": []}
            p0 = pdf.pages[0]
            images = p0.images if hasattr(p0, 'images') else []
            w = p0.width or 1.0
            h = p0.height or 1.0
            centers = []
            for im in images:
                # pdfplumber images have x0,x1,top,bottom or x0,x1,y0,y1 depending on version
                x0 = im.get('x0') or im.get('x0', 0)
                x1 = im.get('x1') or im.get('x1', 0)
                y0 = im.get('y0') or im.get('top') or 0
                y1 = im.get('y1') or im.get('bottom') or 0
                # compute center normalized
                cx = ((x0 + x1) / 2.0) / w
                cy = ((y0 + y1) / 2.0) / h
                centers.append((round(cx, 4), round(cy, 4)))
            return {"image_count": len(centers), "centers": centers}
    except Exception:
        return {"image_count": 0, "centers": []}


# new: compute perceptual hash (pHash) of page 1 using pdfplumber's page.to_image() if available
def compute_first_page_phash(path: Path, hash_size=16, res=150):
    """Return a hex string of pHash for page 1, or None if unavailable.
    hash_size controls the pHash size (larger = more bits).
    """
    if imagehash is None:
        return None
    try:
        with pdfplumber.open(path) as pdf:
            if not pdf.pages:
                return None
            p0 = pdf.pages[0]
            # pdfplumber Page.to_image() returns an object with .original (PIL.Image) in many versions
            try:
                pi = p0.to_image(resolution=res)
                img = getattr(pi, 'original', None)
                if img is None:
                    # fallback: try to export a crop image via page.crop_to_bbox? skip if not available
                    return None
            except Exception:
                # fallback: cannot render page image
                return None
            try:
                ph = imagehash.phash(img, hash_size=hash_size)
                return str(ph)
            except Exception:
                return None
    except Exception:
        return None


def cluster_files(files, shingles_map, min_sim=0.25):
    clusters = []  # list of lists of filenames
    reps = []  # representative shingle set per cluster
    for f in files:
        sset = shingles_map.get(f, set())
        placed = False
        for ci, repset in enumerate(reps):
            sim = jaccard(sset, repset)
            if sim >= min_sim:
                clusters[ci].append(f)
                # update repset as union (makes cluster broader)
                reps[ci] = repset | sset
                placed = True
                break
        if not placed:
            clusters.append([f])
            reps.append(set(sset))
    return clusters


# --- New: inspector helper to explain why two clusters are separate ---
def inspect_clusters(cluster_a_files, cluster_b_files, shingles_map, shingle_size=5, max_pairwise_report=5):
    """Compute union representative shingles and pairwise similarities between two clusters.

    Returns a dictionary with statistics and prints a concise report to stdout.
    """
    def union_shingles(files):
        rep = set()
        for f in files:
            rep |= shingles_map.get(f, set())
        return rep

    rep_a = union_shingles(cluster_a_files)
    rep_b = union_shingles(cluster_b_files)

    rep_j = jaccard(rep_a, rep_b)

    # compute max pairwise jaccard between members
    max_pair = 0.0
    best_pair = (None, None)
    pair_scores = []
    for fa, fb in itertools.product(cluster_a_files, cluster_b_files):
        sa = shingles_map.get(fa, set())
        sb = shingles_map.get(fb, set())
        sim = jaccard(sa, sb)
        pair_scores.append((sim, fa, fb))
        if sim > max_pair:
            max_pair = sim
            best_pair = (fa, fb)

    pair_scores.sort(reverse=True, key=lambda x: x[0])

    # report
    print('\n--- Cluster inspector report ---')
    print(f'Cluster A: {len(cluster_a_files)} file(s), union shingles={len(rep_a)}')
    print(f'Cluster B: {len(cluster_b_files)} file(s), union shingles={len(rep_b)}')
    print(f'Jaccard(rep_A, rep_B) = {rep_j:.4f}')
    print(f'Max pairwise file Jaccard = {max_pair:.4f} between {best_pair[0].name if best_pair[0] else None} and {best_pair[1].name if best_pair[1] else None}')

    print('\nTop pairwise similarities:')
    for sim, fa, fb in pair_scores[:max_pairwise_report]:
        print(f'  {sim:.4f}: {fa.name}  <->  {fb.name}')

    # show some basic metadata for top files
    def meta(f):
        s = shingles_map.get(f, set())
        return {'name': f.name, 'shingles': len(s)}

    top_files = sorted(set(cluster_a_files + cluster_b_files), key=lambda x: -len(shingles_map.get(x, set())))[:10]
    print('\nTop files by shingles (name / shingles):')
    for f in top_files:
        print(f'  {f.name}: {len(shingles_map.get(f, set()))}')

    print('--- End report ---\n')

    return {
        'rep_jaccard': rep_j,
        'max_pairwise': max_pair,
        'best_pair': best_pair,
        'pair_scores': pair_scores,
        'rep_a_size': len(rep_a),
        'rep_b_size': len(rep_b),
    }


# --- New: export templates metadata and assign-by-template logic ---
def export_templates(out_dir: Path, clusters, base_pdf_dir: Path, shingles_map):
    """Write template metadata (template.json) into each cluster folder. Template contains rep filename and logo signature."""
    for i, cluster in enumerate(clusters, start=1):
        cid = f'cluster_{i:03d}'
        cluster_dir = out_dir / cid
        cluster_dir.mkdir(parents=True, exist_ok=True)
        rep = cluster[0]
        rep_path = base_pdf_dir / rep.name
        logo_sig = extract_first_page_logo_signature(rep_path)
        logo_phash = compute_first_page_phash(rep_path)
        meta = {
            'rep': rep.name,
            'logo_signature': logo_sig,
            'logo_phash': logo_phash,
            'cluster_size': len(cluster)
        }
        with (cluster_dir / 'template.json').open('w', encoding='utf-8') as fh:
            json.dump(meta, fh, ensure_ascii=False, indent=2)


def compare_logo_sigs(sig_a, sig_b, max_center_dist=0.08, max_hamming=6):
    """Return True if logo signatures match closely.
    Prefer pHash comparison (if both present); otherwise fall back to center-based comparison.
    """
    if not sig_a or not sig_b:
        return False
    # try phash if available
    pa = sig_a.get('logo_phash') if isinstance(sig_a, dict) else None
    pb = sig_b.get('logo_phash') if isinstance(sig_b, dict) else None
    if pa and pb and imagehash is not None:
        try:
            ha = imagehash.hex_to_hash(pa)
            hb = imagehash.hex_to_hash(pb)
            dist = ha - hb
            return dist <= max_hamming
        except Exception:
            pass
    # fallback to image_count + centers
    if sig_a.get('image_count', 0) == 0 or sig_b.get('image_count', 0) == 0:
        return False
    if sig_a.get('image_count') != sig_b.get('image_count'):
        return False
    ca = sig_a.get('centers', [])
    cb = sig_b.get('centers', [])
    if not ca or not cb or len(ca) != len(cb):
        return False
    dists = []
    for (ax, ay), (bx, by) in zip(ca, cb):
        d = math.hypot(ax - bx, ay - by)
        dists.append(d)
    mean_d = sum(dists) / len(dists) if dists else 1.0
    return mean_d <= max_center_dist


def assign_by_templates(out_dir: Path, base_pdf_dir: Path, pdf_files, shingles_map, min_sim=0.25, logo_boost=0.10, assign_threshold=None):
    """Assign each PDF to the best template. Returns stats dict and writes assignment CSV and assigned symlinks."""
    templates = []  # list of (cid, rep_path, rep_shingles, logo_sig, logo_phash)
    for d in sorted(out_dir.iterdir()):
        if not d.is_dir() or not d.name.startswith('cluster_'):
            continue
        tfile = d / 'template.json'
        if not tfile.exists():
            continue
        meta = json.load(tfile.open(encoding='utf-8'))
        rep_name = meta.get('rep')
        rep_path = base_pdf_dir / rep_name
        rep_shingles = shingles_map.get(rep_path, set())
        logo_sig = meta.get('logo_signature', {})
        logo_phash = meta.get('logo_phash')
        templates.append((d.name, rep_path, rep_shingles, logo_sig, logo_phash))

    if not templates:
        print('No templates found in', out_dir)
        return {'assigned': 0, 'unassigned': len(pdf_files)}

    assigned_dir = out_dir / 'assigned'
    assigned_dir.mkdir(parents=True, exist_ok=True)

    assign_threshold = assign_threshold if assign_threshold is not None else min_sim

    rows = []
    assigned_count = 0
    for p in pdf_files:
        best_score = -1.0
        best_template = None
        p_sh = shingles_map.get(p, set())
        p_logo_sig = extract_first_page_logo_signature(p)
        p_phash = compute_first_page_phash(p)
        for cid, rep_path, rep_sh, logo_sig, logo_phash in templates:
            sim = jaccard(p_sh, rep_sh)
            logo_match = False
            # prefer phash if present
            if p_phash and logo_phash and imagehash is not None:
                try:
                    ha = imagehash.hex_to_hash(p_phash)
                    hb = imagehash.hex_to_hash(logo_phash)
                    dist = ha - hb
                    logo_match = (dist <= 6)
                except Exception:
                    logo_match = compare_logo_sigs(p_logo_sig, logo_sig)
            else:
                logo_match = compare_logo_sigs(p_logo_sig, logo_sig)
            score = sim + (logo_boost if logo_match else 0.0)
            if score > best_score:
                best_score = score
                best_template = (cid, rep_path, logo_sig, sim, logo_match)
        assigned_to = None
        if best_score >= assign_threshold:
            cid = best_template[0]
            target_dir = assigned_dir / cid
            target_dir.mkdir(parents=True, exist_ok=True)
            # symlink into assigned folder
            tpath = target_dir / p.name
            if tpath.exists():
                try:
                    tpath.unlink()
                except Exception:
                    pass
            try:
                tpath.symlink_to(p)
            except Exception:
                shutil.copy2(p, tpath)
            assigned_to = cid
            assigned_count += 1
        rows.append({
            'pdf': p.name,
            'best_template': best_template[0] if best_template else None,
            'jaccard': round(best_template[3], 4) if best_template else 0.0,
            'logo_match': bool(best_template[4]) if best_template else False,
            'score': round(best_score, 4),
            'assigned_to': assigned_to
        })

    # write CSV
    csvp = out_dir / 'assignment.csv'
    with csvp.open('w', newline='', encoding='utf-8') as mf:
        writer = csv.writer(mf)
        writer.writerow(['pdf', 'best_template', 'jaccard', 'logo_match', 'score', 'assigned_to'])
        for r in rows:
            writer.writerow([r['pdf'], r['best_template'], r['jaccard'], r['logo_match'], r['score'], r['assigned_to']])

    print(f'Assigned {assigned_count}/{len(pdf_files)} PDFs to templates. Wrote {csvp}')
    return {'assigned': assigned_count, 'unassigned': len(pdf_files) - assigned_count}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--pdf-dir', default='pdfs', help='relative path to pdfs directory (inside this folder)')
    ap.add_argument('--out-dir', default='templates', help='output templates dir')
    ap.add_argument('--min-sim', type=float, default=0.25, help='Jaccard similarity threshold for clustering')
    ap.add_argument('--shingle-size', type=int, default=5, help='k for k-word shingles')
    ap.add_argument('--max-pages', type=int, default=None, help='max pages to extract per PDF (default all)')
    # make symlinks default True (help text expected this behaviour)
    ap.add_argument('--symlinks', action='store_true', default=True, help='create symlinks instead of copies (default True)')
    ap.add_argument('--no-symlinks', action='store_true', help='disable symlinks; use copies instead')
    ap.add_argument('--inspect-clusters', nargs=2, metavar=('CLUSTER_A', 'CLUSTER_B'),
                    help='Inspect two clusters (e.g. cluster_002 cluster_003) and print similarity diagnostics. If used the script will exit after the report.')

    # new flags
    ap.add_argument('--export-templates', action='store_true', help='Export template metadata (template.json) for each cluster')
    ap.add_argument('--assign-templates', action='store_true', help='Assign all PDFs to templates found in out-dir')
    ap.add_argument('--assign-threshold', type=float, default=None, help='Minimum score (jaccard + logo boost) to assign a PDF to a template')
    ap.add_argument('--logo-boost', type=float, default=0.10, help='Score boost to add when logo signatures match')

    args = ap.parse_args()

    base = Path(__file__).resolve().parent
    pdf_dir = (base / args.pdf_dir).resolve()
    out_dir = (base / args.out_dir).resolve()

    if not pdf_dir.exists() or not pdf_dir.is_dir():
        print('pdf-dir does not exist:', pdf_dir)
        return 1

    out_dir.mkdir(parents=True, exist_ok=True)

    pdf_files = sorted([p for p in pdf_dir.iterdir() if p.is_file() and p.suffix.lower() == '.pdf'])
    if not pdf_files:
        print('No PDF files found in', pdf_dir)
        return 0

    print(f'Found {len(pdf_files)} PDF(s). Extracting text...')

    shingles_map = {}
    metadata = {}
    for p in pdf_files:
        txt = extract_pdf_text(p, max_pages=args.max_pages)
        nt = norm_text(txt)
        toks = tokens_from_text(nt)
        sh = shingles(toks, k=args.shingle_size)
        shingles_map[p] = sh
        metadata[p] = {'text_len': len(nt), 'tokens': len(toks)}
        print(f'  {p.name}: text_len={len(nt)}, tokens={len(toks)}, shingles={len(sh)}')

    # If assign-templates requested without templates existing, we still allow running by building clusters first
    # but if user only wants assign step and templates are already present, skip clustering by setting build_clusters=False
    build_clusters = True
    clusters = []
    if args.assign_templates and any((out_dir / d).exists() for d in out_dir.iterdir() if d.name.startswith('cluster_')):
        # templates exist; we'll load them later and skip writing clusters mapping
        build_clusters = False

    if build_clusters:
        print('Clustering (min_sim=%s)...' % args.min_sim)
        clusters = cluster_files(pdf_files, shingles_map, min_sim=args.min_sim)
        print(f'Produced {len(clusters)} clusters')

        # create cluster dirs and symlinks/copies
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
                    # create symlink if requested; otherwise copy
                    try:
                        use_symlink = args.symlinks and not args.no_symlinks
                        if use_symlink:
                            if target.exists():
                                target.unlink()
                            target.symlink_to(p)
                        else:
                            # copy file if it doesn't exist
                            if not target.exists():
                                shutil.copy2(p, target)
                    except Exception:
                        # fallback to copy
                        if not target.exists():
                            shutil.copy2(p, target)
                    writer.writerow([p.name, cid, rep.name, len(cluster)])

        print('Wrote mapping to', mapping_csv)
    else:
        # load clusters from existing cluster dirs (only cluster_{xxx} dirs)
        clusters = []
        for d in sorted(out_dir.iterdir()):
            if d.is_dir() and d.name.startswith('cluster_'):
                files = sorted([f for f in d.iterdir() if f.is_file() and f.suffix.lower() == '.pdf'])
                # map to original pdf paths in pdf_dir
                cluster_files_list = [pdf_dir / f.name for f in files]
                clusters.append(cluster_files_list)
        print(f'Loaded {len(clusters)} clusters from existing template folders')

    # If user requested an inspection of two clusters, perform it and exit early
    if args.inspect_clusters:
        a_id, b_id = args.inspect_clusters
        # normalize possible id formats: 'cluster_002', '002', '2'
        def find_cluster_index(cid):
            # direct match
            for idx, cluster in enumerate(clusters, start=1):
                name = f'cluster_{idx:03d}'
                if cid == name or cid == str(idx) or cid == f'{idx:03d}':
                    return idx - 1
            return None

        ai = find_cluster_index(a_id)
        bi = find_cluster_index(b_id)
        if ai is None or bi is None:
            print('Could not find one of the requested clusters:', a_id, b_id)
            print('Available clusters:', [f'cluster_{i:03d}' for i in range(1, len(clusters) + 1)])
            return 2

        cluster_a_files = clusters[ai]
        cluster_b_files = clusters[bi]
        inspect_clusters(cluster_a_files, cluster_b_files, shingles_map, shingle_size=args.shingle_size)
        return 0

    # export templates metadata if requested
    if args.export_templates:
        print('Exporting templates metadata into each cluster dir...')
        export_templates(out_dir, clusters, pdf_dir, shingles_map)
        print('Templates exported.')

    # assign all pdfs using templates if requested
    if args.assign_templates:
        print('Assigning PDFs using templates found in', out_dir)
        stats = assign_by_templates(out_dir, pdf_dir, pdf_files, shingles_map, min_sim=args.min_sim, logo_boost=args.logo_boost, assign_threshold=args.assign_threshold)
        print('Assign stats:', stats)

    print('Cluster summary:')
    for i, c in enumerate(clusters, start=1):
        print(f'  cluster_{i:03d}: {len(c)} pdf(s)')

    return 0


if __name__ == '__main__':
    raise SystemExit(main())
