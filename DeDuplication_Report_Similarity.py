"""
pbir_similarity.py
Compare multiple Power BI report projects (PBIR / project folder) by visual-level similarity.

Usage:
  - Put all your report folders (e.g., Report1.Report, Report2.Report, ...) under a single parent folder.
  - Set REPORTS_ROOT to that parent folder (or pass it via CLI if you add argparse).
  - Run: python pbir_similarity.py
Outputs:
  - Prints pairwise similarity matrix, groups for thresholds, master/grandmaster detection.
  - Saves an Excel file 'pbir_similarity_matrix.xlsx' in current working dir.
"""

import json, os
from pathlib import Path
from itertools import combinations
from collections import defaultdict
import pandas as pd
import time

# ---------------- Config ----------------

def get_reports_root():
    print(r"Example path format: C:\Users\YourName\Documents\Deduplication\Power BI Files\pbir files")

    while True:
        try:
            # Ask user for input
            user_input = input("Enter the folder path of PBIP files: ").strip().strip('"').strip("'")

            # Check if input is empty
            if not user_input:
                print("⚠️  Folder path cannot be empty. Please try again.\n")
                continue

            # Check if path exists
            if not os.path.exists(user_input):
                print("❌ The provided path does not exist. Please enter a valid folder path.\n")
                continue

            # Check if it's a directory
            if not os.path.isdir(user_input):
                print("⚠️  The provided path is not a folder. Please enter a valid directory path.\n")
                continue

            # Success
            print(f"✅ PBIP Folder Path set to: {user_input}")
            return user_input

        except KeyboardInterrupt:
            print("\nOperation cancelled by user. Exiting...")
            exit(1)
        except Exception as e:
            print(f"⚠️  An unexpected error occurred: {e}\nPlease try again.\n")
def get_output_path():
    print(r"Example output path format: C:\Users\YourName\Documents\Deduplication\Results")

    while True:
        try:
            user_input = input("Enter the folder path where output Excel file should be saved: ").strip().strip('"').strip("'")

            if not user_input:
                print("⚠️  Output folder path cannot be empty. Please try again.\n")
                continue

            if not os.path.exists(user_input):
                print("❌ The provided output path does not exist. Please enter a valid folder path.\n")
                continue

            if not os.path.isdir(user_input):
                print("⚠️  The provided output path is not a folder. Please enter a valid directory path.\n")
                continue

            print(f"✅ Output Folder Path set to: {user_input}")
            return user_input

        except KeyboardInterrupt:
            print("\nOperation cancelled by user. Exiting...")
            exit(1)
        except Exception as e:
            print(f"⚠️  An unexpected error occurred: {e}\nPlease try again.\n")

# Use the function
# REPORTS_ROOT = get_reports_root()
# OUTPUT_PATH = get_output_path()
REPORTS_ROOT = r"C:\Users\VaishnavKamartiMAQSo\Desktop\VS code explorations\DeDeuplication\Download_convert_reports\converted_pbip_files"
OUTPUT_PATH = r"C:\Users\VaishnavKamartiMAQSo\Desktop\VS code explorations\DeDeuplication\Download_convert_reports\similarity_output_index"
print("🕒 Running the script... please wait.")
time.sleep(2)

#REPORTS_ROOT = r"C:\Users\NikhilTiwariMAQSoftw\OneDrive - MAQ Software\Documents\Deduplication\Power BI Files\pbir files"  # <-- change this
VISUAL_MATCH_THRESHOLD = 0.9    # per-visual Jaccard to consider two visuals matching
GROUP_THRESHOLDS = [0.7, 0.8, 0.9, 0.95, 1.0]
MASTER_THRESHOLD = 0.95         # per-visual threshold when asserting full coverage (master)
# OUT_XLSX = "Report_similarity_matrix.xlsx"
OUT_XLSX = os.path.join(OUTPUT_PATH, "Report_similarity_matrix.xlsx")

# ----------------------------------------

def jaccard(a,b):
    if not a and not b: return 1.0
    if not a or not b: return 0.0
    a,b = set(a), set(b)
    return len(a & b) / len(a | b)

def greedy_visual_match(visuals_a, visuals_b, per_visual_threshold=VISUAL_MATCH_THRESHOLD):
    """
    Return matched_count, pairs, scores
    pairs: list of (index_in_A, index_in_B)
    """
    unmatched_b = set(range(len(visuals_b)))
    pairs = []
    scores = []
    for i, va in enumerate(visuals_a):
        best_j = -1.0
        best_b_idx = None
        for j in list(unmatched_b):
            vb = visuals_b[j]
            score = jaccard(va['fields'], vb['fields'])
            # prefer same visual type on equal scores
            if abs(score - best_j) < 1e-9 and va.get('type') and vb.get('type') and va['type'] == vb['type']:
                best_b_idx = j
                best_j = score
            elif score > best_j:
                best_b_idx = j
                best_j = score
        if best_j >= per_visual_threshold and best_b_idx is not None:
            unmatched_b.remove(best_b_idx)
            pairs.append((i, best_b_idx))
            scores.append(best_j)
    return len(pairs), pairs, scores

def report_similarity(visuals_a, visuals_b, per_visual_threshold=VISUAL_MATCH_THRESHOLD):
    if not visuals_a and not visuals_b:
        return 1.0, 0, []
    matched_count, pairs, scores = greedy_visual_match(visuals_a, visuals_b, per_visual_threshold)
    denom = len(visuals_a) + len(visuals_b)
    score = 0.0 if denom == 0 else (2.0 * matched_count) / denom
    return score, matched_count, list(zip(pairs, scores))

# ---------------- JSON field extraction heuristics ----------------
def extract_fields_from_json(obj):
    found = set()
    if isinstance(obj, dict):
        for k,v in obj.items():
            kl = k.lower()
            if kl in ('queryref','queryname','field','displayname','name','column','expr','measure'):
                if isinstance(v, str):
                    found.add(v.strip().lower())
                elif isinstance(v, dict) and 'expression' in v:
                    found.add(str(v['expression']).strip().lower())
            if isinstance(v, (dict, list)):
                found |= extract_fields_from_json(v)
            elif isinstance(v, str) and kl in ('displayname','name'):
                found.add(v.strip().lower())
    elif isinstance(obj, list):
        for item in obj:
            if isinstance(item, str):
                found.add(item.strip().lower())
            else:
                found |= extract_fields_from_json(item)
    return found

def parse_report_visuals(report_folder):
    """
    Heuristic parser for a report project folder (PBIR).
    Returns list of visuals: each is {'id':..., 'type':..., 'fields': set([...])}
    """
    report_path = Path(report_folder)
    visuals = []

    # First, look for explicit visual.json (enhanced PBIR format)
    for visual_file in report_path.rglob('**/visual.json'):
        try:
            with open(visual_file, 'r', encoding='utf-8') as f:
                vdoc = json.load(f)
            vtype = (vdoc.get('visualType') or vdoc.get('type') or '').lower()
            fields = set()
            if isinstance(vdoc, dict):
                if 'fields' in vdoc and isinstance(vdoc['fields'], list):
                    for it in vdoc['fields']:
                        if isinstance(it, str): fields.add(it.strip().lower())
                if 'projections' in vdoc and isinstance(vdoc['projections'], dict):
                    for arr in vdoc['projections'].values():
                        if isinstance(arr, list):
                            for it in arr:
                                if isinstance(it, str): fields.add(it.strip().lower())
                                elif isinstance(it, dict):
                                    fields |= extract_fields_from_json(it)
                fields |= extract_fields_from_json(vdoc)
            visuals.append({'id': visual_file.name, 'type': vtype, 'fields': {f for f in fields if f}})
        except Exception:
            pass

    # Fallback: scan all jsons and look for explicit 'fields' or 'projections' keys
    if not visuals:
        for jf in report_path.rglob('*.json'):
            try:
                with open(jf, 'r', encoding='utf-8') as f:
                    doc = json.load(f)
                if not isinstance(doc, dict): 
                    continue
                fields = set()
                vtype = (doc.get('visualType') or doc.get('type') or '').lower()
                if 'fields' in doc and isinstance(doc['fields'], list):
                    for it in doc['fields']:
                        if isinstance(it, str): fields.add(it.strip().lower())
                if 'projections' in doc and isinstance(doc['projections'], dict):
                    for arr in doc['projections'].values():
                        if isinstance(arr, list):
                            for it in arr:
                                if isinstance(it, str): fields.add(it.strip().lower())
                                elif isinstance(it, dict):
                                    fields |= extract_fields_from_json(it)
                fields |= extract_fields_from_json(doc)
                # remove tokens that equal the visual type name (noise)
                fields = {f for f in fields if f and f != vtype}
                if fields:
                    visuals.append({'id': jf.name, 'type': vtype, 'fields': fields})
            except Exception:
                pass

    # deduplicate
    seen = set(); cleaned = []
    for v in visuals:
        sig = (v['type'], tuple(sorted(v['fields'])))
        if sig not in seen:
            seen.add(sig); cleaned.append(v)
    return cleaned

# ---------------- grouping, master detection ----------------
def find_connected_components(nodes, edges):
    visited = set(); comps = []
    for n in nodes:
        if n in visited: continue
        stack = [n]; comp = set()
        while stack:
            cur = stack.pop()
            if cur in visited: continue
            visited.add(cur); comp.add(cur)
            for nb in edges[cur]:
                if nb not in visited: stack.append(nb)
            # treat edges as undirected for connectivity (also inspect reverse)
            for other, outs in edges.items():
                if cur in outs and other not in visited:
                    stack.append(other)
        comps.append(comp)
    return comps

def detect_masters_tiebreak(names, visuals_by_report, per_visual_threshold=MASTER_THRESHOLD):
    masters = defaultdict(list)
    for r1, r2 in combinations(names, 2):
        v1 = visuals_by_report[r1]; v2 = visuals_by_report[r2]
        m12, _, _ = greedy_visual_match(v1, v2, per_visual_threshold)
        m21, _, _ = greedy_visual_match(v2, v1, per_visual_threshold)
        # r1 master of r2 if r1 covers r2 and is strict superset OR lexicographically smaller on tie
        if m12 == len(v2) and (len(v1) > len(v2) or (len(v1)==len(v2) and r1 < r2)):
            masters[r1].append(r2)
        if m21 == len(v1) and (len(v2) > len(v1) or (len(v2)==len(v1) and r2 < r1)):
            masters[r2].append(r1)
    return dict(masters)

def transitive_closure(master_edges):
    adj = defaultdict(set); nodes = set()
    for m, childs in master_edges.items():
        nodes.add(m)
        for c in childs:
            nodes.add(c); adj[m].add(c)
    closure = {n:set() for n in nodes}
    for n in nodes:
        stack = list(adj[n])
        while stack:
            cur = stack.pop()
            if cur in closure[n]: continue
            closure[n].add(cur)
            for nxt in adj.get(cur, []):
                if nxt not in closure[n]: stack.append(nxt)
    return closure

# -------------------- main run --------------------
def main():
    root = Path(REPORTS_ROOT)
    report_dirs = sorted([p for p in root.iterdir() if p.is_dir()])
    if not report_dirs:
        print("No report folders found under:", root); return

    visuals_by_report = {}
    for r in report_dirs:
        visuals_by_report[r.name] = parse_report_visuals(r)

    # debug print
    # for name, vis in visuals_by_report.items():
        # print(f"\n{name}: {len(vis)} visuals")
        # print("")
        # for v in vis:
            # print(f" - {v['id']} (type={v['type']}) fields={sorted(v['fields'])}")
            # print()

    # pairwise similarity matrix
    names = sorted(list(visuals_by_report.keys()))
    sim = pd.DataFrame(index=names, columns=names, dtype=float)
    for a in names:
        for b in names:
            s, matched, _ = report_similarity(visuals_by_report[a], visuals_by_report[b], VISUAL_MATCH_THRESHOLD)
            sim.loc[a,b] = round(s,4)

    # print("\nSimilarity matrix:\n", sim)
    sim.to_excel(OUT_XLSX)
    # print("\nSaved similarity matrix as:", OUT_XLSX,"to this path: ",os.getcwd())
    # print(f"\n✅ Saved similarity matrix to this path: {OUT_XLSX}")


    # groups at thresholds
    for thr in GROUP_THRESHOLDS:
        edges = defaultdict(set)
        for i,j in combinations(names, 2):
            if sim.loc[i,j] >= thr:
                edges[i].add(j); edges[j].add(i)
        comps = find_connected_components(names, edges)
        # print(f"\nThreshold {int(thr*100)}% -> groups: {comps}")

    # Calculate similarity stats
    total_reports = len(names)
    similar_groups = [c for c in comps if len(c) > 1]
    similar_count = sum(len(c) for c in similar_groups)
    efficiency = (similar_count / total_reports) * 100 if total_reports else 0

    

    # masters
    masters = detect_masters_tiebreak(names, visuals_by_report, MASTER_THRESHOLD)
    closure = transitive_closure(masters)
    # print("\nDirect masters:", masters)
    # print("\nTransitive closure:", closure)

    # elimination heuristic (keep roots = reports that are not child of any master)
    children = set(c for childs in masters.values() for c in childs)
    roots = sorted(list(set(names) - children))
    eliminate = sorted(list(children))
    print(f"\n  Reports to keep:", roots)
    print(f"\n  Reports eligible for elimination (have a master):", eliminate)
    print(f"\n🔹 At {int(thr*100)}% threshold:")
    print(f"   → {similar_count} out of {total_reports} reports were identified as similar.")
    print(f"   → Efficiency: {efficiency:.2f}%")
    print("\n✅ Script execution completed successfully!")
    print(f"\n✅ Saved similarity matrix to this path: {OUT_XLSX}")
    
    time.sleep(60)



if __name__ == "__main__":
    main()
