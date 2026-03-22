#!/usr/bin/env python3
## SCRIPT_RUNNER
## param: input,               label=Input HTML Profile,       type=file-in,  filter=HTML Profile (*.html)|*.html;*.htm
## param: --out,               label=Output File,              type=file-out, filter=HTML Profile (*.html)|*.html,       optional=true
## param: --min-query-time,    label=Min Query Time (s),       type=float,    default=0.05,  optional=true
## param: --min-flat-utime,    label=Min Flat Profile utime (s), type=float,  default=0.001, optional=true
## param: --min-obj-cpu,       label=Min Object CPU Time (s),  type=float,    default=0.05,  optional=true
## param: --keep-rtime-graph,  label=Keep rtime Call Graph,    type=flag,                    optional=true
"""
Infor LN Call Graph Profile Trimmer
====================================
Reduces file size of HTML Call Graph Profiles without deleting meaningful data.

Strategies applied (all safe, no important data lost):
  1. Filter Query Summary rows below a minimum rtime threshold
  2. Filter Flat Profile (CPU-sorted) rows below a minimum utime threshold  
  3. Filter Object Summary rows below a minimum cpu-time threshold
  4. Optionally remove the entire rtime-sorted Call Graph (1.1.x) section
     — it is a duplicate view of the same data sorted differently.
     The (rtime) links in the Query Summary will become inactive, but
     all CPU-sorted (0.1.x) links remain intact.

Usage:
  python trim_ln_profile.py <input.html> [options]

Options:
  --out FILE                Output filename (default: <input>_trimmed.html)
  --min-query-time S        Drop Query Summary rows with rtime < S  (default 0.05)
  --min-flat-utime S        Drop Flat Profile rows with utime < S   (default 0.1)
  --min-obj-cpu S           Drop Object Summary rows with cpu < S   (default 0.05)
  --keep-rtime-graph        Keep the rtime-sorted (1.1.x) Call Graph (default: remove it)
  --dry-run                 Print savings estimate only, don't write output
"""

import re
import sys
import argparse
from pathlib import Path


# ─── Section detection helpers ───────────────────────────────────────────────

SECTION_MARKERS = {
    "query_summary":   re.compile(r'<A ID="query"></A>'),
    "object_summary":  re.compile(r'<A ID="object\.0\.1"></A>'),
    "callgraph_cpu":   re.compile(r'<A ID="profile\.0\.1"></A>'),
    "flatprofile_cpu": re.compile(r'<A ID="flatprofile\.0\.1"></A>'),
    "callgraph_rt":    re.compile(r'<A ID="profile\.1\.1"></A>'),   # may not exist as explicit anchor
    "cg_1_1_first":    re.compile(r'<A ID="1\.1\.1"></A>'),         # first rtime CG node
    "flatprofile_rt":  re.compile(r'<A ID="flatprofile\.1\.1"></A>'),
}


def find_sections(lines):
    """Return dict mapping section name -> line index."""
    found = {}
    for i, line in enumerate(lines):
        for name, pat in SECTION_MARKERS.items():
            if name not in found and pat.search(line):
                found[name] = i
    return found


# ─── Row-level filters ────────────────────────────────────────────────────────

# Query Summary row: first <td> is query id, second is total rtime
# <tr ><td>...NNN</td><td>  X.XXX  </td>...
RE_QUERY_ROW  = re.compile(r'^<tr [^>]*><td><A HREF="#query id')
RE_QUERY_TIME = re.compile(r'<td>\s*([\d.]+)\s*</td>')   # second td = total rtime

# Flat Profile row — each row is one long line
RE_FLAT_ROW   = re.compile(r'^<tr ><td data=')
# utime is the 2nd <td data=...><b>...</b> cell; data= may have leading spaces
RE_FLAT_UTIME = re.compile(r'<td data="\s*([\d.]+)"\s*><b>\s*([\d.]+)\s*</b></td>')

# Object Summary – per-DLL table rows  (e.g. <tr><td><A HREF="#otfglddll5456.0">...)
RE_OBJ_ROW  = re.compile(r'^<tr><td><A HREF="#ot')
RE_OBJ_CPU  = re.compile(r'<td><B>\s*([\d.]+)</B></td>')


def get_second_float(line, pattern):
    """Extract second float match from line using pattern."""
    matches = pattern.findall(line)
    if len(matches) >= 2:
        try:
            return float(matches[1])
        except (ValueError, IndexError):
            pass
    return None


def get_first_float(line, pattern):
    matches = pattern.findall(line)
    if matches:
        try:
            return float(matches[0])
        except ValueError:
            pass
    return None


# ─── Main trimmer ─────────────────────────────────────────────────────────────

def trim(lines, sections, args):
    keep_rtime_graph = args.keep_rtime_graph
    min_query   = args.min_query_time
    min_flat    = args.min_flat_utime
    min_obj     = args.min_obj_cpu

    # Determine rtime call graph range
    cg1_start = sections.get("cg_1_1_first") or sections.get("callgraph_rt")
    # The rtime section runs from cg1_start to end of file (after rtime flat profile)
    # We'll remove everything from the first 1.1.1 node line back to the preceding <H2>
    # Actually we remove from the <H3> or <H2> header just before 1.1.1

    # Find the H2/H3 header line just before cg1_start
    cg1_header = cg1_start
    if cg1_start:
        for back in range(cg1_start - 1, max(cg1_start - 20, 0), -1):
            if re.search(r'<H[23]>', lines[back], re.IGNORECASE):
                cg1_header = back
                break

    result = []
    total_orig = len(lines)

    # Counters for stats
    q_kept = q_dropped = 0
    f_kept = f_dropped = 0
    o_kept = o_dropped = 0
    cg_lines_removed = 0

    in_query_section  = False
    in_object_section = False
    in_flat_section   = False
    in_rtime_graph    = False

    # Track which sections we've passed
    passed_query_end  = False

    i = 0
    while i < len(lines):
        line = lines[i]

        # ── Detect section boundaries ────────────────────────────────────────
        if SECTION_MARKERS["query_summary"].search(line):
            in_query_section  = True
            in_object_section = False
            in_flat_section   = False
        elif SECTION_MARKERS["object_summary"].search(line):
            in_query_section  = False
            in_object_section = True
            in_flat_section   = False
        elif SECTION_MARKERS["callgraph_cpu"].search(line):
            in_object_section = False
            in_flat_section   = False
        elif SECTION_MARKERS["flatprofile_cpu"].search(line):
            in_flat_section   = True
            in_query_section  = False
            in_object_section = False
        elif SECTION_MARKERS["flatprofile_rt"].search(line):
            in_flat_section   = True   # rtime flat, same structure

        # Close query section at its </table>
        if in_query_section and '</table>' in line and not RE_QUERY_ROW.search(line):
            in_query_section = False
            passed_query_end = True

        # ── rtime Call Graph removal ─────────────────────────────────────────
        if not keep_rtime_graph and cg1_start and i >= cg1_header:
            in_rtime_graph = True

        if in_rtime_graph:
            cg_lines_removed += 1
            i += 1
            continue

        # ── Query Summary filter ─────────────────────────────────────────────
        if in_query_section and RE_QUERY_ROW.search(line):
            # Extract total rtime (second <td> value)
            tds = RE_QUERY_TIME.findall(line)
            rtime = float(tds[1]) if len(tds) >= 2 else 999.0
            if rtime < min_query:
                q_dropped += 1
                i += 1
                continue
            else:
                q_kept += 1

        # ── Object Summary filter ────────────────────────────────────────────
        if in_object_section and RE_OBJ_ROW.search(line):
            # cpu time is in the bold <B> cell
            cpu = get_first_float(line, RE_OBJ_CPU)
            if cpu is not None and cpu < min_obj:
                o_dropped += 1
                i += 1
                continue
            else:
                o_kept += 1

        # ── Flat Profile filter ──────────────────────────────────────────────
        if in_flat_section and RE_FLAT_ROW.search(line):
            # utime is the second data= cell
            matches = RE_FLAT_UTIME.findall(line)
            # matches are (data_val, display_val) tuples for bold cells
            # utime is second bold td (index 1)
            utime = float(matches[1][0]) if len(matches) >= 2 else 999.0
            if utime < min_flat:
                f_dropped += 1
                i += 1
                continue
            else:
                f_kept += 1

        result.append(line)
        i += 1

    stats = {
        "original_lines":   total_orig,
        "output_lines":     len(result),
        "lines_saved":      total_orig - len(result),
        "pct_saved":        100.0 * (total_orig - len(result)) / total_orig,
        "query_kept":       q_kept,
        "query_dropped":    q_dropped,
        "flat_kept":        f_kept,
        "flat_dropped":     f_dropped,
        "obj_kept":         o_kept,
        "obj_dropped":      o_dropped,
        "cg_lines_removed": cg_lines_removed,
    }
    return result, stats


# ─── CLI ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Trim Infor LN HTML Call Graph Profile files")
    parser.add_argument("input", help="Input HTML profile file")
    parser.add_argument("--out", help="Output file path")
    parser.add_argument("--min-query-time", type=float, default=0.05,
                        help="Drop Query Summary rows with total rtime < this (default 0.05s)")
    parser.add_argument("--min-flat-utime", type=float, default=0.001,
                        help="Drop Flat Profile rows with utime < this (default 0.001s)")
    parser.add_argument("--min-obj-cpu", type=float, default=0.05,
                        help="Drop Object Summary rows with cpu < this (default 0.05s)")
    parser.add_argument("--keep-rtime-graph", action="store_true",
                        help="Keep the rtime-sorted (1.1.x) call graph section (large!)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Estimate savings without writing output")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"ERROR: File not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    out_path = Path(args.out) if args.out else input_path.with_stem(input_path.stem + "_trimmed")

    print(f"Reading {input_path} …")
    with open(input_path, "r", encoding="utf-8", errors="replace") as f:
        lines = f.readlines()

    orig_bytes = input_path.stat().st_size

    print("Locating sections …")
    sections = find_sections(lines)
    found_names = list(sections.keys())
    print(f"  Sections found: {', '.join(found_names)}")

    print("Trimming …")
    result, stats = trim(lines, sections, args)

    print()
    print("=" * 55)
    print("  TRIM SUMMARY")
    print("=" * 55)
    print(f"  Original : {stats['original_lines']:>8,} lines  ({orig_bytes/1024/1024:.1f} MB)")
    print(f"  Output   : {stats['output_lines']:>8,} lines")
    print(f"  Saved    : {stats['lines_saved']:>8,} lines  ({stats['pct_saved']:.1f}%)")
    print()
    print(f"  Query Summary  : kept {stats['query_kept']}, dropped {stats['query_dropped']} rows (rtime < {args.min_query_time}s)")
    print(f"  Object Summary : kept {stats['obj_kept']}, dropped {stats['obj_dropped']} rows (cpu < {args.min_obj_cpu}s)")
    print(f"  Flat Profile   : kept {stats['flat_kept']}, dropped {stats['flat_dropped']} rows (utime < {args.min_flat_utime}s)")
    if not args.keep_rtime_graph:
        print(f"  rtime Call Graph removed: {stats['cg_lines_removed']:,} lines")
    print("=" * 55)

    if not args.dry_run:
        with open(out_path, "w", encoding="utf-8") as f:
            f.writelines(result)
            # If the rtime call graph was removed, the </BODY></HTML> tags at the
            # very end of the original file were also removed — add them back.
            if not args.keep_rtime_graph and stats["cg_lines_removed"] > 0:
                f.write("\n</BODY>\n</HTML>\n")
        out_bytes = out_path.stat().st_size
        print(f"\n  Output written: {out_path}")
        print(f"  File size: {out_bytes/1024/1024:.1f} MB  (was {orig_bytes/1024/1024:.1f} MB, -{100*(1-out_bytes/orig_bytes):.0f}%)")
    else:
        print("\n  (dry-run — no output written)")


if __name__ == "__main__":
    main()
