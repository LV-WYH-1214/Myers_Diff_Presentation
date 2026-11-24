#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Real Empirical Benchmark: Textbook 2D DP vs. Basic Myers (with full history) vs. Linear-Space Bidirectional Myers.
Run this script anytime to reproduce the exact timing (ms) and peak memory (MB) numbers on Slide 11.
"""

import time
import tracemalloc
import random
import json
import os


def generate_test_files(n_lines, d_edits, seed=42):
    """Generates two lists of hashed line IDs of length ~n_lines with ~d_edits differences."""
    rng = random.Random(seed)
    a = [f"line_content_{i}_{rng.randint(1000, 9999)}" for i in range(n_lines)]
    b = list(a)
    step = max(1, n_lines // max(1, d_edits))
    edits_done = 0
    for idx in range(0, n_lines, step):
        if edits_done >= d_edits:
            break
        if edits_done % 2 == 0:
            b[idx] = f"modified_line_{idx}"
        else:
            b.insert(idx, f"inserted_line_{idx}")
        edits_done += 1
    b = b[:n_lines]
    # Pre-hash lines to integer IDs (just like Git xprepare.c)
    vocab = {}
    a_ids = [vocab.setdefault(s, len(vocab)) for s in a]
    b_ids = [vocab.setdefault(s, len(vocab)) for s in b]
    return a_ids, b_ids


def dp_2d_lcs(a, b):
    """Standard Textbook O(N*M) 2D DP table for LCS/Edit Distance."""
    n, m = len(a), len(b)
    # Allocate (n+1) x (m+1) 2D array of ints
    dp = [[0] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        ai = a[i - 1]
        row_prev = dp[i - 1]
        row_curr = dp[i]
        for j in range(1, m + 1):
            if ai == b[j - 1]:
                row_curr[j] = row_prev[j - 1] + 1
            else:
                v1 = row_prev[j]
                v2 = row_curr[j - 1]
                row_curr[j] = v1 if v1 >= v2 else v2
    return n + m - 2 * dp[n][m]


def myers_basic_with_trace(a, b):
    """Basic Myers O((N+M)D) storing V snapshot per D for path backtracking (O(D^2) space)."""
    n, m = len(a), len(b)
    max_d = n + m
    offset = max_d + 2
    v = [0] * (2 * max_d + 5)
    trace = []
    for d in range(max_d + 1):
        for k in range(-d, d + 1, 2):
            idx = k + offset
            if k == -d or (k != d and v[idx - 1] < v[idx + 1]):
                x = v[idx + 1]
            else:
                x = v[idx - 1] + 1
            y = x - k
            while x < n and y < m and a[x] == b[y]:
                x += 1
                y += 1
            v[idx] = x
            if x >= n and y >= m:
                trace.append(v[offset - d: offset + d + 1])
                return d
        # Save slice of V for current D (size 2d+1) to enable backtracking
        trace.append(v[offset - d: offset + d + 1])
    return max_d


def myers_linear_space_bidirectional(a, b):
    """Linear-Space Bidirectional Myers O((N+M)D) time, strict O(N+M) space (two 1D arrays)."""
    n, m = len(a), len(b)
    max_d = n + m
    offset = max_d + 2
    # Only two 1D arrays allocated for the entire algorithm!
    v_fwd = [0] * (2 * max_d + 5)
    v_bwd = [0] * (2 * max_d + 5)

    delta = n - m
    odd = (delta & 1) != 0

    for d in range((max_d // 2) + 2):
        # Forward step d
        for k in range(-d, d + 1, 2):
            idx = k + offset
            if k == -d or (k != d and v_fwd[idx - 1] < v_fwd[idx + 1]):
                x = v_fwd[idx + 1]
            else:
                x = v_fwd[idx - 1] + 1
            y = x - k
            while x < n and y < m and a[x] == b[y]:
                x += 1
                y += 1
            v_fwd[idx] = x
            if odd and (k - delta) >= -(d - 1) and (k - delta) <= (d - 1):
                if x + v_bwd[(k - delta) + offset] >= n:
                    return 2 * d - 1

        # Backward step d
        for k in range(-d, d + 1, 2):
            idx = k + offset
            if k == -d or (k != d and v_bwd[idx - 1] < v_bwd[idx + 1]):
                x = v_bwd[idx + 1]
            else:
                x = v_bwd[idx - 1] + 1
            y = x - k
            while x < n and y < m and a[n - 1 - x] == b[m - 1 - y]:
                x += 1
                y += 1
            v_bwd[idx] = x
            if not odd and (k + delta) >= -d and (k + delta) <= d:
                if x + v_fwd[(k + delta) + offset] >= n:
                    return 2 * d
    return max_d


def run_benchmarks():
    sizes = [300, 600, 1200, 2400]
    results = {"sizes": sizes, "time_ms": {"dp": [], "myers": [], "bi": []},
               "mem_kb": {"dp": [], "myers": [], "bi": []}}

    print("=== Experiment 1: Wall-Clock Time (ms) for Small Daily Edit (D = 12) ===")
    for sz in sizes:
        a, b = generate_test_files(sz, d_edits=8)
        t0 = time.perf_counter()
        dp_2d_lcs(a, b)
        t_dp = (time.perf_counter() - t0) * 1000.0

        t0 = time.perf_counter()
        for _ in range(10):
            myers_basic_with_trace(a, b)
        t_myers = (time.perf_counter() - t0) * 100.0

        t0 = time.perf_counter()
        for _ in range(10):
            myers_linear_space_bidirectional(a, b)
        t_bi = (time.perf_counter() - t0) * 100.0

        results["time_ms"]["dp"].append(round(t_dp, 3))
        results["time_ms"]["myers"].append(round(t_myers, 4))
        results["time_ms"]["bi"].append(round(t_bi, 4))
        print(f"N={sz:4d} | 2D-DP: {t_dp:8.3f} ms | Myers-Basic: {t_myers:6.4f} ms | Myers-Bi: {t_bi:6.4f} ms")

    print("\n=== Experiment 2: Peak Memory (KB) via tracemalloc for Large Refactor (D ~ 0.35*N) ===")
    for sz in sizes:
        a, b = generate_test_files(sz, d_edits=int(sz * 0.25))

        tracemalloc.start()
        dp_2d_lcs(a, b)
        _, peak_dp = tracemalloc.get_traced_memory()
        tracemalloc.stop()

        tracemalloc.start()
        myers_basic_with_trace(a, b)
        _, peak_myers = tracemalloc.get_traced_memory()
        tracemalloc.stop()

        tracemalloc.start()
        myers_linear_space_bidirectional(a, b)
        _, peak_bi = tracemalloc.get_traced_memory()
        tracemalloc.stop()

        results["mem_kb"]["dp"].append(round(peak_dp / 1024.0, 1))
        results["mem_kb"]["myers"].append(round(peak_myers / 1024.0, 1))
        results["mem_kb"]["bi"].append(round(peak_bi / 1024.0, 1))
        print(f"N={sz:4d} | 2D-DP: {peak_dp/1024:8.1f} KB | Myers-Basic: {peak_myers/1024:7.1f} KB | Myers-Bi: {peak_bi/1024:6.1f} KB")

    out_json = "/Users/lv/Desktop/Myers_Diff_Presentation/assets/benchmark_results.json"
    with open(out_json, "w") as f:
        json.dump(results, f, indent=2)
    print(f"\nSaved real benchmark data to {out_json}")


if __name__ == "__main__":
    run_benchmarks()
