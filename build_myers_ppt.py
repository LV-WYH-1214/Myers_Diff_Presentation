#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generator for Myers Diff Algorithm Presentation (.pptx) + High-DPI Figures + Animated GIFs.
"""

import os
import json
import math
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.path import Path as MplPath
from PIL import Image

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Configure Matplotlib Chinese Font (use Hiragino Sans GB / Arial Unicode MS for complete Simplified Chinese coverage)
plt.rcParams['font.sans-serif'] = ['Hiragino Sans GB', 'Arial Unicode MS', 'Heiti TC', 'PingFang HK', 'sans-serif']
plt.rcParams['font.monospace'] = ['Hiragino Sans GB', 'Arial Unicode MS', 'Heiti TC', 'monospace']
plt.rcParams['axes.unicode_minus'] = False

OUT_DIR = "/Users/lv/Desktop/Myers_Diff_Presentation"
ASSETS_DIR = os.path.join(OUT_DIR, "assets")
os.makedirs(ASSETS_DIR, exist_ok=True)

# Classic example from Eugene Myers' 1986 paper: "An O(ND) Difference Algorithm and Its Variations"
A_STR = "ABCABBA"
B_STR = "CBABAC"
N = len(A_STR)  # 7
M = len(B_STR)  # 6


def compute_myers_trace(A, B):
    """
    Runs Myers' O((N+M)D) greedy edit-graph algorithm on strings A and B,
    recording every step for exact visualization and animation.
    Returns:
      history: list of dicts per D, containing V snapshot, explored edges, snakes
      final_path: list of (x1, y1, x2, y2, edge_type) from (0,0) to (N,M)
    """
    n, m = len(A), len(B)
    max_d = n + m
    v = {1: 0}
    history = []
    parent = {}  # (d, k) -> (prev_x, prev_y, mid_x, mid_y, end_x, end_y, move_type)

    found_d = None
    for d in range(0, max_d + 1):
        step_edges = []
        v_next = dict(v)
        for k in range(-d, d + 1, 2):
            if k == -d or (k != d and v.get(k - 1, -1) < v.get(k + 1, -1)):
                # Move DOWN from k + 1
                prev_k = k + 1
                px = v[prev_k]
                py = px - prev_k
                mx, my = px, py + 1
                move_type = 'down' if d > 0 else 'start'
            else:
                # Move RIGHT from k - 1
                prev_k = k - 1
                px = v[prev_k]
                py = px - prev_k
                mx, my = px + 1, py
                move_type = 'right'

            if d == 0:
                px, py, mx, my = 0, 0, 0, 0

            # Slide down diagonal snake
            ex, ey = mx, my
            snakes = []
            while ex < n and ey < m and A[ex] == B[ey]:
                snakes.append((ex, ey, ex + 1, ey + 1))
                ex += 1
                ey += 1

            v_next[k] = ex
            if 0 <= mx <= n and 0 <= my <= m:
                parent[(d, k)] = (px, py, mx, my, ex, ey, move_type, prev_k, snakes)
                step_edges.append({
                    'd': d, 'k': k,
                    'px': px, 'py': py,
                    'mx': mx, 'my': my,
                    'ex': ex, 'ey': ey,
                    'move_type': move_type,
                    'snakes': snakes,
                    'in_bounds': (0 <= ex <= n and 0 <= ey <= m)
                })

            if ex >= n and ey >= m and found_d is None:
                found_d = d
                break

        v = v_next
        history.append({
            'd': d,
            'v': dict(v),
            'edges': step_edges
        })
        if found_d is not None:
            break

    # Backtrack optimal path from (found_d, n - m)
    path_segments = []
    curr_k = n - m
    for d in range(found_d, -1, -1):
        px, py, mx, my, ex, ey, move_type, prev_k, snakes = parent[(d, curr_k)]
        # Add snakes in reverse
        for sx1, sy1, sx2, sy2 in reversed(snakes):
            path_segments.append((sx1, sy1, sx2, sy2, 'snake'))
        if d > 0:
            path_segments.append((px, py, mx, my, move_type))
        curr_k = prev_k

    path_segments.reverse()
    return history, path_segments, found_d


HISTORY, OPTIMAL_PATH, OPTIMAL_D = compute_myers_trace(A_STR, B_STR)


def draw_edit_grid_base(ax, title=""):
    """Draws the base Edit Graph grid for A='ABCABBA' (N=7) and B='CBABAC' (M=6)."""
    ax.set_xlim(-0.8, N + 0.8)
    ax.set_ylim(M + 0.8, -0.8)  # Invert Y so (0,0) is top-left
    ax.set_aspect('equal')
    ax.axis('off')

    # Draw grid lines
    for x in range(N + 1):
        ax.plot([x, x], [0, M], color='#CBD5E1', lw=1.2, zorder=1)
    for y in range(M + 1):
        ax.plot([0, N], [y, y], color='#CBD5E1', lw=1.2, zorder=1)

    # Draw all matching diagonals (Snakes)
    for x in range(N):
        for y in range(M):
            if A_STR[x] == B_STR[y]:
                ax.plot([x, x + 1], [y, y + 1], color='#F59E0B', lw=2.2, ls='--', alpha=0.75, zorder=2)

    # Draw vertices
    for x in range(N + 1):
        for y in range(M + 1):
            ax.scatter([x], [y], color='#64748B', s=28, zorder=3)

    # Labels for A (top horizontal edges)
    for x, ch in enumerate(A_STR):
        ax.text(x + 0.5, -0.38, ch, fontsize=13, fontweight='bold', color='#1E293B',
                ha='center', va='center',
                bbox=dict(boxstyle='round,pad=0.2', facecolor='#E2E8F0', edgecolor='none'))
    # Labels for B (left vertical edges)
    for y, ch in enumerate(B_STR):
        ax.text(-0.42, y + 0.5, ch, fontsize=13, fontweight='bold', color='#1E293B',
                ha='center', va='center',
                bbox=dict(boxstyle='round,pad=0.2', facecolor='#E2E8F0', edgecolor='none'))

    # Coordinate numbers
    for x in range(N + 1):
        ax.text(x, -0.08, str(x), fontsize=8.5, color='#64748B', ha='center', va='bottom')
    for y in range(M + 1):
        ax.text(-0.1, y, str(y), fontsize=8.5, color='#64748B', ha='right', va='center')

    # Highlight (0,0) and (N,M)
    ax.scatter([0], [0], color='#2563EB', s=90, zorder=6)
    ax.scatter([N], [M], color='#10B981', s=110, marker='*', zorder=6)
    if title:
        ax.set_title(title, fontsize=13, fontweight='bold', color='#0F172A', pad=12)


# ==============================================================================
# FIGURE 1: Pain Points of Standard 2D LCS DP vs Myers O(ND)
# ==============================================================================
def make_fig1_pain_points():
    fig, axes = plt.subplots(1, 2, figsize=(12.2, 3.85), dpi=240)
    fig.patch.set_facecolor('#FFFFFF')

    # Left subplot: 2D DP full table vs Myers narrow band on 40x40 matrix
    ax = axes[0]
    grid_n = 40
    mat_dp = np.ones((grid_n, grid_n)) * 0.25
    # Show Myers explored region (narrow diamond around diagonal)
    for r in range(grid_n):
        for c in range(grid_n):
            if abs(r - c) <= 4 and (r + c) <= grid_n * 1.85:
                mat_dp[r, c] = 0.9

    ax.imshow(mat_dp, cmap='Blues', vmin=0, vmax=1, aspect='equal')
    ax.plot([0, grid_n - 1], [0, grid_n - 1], color='#F59E0B', lw=2.5, ls='--', label='主对角线 k=0 (代码未改动区)')
    ax.set_title("痛点 1：课本二维 DP 像“压路机”填满 N×M 全表\nvs. Myers 像“激光束”仅搜对角线窄带 O((N+M)D)",
                 fontsize=11.0, fontweight='bold', color='#0F172A', pad=8)
    ax.set_xlabel("原文件 A 行号 (N 行)", fontsize=9.5, color='#334155')
    ax.set_ylabel("新文件 B 行号 (M 行)", fontsize=9.5, color='#334155')
    ax.text(grid_n * 0.18, grid_n * 0.78, "课本 DP 盲目计算区\nN=2400 时填满 576 万格\n本机实测耗时 275.3 ms",
            fontsize=8.8, color='#1E3A8A', bbox=dict(boxstyle='round,pad=0.35', facecolor='#EFF6FF', edgecolor='#93C5FD'))
    ax.text(grid_n * 0.44, grid_n * 0.24, "Myers 搜索锥区 (D=12)\n仅访问 ~0.5% 节点!\n实测仅需 0.13 ms (快 2,131 倍)",
            fontsize=8.8, fontweight='bold', color='#FFFFFF',
            bbox=dict(boxstyle='round,pad=0.35', facecolor='#2563EB', edgecolor='#1D4ED8'))
    ax.legend(loc='lower right', fontsize=8.0)

    # Right subplot: Bad LCS diff vs Intuitive Myers/Git diff
    ax2 = axes[1]
    ax2.axis('off')
    ax2.set_title("痛点 2：普通 LCS 容易把孤立括号 `}` 强行配对，产出破碎 Diff",
                  fontsize=11.0, fontweight='bold', color='#0F172A', pad=8)

    # Draw two code diff boxes
    box_bad = mpatches.FancyBboxPatch((0.02, 0.52), 0.96, 0.44, boxstyle="round,pad=0.02",
                                      facecolor='#FEF2F2', edgecolor='#FCA5A5', lw=1.5)
    ax2.add_patch(box_bad)
    ax2.text(0.04, 0.92, "[反例] 普通 LCS 回溯：把旧函数结尾的 `}` 强行配给新函数",
             fontsize=9.2, fontweight='bold', color='#991B1B', va='top')
    bad_code = (
        "  void foo() {\n"
        "-     old_logic();\n"
        "+     new_helper() { step1(); }\n"
        "  }              <-- 错位匹配！代码被切碎\n"
        "+ void foo() { new_logic(); }"
    )
    ax2.text(0.05, 0.80, bad_code, fontsize=8.4, color='#7F1D1D', va='top', linespacing=1.32)

    box_good = mpatches.FancyBboxPatch((0.02, 0.02), 0.96, 0.44, boxstyle="round,pad=0.02",
                                       facecolor='#ECFDF5', edgecolor='#6EE7B7', lw=1.5)
    ax2.add_patch(box_good)
    ax2.text(0.04, 0.42, "[正例] Myers 贪心滑行：优先连续匹配，输出整洁可读代码块",
             fontsize=9.2, fontweight='bold', color='#065F46', va='top')
    good_code = (
        "+ void new_helper() { step1(); }\n"
        "  void foo() {   <-- 新增函数整块在前，修改逻辑在后\n"
        "-     old_logic();\n"
        "+     new_logic();\n"
        "  }"
    )
    ax2.text(0.05, 0.31, good_code, fontsize=8.4, color='#064E3B', va='top', linespacing=1.30)

    plt.tight_layout()
    path = os.path.join(ASSETS_DIR, "fig1_dp_vs_myers_pain.png")
    plt.savefig(path, dpi=240)
    plt.close(fig)
    return path


# ==============================================================================
# FIGURE 2: Edit Graph Modeling & Optimal Path (A="ABCABBA", B="CBABAC")
# ==============================================================================
def make_fig2_edit_graph():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11.5, 5.0), dpi=240,
                                   gridspec_kw={'width_ratios': [1.35, 1.0]})
    fig.patch.set_facecolor('#FFFFFF')

    draw_edit_grid_base(ax1, "编辑图 (Edit Graph)：A = \"ABCABBA\" → B = \"CBABAC\"")

    # Overlay optimal path on ax1
    for x1, y1, x2, y2, etype in OPTIMAL_PATH:
        if etype == 'right':
            ax1.annotate("", xy=(x2, y2), xytext=(x1, y1),
                         arrowprops=dict(arrowstyle="->", color="#EF4444", lw=3.0, shrinkA=3, shrinkB=3), zorder=5)
        elif etype == 'down':
            ax1.annotate("", xy=(x2, y2), xytext=(x1, y1),
                         arrowprops=dict(arrowstyle="->", color="#10B981", lw=3.0, shrinkA=3, shrinkB=3), zorder=5)
        elif etype == 'snake':
            ax1.annotate("", xy=(x2, y2), xytext=(x1, y1),
                         arrowprops=dict(arrowstyle="->", color="#2563EB", lw=3.5, shrinkA=2, shrinkB=2), zorder=5)

    # Right panel: Rule explanation & Diff translation
    ax2.axis('off')
    rules_text = [
        ("1. 图的顶点与坐标", "网格大小为 (N+1)×(M+1)，起点 (0,0)，终点 (7,6)。"),
        ("2. 三种移动方向（边的权重）", ""),
        ("   • 向右走 (x→x+1, y)", "删除 A[x]（红色箭头，编辑代价 +1）"),
        ("   • 向下走 (x, y→y+1)", "插入 B[y]（绿色箭头，编辑代价 +1）"),
        ("   • 斜向对角线 (x+1, y+1)", "当 A[x]=B[y] 时保留字符（蓝色斜线 Snake，代价 0！）"),
        ("3. 路径与 Diff 脚本一一对应", "图中高亮的最短路径 (D = 5 步非对角线边) 对应："),
    ]
    y_pos = 0.96
    for header, desc in rules_text:
        ax2.text(0.02, y_pos, header, fontsize=10.2, fontweight='bold', color='#0F172A', va='top')
        y_pos -= 0.05
        if desc:
            ax2.text(0.05, y_pos, desc, fontsize=9.2, color='#334155', va='top')
            y_pos -= 0.058

    # Draw visual diff badges
    diff_ops = [
        ("-A", "#FEE2E2", "#991B1B"),
        ("-B", "#FEE2E2", "#991B1B"),
        (" C", "#DBEAFE", "#1E40AF"),
        ("+B", "#D1FAE5", "#065F46"),
        (" A", "#DBEAFE", "#1E40AF"),
        (" B", "#DBEAFE", "#1E40AF"),
        ("-B", "#FEE2E2", "#991B1B"),
        (" A", "#DBEAFE", "#1E40AF"),
        ("+C", "#D1FAE5", "#065F46"),
    ]
    x_pos = 0.03
    for label, bg, fg in diff_ops:
        ax2.text(x_pos, 0.25, label, fontsize=10.0, fontweight='bold',
                 color=fg, bbox=dict(boxstyle='round,pad=0.25', facecolor=bg, edgecolor=fg, lw=1.2))
        x_pos += 0.105

    ax2.text(0.02, 0.16,
             "核心等价定理：\n"
             "求最小编辑距离 D (或最长公共子序列 LCS)\n"
             "<==> 在编辑图中寻找从 (0,0) 到 (N,M) 经过最少非对角线边的路径！\n"
             "公式恒等式：D = N + M - 2·|LCS|  (本例 D = 7 + 6 - 2×4 = 5)",
             fontsize=9.2, color='#1E293B',
             bbox=dict(boxstyle='round,pad=0.4', facecolor='#FEF3C7', edgecolor='#F59E0B', lw=1.5),
             va='top')

    plt.tight_layout()
    path = os.path.join(ASSETS_DIR, "fig2_edit_graph_intro.png")
    plt.savefig(path, dpi=240)
    plt.close(fig)
    return path


# ==============================================================================
# FIGURE 3: Coordinate Transformation (x,y) -> (D, k = x - y)
# ==============================================================================
def make_fig3_k_diagonals():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11.5, 4.9), dpi=240,
                                   gridspec_kw={'width_ratios': [1.25, 1.0]})
    fig.patch.set_facecolor('#FFFFFF')

    draw_edit_grid_base(ax1, "对角线坐标系 k = x - y (k ∈ [-M, N])")

    # Draw k-diagonals
    k_colors = {
        0: '#2563EB', 1: '#7C3AED', -1: '#7C3AED',
        2: '#059669', -2: '#059669', 3: '#D97706', -3: '#D97706'
    }
    for k in range(-M, N + 1):
        pts = [(x, x - k) for x in range(N + 1) if 0 <= x - k <= M]
        if len(pts) >= 2:
            xs, ys = zip(*pts)
            col = k_colors.get(k, '#94A3B8')
            lw = 2.2 if abs(k) <= 2 else 1.0
            ax1.plot(xs, ys, color=col, lw=lw, ls=':', alpha=0.7, zorder=1)
            # Label k at the end of diagonal
            lx, ly = pts[-1]
            ax1.text(lx + 0.22, ly + 0.22, f"k={k}", fontsize=8.5, fontweight='bold', color=col)

    ax2.axis('off')
    ax2.set_title("为什么定义 k = x - y 能实现状态降维？", fontsize=12, fontweight='bold', color='#0F172A', pad=10)

    cards = [
        ("性质 1：斜滑不换“车道 k”",
         "• 免费滑梯 (x+1, y+1) 保持 k = x-y 恒定不变；\n"
         "• 右移删除 k -> k+1；下移插入 k -> k-1。",
         "#EFF6FF", "#3B82F6"),
        ("性质 2：奇偶性引理（查一半车道）",
         "• 起点 k=0，每走 1 步代价边 k 必 ±1；\n"
         "• 走 D 步后 k 必与 D 同奇偶 ==> 第 D 轮仅查 D+1 条线。",
         "#F5F3FF", "#8B5CF6"),
        ("性质 3：每条车道只留最远“车头” V[k]",
         "• 同一条线 y = x - k 上，x 越大离终点越近；\n"
         "• V[k] = 走 D 步到达车道 k 的【最大横坐标 x】。",
         "#ECFDF5", "#10B981"),
    ]
    y_top = 0.92
    for title, body, bg, border in cards:
        ax2.text(0.02, y_top, f"【{title}】\n{body}", fontsize=9.3, color='#1E293B', va='top',
                 linespacing=1.45,
                 bbox=dict(boxstyle='round,pad=0.45', facecolor=bg, edgecolor=border, lw=1.5))
        y_top -= 0.31

    plt.tight_layout()
    path = os.path.join(ASSETS_DIR, "fig3_k_diagonals.png")
    plt.savefig(path, dpi=240)
    plt.close(fig)
    return path


# ==============================================================================
# FIGURE 4: State Transition Diagram (How V[k] is computed at step D)
# ==============================================================================
def make_fig4_state_transition():
    fig, ax = plt.subplots(figsize=(7.2, 5.2), dpi=240)
    fig.patch.set_facecolor('#FFFFFF')
    ax.set_xlim(-0.5, 5.5)
    ax.set_ylim(5.5, -0.5)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title("状态转移网格图：第 D 步如何更新对角线 k 上的最远点 V[k]？",
                 fontsize=12.5, fontweight='bold', color='#0F172A', pad=12)

    # Draw local grid
    for i in range(6):
        ax.plot([i, i], [0, 5], color='#E2E8F0', lw=1.2, zorder=1)
        ax.plot([0, 5], [i, i], color='#E2E8F0', lw=1.2, zorder=1)

    # Draw 3 diagonals: k+1, k, k-1
    # Let k=0 be main diagonal (x=y), k+1 be y=x-1, k-1 be y=x+1
    ax.plot([0, 5], [0, 5], color='#2563EB', lw=2.2, ls=':', label='目标对角线 k')
    ax.plot([1, 5], [0, 4], color='#10B981', lw=2.0, ls=':', label='上方对角线 k+1')
    ax.plot([0, 4], [1, 5], color='#EF4444', lw=2.0, ls=':', label='下方对角线 k-1')

    ax.text(5.05, 5.15, "对角线 k", color='#2563EB', fontweight='bold', fontsize=9.5)
    ax.text(5.05, 4.05, "对角线 k+1", color='#10B981', fontweight='bold', fontsize=9.5)
    ax.text(3.6, 5.25, "对角线 k-1", color='#EF4444', fontweight='bold', fontsize=9.5)

    # Point on k+1: (2, 1) -> move DOWN to (2, 2) on k
    ax.scatter([2], [1], color='#10B981', s=120, zorder=5)
    ax.text(2.0, 0.55, "V[k+1] = 2\n(D-1 步终点)", fontsize=8.8, fontweight='bold', color='#065F46', ha='center',
            bbox=dict(boxstyle='round,pad=0.2', facecolor='#D1FAE5', edgecolor='#10B981'))
    ax.annotate("", xy=(2, 2), xytext=(2, 1),
                arrowprops=dict(arrowstyle="->", color="#10B981", lw=3.0), zorder=4)
    ax.text(2.2, 1.45, "向下插入\nx 不变 (x=2)", fontsize=8.2, color='#065F46', fontweight='bold',
            bbox=dict(boxstyle='round,pad=0.15', facecolor='#FFFFFF', edgecolor='none', alpha=0.85))

    # Point on k-1: (1, 2) -> move RIGHT to (2, 2) on k
    ax.scatter([1], [2], color='#EF4444', s=120, zorder=5)
    ax.text(0.32, 2.0, "V[k-1] = 1\n(D-1 步终点)", fontsize=8.8, fontweight='bold', color='#991B1B', ha='center',
            bbox=dict(boxstyle='round,pad=0.2', facecolor='#FEE2E2', edgecolor='#EF4444'))
    ax.annotate("", xy=(2, 2), xytext=(1, 2),
                arrowprops=dict(arrowstyle="->", color="#EF4444", lw=3.0), zorder=4)
    ax.text(1.42, 2.48, "向右删除\nx = V[k-1]+1", fontsize=8.2, color='#991B1B', fontweight='bold', ha='center',
            bbox=dict(boxstyle='round,pad=0.15', facecolor='#FFFFFF', edgecolor='none', alpha=0.85))

    # Landing point (2,2) -> slide along snake to (4,4)
    ax.scatter([2], [2], color='#F59E0B', s=130, zorder=6)
    ax.text(2.55, 2.05, "① 择优落脚点\nx = max(V[k-1]+1, V[k+1])", fontsize=8.8, fontweight='bold', color='#B45309',
            bbox=dict(boxstyle='round,pad=0.2', facecolor='#FFFBEB', edgecolor='#F59E0B', lw=1.0))

    ax.annotate("", xy=(4, 4), xytext=(2, 2),
                arrowprops=dict(arrowstyle="->", color="#2563EB", lw=4.0), zorder=5)
    ax.text(3.35, 2.95, "② 顺对角线免费滑到底 (Snake)!\nwhile A[x]==B[y]: x++, y++",
            fontsize=8.8, fontweight='bold', color='#1D4ED8',
            bbox=dict(boxstyle='round,pad=0.25', facecolor='#DBEAFE', edgecolor='#2563EB'))

    ax.scatter([4], [4], color='#2563EB', s=160, marker='*', zorder=7)
    ax.text(4.0, 4.48, "更新 V[k] = 4", fontsize=9.5, fontweight='bold', color='#1E40AF', ha='center',
            bbox=dict(boxstyle='round,pad=0.25', facecolor='#EFF6FF', edgecolor='#2563EB', lw=1.5))

    plt.tight_layout()
    path = os.path.join(ASSETS_DIR, "fig4_state_transition.png")
    plt.savefig(path, dpi=240)
    plt.close(fig)
    return path


# ==============================================================================
# FIGURE 5: 6-Panel Step-by-Step Trace (D = 0, 1, 2, 3, 4, 5) + Animated GIF
# ==============================================================================
def _draw_trace_panel_group(d_list, panel_titles, figsize, dpi, out_filename, badge_fontsize=7.6):
    """Render a crisp 1x3 row of Myers trace snapshots for the requested D rounds."""
    fig, axes = plt.subplots(1, len(d_list), figsize=figsize, dpi=dpi)
    fig.patch.set_facecolor('#FFFFFF')
    if len(d_list) == 1:
        axes = [axes]

    for idx, d_val in enumerate(d_list):
        ax = axes[idx]
        draw_edit_grid_base(ax, "")
        ax.set_title(panel_titles[idx], fontsize=9.6 if len(d_list) == 3 and figsize[0] < 9 else 10.6,
                     fontweight='bold', color='#0F172A', pad=6)
        for prev_d in range(d_val):
            for e in HISTORY[prev_d]['edges']:
                if prev_d > 0:
                    ax.plot([e['px'], e['mx']], [e['py'], e['my']], color='#94A3B8', lw=1.4, alpha=0.45, zorder=3)
                for sx1, sy1, sx2, sy2 in e['snakes']:
                    ax.plot([sx1, sx2], [sy1, sy2], color='#60A5FA', lw=1.8, alpha=0.5, zorder=3)

        for e in HISTORY[d_val]['edges']:
            if d_val > 0:
                col = '#EF4444' if e['move_type'] == 'right' else '#10B981'
                ax.annotate("", xy=(e['mx'], e['my']), xytext=(e['px'], e['py']),
                            arrowprops=dict(arrowstyle="->", color=col, lw=2.3, shrinkA=2, shrinkB=2), zorder=5)
            for sx1, sy1, sx2, sy2 in e['snakes']:
                ax.annotate("", xy=(sx2, sy2), xytext=(sx1, sy1),
                            arrowprops=dict(arrowstyle="->", color='#2563EB', lw=2.9, shrinkA=1, shrinkB=1), zorder=5)
            if 0 <= e['ex'] <= N and 0 <= e['ey'] <= M:
                ax.scatter([e['ex']], [e['ey']], color='#F59E0B', edgecolor='#B45309', s=68, zorder=6)
                dy_off = 0.32 if e['ey'] == 0 else -0.18
                ax.text(e['ex'] + 0.14, e['ey'] + dy_off, f"V[{e['k']:+d}]={e['ex']}",
                        fontsize=badge_fontsize, fontweight='bold', color='#92400E',
                        bbox=dict(boxstyle='round,pad=0.10', facecolor='#FEF3C7', edgecolor='#F59E0B', lw=0.7, alpha=0.92),
                        zorder=7)

        if d_val == OPTIMAL_D:
            for x1, y1, x2, y2, _ in OPTIMAL_PATH:
                ax.plot([x1, x2], [y1, y2], color='#F59E0B', lw=5.0, alpha=0.45, zorder=4)

    plt.tight_layout()
    out_path = os.path.join(ASSETS_DIR, out_filename)
    plt.savefig(out_path, dpi=dpi)
    plt.close(fig)
    return out_path


def make_fig5_and_gif():
    d_descriptions = [
        "D = 0：起点 (0,0)，停在 V[0]=0",
        "D = 1：向右至 (1,0)、向下至 (0,1)",
        "D = 2：触发 3 条对角滑梯，达 (3,1)/(2,4)",
        "D = 3：顺长滑梯推进至 (5,4) 与 (4,5)",
        "D = 4：最远达 (7,5) 与 (4,6)，距终点一步",
        "D = 5：命中终点 (7,6) ★ 最短路径完成！"
    ]

    # 1a. Large 1x3 figure for Part 1 Slide 6 (D = 0 -> 1 -> 2)
    f5a_path = _draw_trace_panel_group(
        [0, 1, 2],
        ["D = 0：起点 (0,0) 无匹配滑梯", "D = 1：分别尝试删/插迈出 1 步", "D = 2：首次触发免费对角滑梯!"],
        figsize=(12.2, 3.90), dpi=240, out_filename="fig5a_trace_d0_to_d2.png", badge_fontsize=7.8
    )

    # 1b. Large 1x3 figure for Part 1 Slide 7 (D = 3 -> 4 -> 5)
    f5b_path = _draw_trace_panel_group(
        [3, 4, 5],
        ["D = 3：长滑梯直冲 (5,4)", "D = 4：推进到 (7,5)", "D = 5：命中 (7,6) 完成 ★"],
        figsize=(7.75, 3.64), dpi=240, out_filename="fig5b_trace_d3_to_d5.png", badge_fontsize=6.5
    )

    # 2. Generate high-res Animated GIF (myers_edit_graph_anim.gif) with synchronized V[k] bar!
    frames = []
    k_cols = list(range(-5, 6))  # k from -5 to +5
    for d_val in range(7):  # 0..5 + final path highlight frame
        f_fig, (f_ax, v_ax) = plt.subplots(2, 1, figsize=(6.3, 5.1), dpi=150,
                                           gridspec_kw={'height_ratios': [4.1, 1.0]})
        f_fig.patch.set_facecolor('#FFFFFF')
        if d_val <= 5:
            title_str = f"Myers 动态搜索 — 第 D = {d_val} 步\n{d_descriptions[d_val]}"
        else:
            title_str = "★ 搜索完成 (D = 5)！回溯生成最短编辑脚本 (SES)：\n-A  -B   C  +B   A   B  -B   A  +C"
        draw_edit_grid_base(f_ax, title_str)

        max_step = min(d_val, 5)
        for prev_d in range(max_step):
            for e in HISTORY[prev_d]['edges']:
                if prev_d > 0:
                    f_ax.plot([e['px'], e['mx']], [e['py'], e['my']], color='#94A3B8', lw=1.8, alpha=0.45, zorder=3)
                for sx1, sy1, sx2, sy2 in e['snakes']:
                    f_ax.plot([sx1, sx2], [sy1, sy2], color='#93C5FD', lw=2.2, alpha=0.5, zorder=3)

        active_ks = set()
        if d_val <= 5:
            for e in HISTORY[d_val]['edges']:
                active_ks.add(e['k'])
                if d_val > 0:
                    col = '#EF4444' if e['move_type'] == 'right' else '#10B981'
                    f_ax.annotate("", xy=(e['mx'], e['my']), xytext=(e['px'], e['py']),
                                  arrowprops=dict(arrowstyle="->", color=col, lw=2.8, shrinkA=2, shrinkB=2), zorder=5)
                for sx1, sy1, sx2, sy2 in e['snakes']:
                    f_ax.annotate("", xy=(sx2, sy2), xytext=(sx1, sy1),
                                  arrowprops=dict(arrowstyle="->", color='#2563EB', lw=3.3, shrinkA=1, shrinkB=1), zorder=5)
                if 0 <= e['ex'] <= N and 0 <= e['ey'] <= M:
                    f_ax.scatter([e['ex']], [e['ey']], color='#F59E0B', edgecolor='#B45309', s=100, zorder=6)
                    f_ax.text(e['ex'] + 0.15, e['ey'] - 0.18, f"k={e['k']}", fontsize=8.0,
                              fontweight='bold', color='#B45309')
        else:
            active_ks.add(1)
            for x1, y1, x2, y2, etype in OPTIMAL_PATH:
                col = '#EF4444' if etype == 'right' else ('#10B981' if etype == 'down' else '#2563EB')
                f_ax.plot([x1, x2], [y1, y2], color='#FDE047', lw=7.5, alpha=0.7, zorder=4)
                f_ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                              arrowprops=dict(arrowstyle="->", color=col, lw=3.2, shrinkA=2, shrinkB=2), zorder=6)

        # Draw synchronized V[k] 1D array bar at bottom of frame
        v_ax.set_xlim(-5.6, 5.6)
        v_ax.set_ylim(0, 1.8)
        v_ax.axis('off')
        v_snapshot = HISTORY[max_step]['v']
        v_ax.text(0, 1.58, f"一维状态数组 V[k] 原地滚动更新 (橙色高亮为本轮 D={max_step} 更新的对角线 k)：",
                  fontsize=8.2, fontweight='bold', color='#1E293B', ha='center')
        for k_val in k_cols:
            is_curr = (k_val in active_ks)
            has_val = (k_val in v_snapshot) and (abs(k_val) <= max_step)
            bg_c = '#FEF3C7' if is_curr else ('#EFF6FF' if has_val else '#F8FAFC')
            bd_c = '#F59E0B' if is_curr else ('#93C5FD' if has_val else '#E2E8F0')
            rect = mpatches.FancyBboxPatch((k_val - 0.42, 0.12), 0.84, 1.15, boxstyle="round,pad=0.04",
                                           facecolor=bg_c, edgecolor=bd_c, lw=1.5 if is_curr else 1.0)
            v_ax.add_patch(rect)
            v_ax.text(k_val, 0.96, f"k={k_val}", fontsize=7.2, color='#64748B', ha='center', va='center')
            val_str = str(v_snapshot[k_val]) if has_val else "·"
            v_ax.text(k_val, 0.44, val_str, fontsize=9.5, fontweight='bold',
                      color='#B45309' if is_curr else ('#1E40AF' if has_val else '#CBD5E1'),
                      ha='center', va='center')

        plt.tight_layout()
        tmp_png = os.path.join(ASSETS_DIR, f"_tmp_frame_{d_val}.png")
        plt.savefig(tmp_png, dpi=150)
        plt.close(f_fig)
        frames.append(Image.open(tmp_png).convert("RGB"))

    gif_path = os.path.join(OUT_DIR, "myers_edit_graph_anim.gif")
    frames[0].save(
        gif_path,
        save_all=True,
        append_images=frames[1:],
        duration=[900, 1100, 1100, 1100, 1100, 1300, 2200],
        loop=0
    )
    return f5a_path, f5b_path, gif_path


# ==============================================================================
# FIGURE 6: Race Comparison Animation (2D DP vs Myers O(ND))
# ==============================================================================
def make_race_gif():
    grid_sz = 24
    # Construct two 24-char strings that differ in only D=4 places
    frames = []
    total_steps = 8
    for step in range(1, total_steps + 1):
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(7.6, 3.65), dpi=140)
        fig.patch.set_facecolor('#FFFFFF')

        # Left: Standard 2D DP fills row by row
        dp_grid = np.zeros((grid_sz, grid_sz))
        rows_done = int((step / total_steps) * grid_sz)
        dp_grid[:rows_done, :] = 0.75
        ax1.imshow(dp_grid, cmap='Reds', vmin=0, vmax=1)
        ax1.set_title(f"课本二维 DP (压路机全表扫描)\n已算：{rows_done * grid_sz} / {grid_sz*grid_sz} 格 ({int(rows_done/grid_sz*100)}%)",
                      fontsize=10.0, fontweight='bold', color='#991B1B')
        ax1.set_xticks([])
        ax1.set_yticks([])

        # Right: Myers explores narrow diagonal cone and finishes by step 4!
        myers_grid = np.zeros((grid_sz, grid_sz))
        reach = min(grid_sz, int((step / 4.0) * grid_sz))
        for r in range(reach):
            for c in range(reach):
                if abs(r - c) <= min(2, step // 2 + 1):
                    myers_grid[r, c] = 0.85
        cells_myers = int(np.sum(myers_grid > 0))
        status_str = "★ 已命中终点！" if step >= 4 else "沿对角线高速滑行..."
        ax2.imshow(myers_grid, cmap='Blues', vmin=0, vmax=1)
        ax2.set_title(f"Myers 算法 (激光束对角线聚焦)\n已访问：{cells_myers} 格 ({status_str})",
                      fontsize=10.0, fontweight='bold', color='#1E40AF')
        ax2.set_xticks([])
        ax2.set_yticks([])

        plt.tight_layout()
        tmp_png = os.path.join(ASSETS_DIR, f"_tmp_race_{step}.png")
        plt.savefig(tmp_png, dpi=140)
        plt.close(fig)
        frames.append(Image.open(tmp_png).convert("RGB"))

    race_gif_path = os.path.join(OUT_DIR, "lcs_vs_myers_race.gif")
    frames[0].save(
        race_gif_path,
        save_all=True,
        append_images=frames[1:],
        duration=550,
        loop=0
    )
    # Also save last frame as static image
    static_race = os.path.join(ASSETS_DIR, "fig_race_static.png")
    frames[-1].save(static_race)
    return static_race, race_gif_path


# ==============================================================================
# FIGURE 7: Bidirectional Search + Divide & Conquer (Linear Space O(N+M))
# ==============================================================================
def make_fig6_bidirectional():
    fig, ax = plt.subplots(figsize=(12.2, 3.85), dpi=240)
    fig.patch.set_facecolor('#FFFFFF')
    ax.set_xlim(0, 11.2)
    ax.set_ylim(5.2, 0)
    ax.axis('off')
    ax.set_title("挖隧道模型：正向与反向各推进 D/2 步在中间会师（寻找“中间蛇”），将空间压至线性 O(N+M)",
                 fontsize=11.5, fontweight='bold', color='#0F172A', pad=8)

    # Left grid diagram (x: 0.4..7.4, y: 0.4..4.8)
    outer = mpatches.Rectangle((0.6, 0.45), 6.8, 4.3, fill=True, facecolor='#F8FAFC', edgecolor='#94A3B8', lw=1.8)
    ax.add_patch(outer)

    # Top-left subproblem box
    sub1 = mpatches.Rectangle((0.6, 0.45), 2.8, 1.8, fill=True, facecolor='#EFF6FF', edgecolor='#3B82F6', lw=2.0, ls='--')
    ax.add_patch(sub1)
    ax.text(2.0, 1.50, "左上方子问题 (规模减半)\n递归求解，复用 V_fwd[k]", fontsize=8.8, fontweight='bold',
            color='#1D4ED8', ha='center', va='center',
            bbox=dict(boxstyle='round,pad=0.25', facecolor='#FFFFFF', edgecolor='#93C5FD', lw=1.2, alpha=0.95), zorder=5)

    # Bottom-right subproblem box
    sub2 = mpatches.Rectangle((4.6, 2.95), 2.8, 1.8, fill=True, facecolor='#ECFDF5', edgecolor='#10B981', lw=2.0, ls='--')
    ax.add_patch(sub2)
    ax.text(6.0, 3.70, "右下方子问题 (规模减半)\n递归求解，复用 V_bwd[k]", fontsize=8.8, fontweight='bold',
            color='#065F46', ha='center', va='center',
            bbox=dict(boxstyle='round,pad=0.25', facecolor='#FFFFFF', edgecolor='#6EE7B7', lw=1.2, alpha=0.95), zorder=5)

    # Middle Snake from (3.4, 2.25) to (4.6, 2.95)
    ax.annotate("", xy=(4.6, 2.95), xytext=(3.4, 2.25),
                arrowprops=dict(arrowstyle="->", color="#F59E0B", lw=5.0), zorder=6)
    ax.scatter([3.4, 4.6], [2.25, 2.95], color='#D97706', s=100, zorder=7)
    ax.text(4.0, 1.75, "中间蛇 (Middle Snake)\n正反向走 D/2 步在黄金重叠段会师！",
            fontsize=9.2, fontweight='bold', color='#B45309', ha='center',
            bbox=dict(boxstyle='round,pad=0.28', facecolor='#FEF3C7', edgecolor='#F59E0B', lw=1.5))

    # Forward & Backward arrows
    ax.annotate("", xy=(3.32, 2.20), xytext=(0.68, 0.50),
                arrowprops=dict(arrowstyle="->", color="#2563EB", lw=2.4, ls=':'), zorder=4)
    ax.text(1.75, 0.72, "正向挖进 (从 (0,0) 走 D/2 步)", fontsize=8.5, color='#1E40AF', fontweight='bold')

    ax.annotate("", xy=(4.68, 3.00), xytext=(7.32, 4.70),
                arrowprops=dict(arrowstyle="->", color="#059669", lw=2.4, ls=':'), zorder=4)
    ax.text(5.15, 4.50, "反向挖进 (从 (N,M) 走 D/2 步)", fontsize=8.5, color='#065F46', fontweight='bold')

    ax.scatter([0.6], [0.45], color='#2563EB', s=90, zorder=7)
    ax.text(0.45, 0.30, "(0, 0)", fontsize=9.0, fontweight='bold', color='#1E293B')
    ax.scatter([7.4], [4.75], color='#10B981', s=110, marker='*', zorder=7)
    ax.text(7.45, 4.95, "(N, M)", fontsize=9.0, fontweight='bold', color='#1E293B')

    # Right side classmate analogy + Hirschberg intuition box
    ax.text(7.75, 0.65,
            "【直观几何模型：两头挖隧道会师】\n"
            "• 痛点：丢掉旧 V[k] 后无法直接回溯路径；\n"
            "  存全套快照又要退化回 O(D^2) 内存。\n\n"
            "• 分治解法（Hirschberg 思想）：\n"
            "  1. 两队从 (0,0) 与 (N,M) 各推进 D/2 步\n"
            "  2. 在公共对角线会师，锁定【中间蛇】\n"
            "  3. 擦掉旧 V[k] 复用，切分左上/右下子问题\n"
            "  ==> 空间压至严格线性 O(N+M)！",
            fontsize=9.0, color='#1E293B', va='top', linespacing=1.48,
            bbox=dict(boxstyle='round,pad=0.45', facecolor='#EFF6FF', edgecolor='#3B82F6', lw=1.4))

    plt.tight_layout()
    path = os.path.join(ASSETS_DIR, "fig6_bidirectional_middle_snake.png")
    plt.savefig(path, dpi=240)
    plt.close(fig)
    return path


# ==============================================================================
# FIGURE 8: Worst Case Degradation & Git Patience / Histogram Diff
# ==============================================================================
def make_fig7_worst_case():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12.2, 3.85), dpi=240)
    fig.patch.set_facecolor('#FFFFFF')

    # Left: Complexity curve as D grows from 1 to N+M
    d_vals = np.linspace(1, 100, 100)
    dp_cost = np.ones_like(d_vals) * 10000
    myers_cost = 200 * d_vals * 0.9
    myers_cutoff = np.minimum(myers_cost, 3200)

    ax1.plot(d_vals, dp_cost, color='#EF4444', lw=2.4, ls='--', label='课本二维 DP: Θ(N·M) 恒定高开销')
    ax1.plot(d_vals, myers_cost, color='#2563EB', lw=2.6, label='标准 Myers: O((N+M)D) 随差异线性增长')
    ax1.plot(d_vals, myers_cutoff, color='#10B981', lw=2.8, label='Git 工程版 Myers (带启发式截断 Cutoff)')
    ax1.axvline(x=35, color='#F59E0B', lw=1.5, ls=':')
    ax1.text(37, 6300, "退化区 (D 很大，如整文件重写)\nGit 启动启发式截断防卡死", fontsize=8.5, color='#B45309',
             bbox=dict(boxstyle='round,pad=0.28', facecolor='#FEF3C7', edgecolor='#F59E0B'))
    ax1.set_title("退化边界 (Worst Case)：当两文件完全不同时 D -> N+M", fontsize=10.8, fontweight='bold', color='#0F172A')
    ax1.set_xlabel("文件差异度 D (占总行数比例 %)", fontsize=9.2)
    ax1.set_ylabel("状态转移操作次数 (相对值)", fontsize=9.2)
    ax1.legend(loc='upper left', fontsize=8.2)
    ax1.grid(True, alpha=0.25)

    # Right: Patience Diff / Histogram Diff anchor concept
    ax2.axis('off')
    ax2.set_title("Git 工业级进化：先找“定海神针”锚点，再分段跑 Myers", fontsize=10.8, fontweight='bold', color='#0F172A')
    explanation = (
        "【痛点：高频重复行 `{}` 迷魂阵】\n"
        "代码中大量空行与 `}` 容易把贪心带进假对角线。\n\n"
        "【Git 破局：Patience / Histogram 锚点】\n"
        "1. 抽取全篇仅出现 1 次的“唯一核心行”（如函数签名）\n"
        "2. O(N log N) 求唯一行 LIS 作为【定海神针】锚点\n"
        "3. 大网格切碎为小区间，再分段安全跑 Myers！"
    )
    ax2.text(0.02, 0.92, explanation, fontsize=9.1, color='#1E293B', va='top', linespacing=1.42,
             bbox=dict(boxstyle='round,pad=0.42', facecolor='#F8FAFC', edgecolor='#CBD5E1', lw=1.4))

    ax2.text(0.02, 0.25,
             "[实测开关] Git 命令行三种差分算法自由切换：\n"
             "• git diff --myers      (默认：日常改动速度最快)\n"
             "• git diff --patience   (先锚定唯一行，专治大段函数移动)\n"
             "• git diff --histogram  (Patience 加强版，按低频直方图锚定，Git 推荐)",
             fontsize=8.6, color='#0F172A', va='top', linespacing=1.32,
             bbox=dict(boxstyle='round,pad=0.35', facecolor='#EFF6FF', edgecolor='#60A5FA', lw=1.4))

    plt.tight_layout()
    path = os.path.join(ASSETS_DIR, "fig7_worst_case_and_git_variants.png")
    plt.savefig(path, dpi=240)
    plt.close(fig)
    return path


# ==============================================================================
# FIGURE 9: Rigorous Correctness Proof (Cut-and-Paste + Induction on D)
# ==============================================================================
def make_fig_proof():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12.2, 3.85), dpi=240,
                                   gridspec_kw={'width_ratios': [1.05, 1.2]})
    fig.patch.set_facecolor('#FFFFFF')

    # Left subplot: Racing analogy + Geometric Cut-and-Paste Argument on diagonal k
    ax1.set_xlim(-0.3, 6.3)
    ax1.set_ylim(5.6, -0.3)
    ax1.set_aspect('equal')
    ax1.axis('off')
    ax1.set_title("引理 1：对角线单调支配性（赛车追赶直觉 + 剪接法）",
                  fontsize=10.8, fontweight='bold', color='#0F172A', pad=8)

    # Shade target quadrant x >= 3, y >= 3
    quad = mpatches.Rectangle((3, 3), 3.0, 2.2, facecolor='#ECFDF5', edgecolor='none', alpha=0.85, zorder=0)
    ax1.add_patch(quad)
    ax1.text(5.15, 3.55, "终点所在象限\n{x >= x1, y >= y1}", fontsize=7.8, color='#065F46', ha='center', va='center')

    # Grid lines
    for x in range(7):
        ax1.plot([x, x], [0, 5.2], color='#CBD5E1', lw=0.9, zorder=1)
    for y in range(6):
        ax1.plot([0, 6], [y, y], color='#CBD5E1', lw=0.9, zorder=1)

    # Diagonal k=0 line
    ax1.plot([0, 5.2], [0, 5.2], color='#2563EB', lw=2.0, ls=':')
    ax1.text(1.95, 2.28, "车道 k=0", fontsize=8.0, color='#2563EB', fontweight='bold', rotation=-38)

    # L-shaped frontier rays starting at P1(3,3)
    ax1.plot([3, 6], [3, 3], color='#F59E0B', lw=2.4, ls='--', zorder=3)
    ax1.plot([3, 3], [3, 5.2], color='#F59E0B', lw=2.4, ls='--', zorder=3)
    ax1.text(4.85, 2.82, "门框边界 y=y1 (x>=x1)", fontsize=8.0, fontweight='bold', color='#B45309', ha='center')

    # Trailing P2=(1,1) vs Dominant P1=(3,3)
    ax1.scatter([1], [1], color='#EF4444', s=100, zorder=6)
    ax1.text(1.05, 0.55, "后车 P2(1,1)\n同在 k=0 但落后", fontsize=8.2, fontweight='bold', color='#991B1B', ha='center',
             bbox=dict(boxstyle='round,pad=0.18', facecolor='#FEE2E2', edgecolor='#EF4444'))

    ax1.scatter([3], [3], color='#10B981', s=130, zorder=7)
    ax1.text(2.05, 3.65, "领先车 P1(3,3)\n(x1 > x2, y1 > y2)", fontsize=8.2, fontweight='bold', color='#065F46', ha='center',
             bbox=dict(boxstyle='round,pad=0.18', facecolor='#D1FAE5', edgecolor='#10B981'))

    # Arbitrary path from P2(1,1) -> (2,1) -> (3,2) -> (4,2) -> Q(4,3) -> T(5,4)
    ax1.plot([1, 2, 3, 4, 4], [1, 1, 2, 2, 3], color='#EF4444', lw=2.5, ls='-', zorder=4)
    ax1.scatter([4], [3], color='#F59E0B', s=95, zorder=8)
    ax1.text(4.3, 2.35, "交点 Q(4,3)", fontsize=8.5, fontweight='bold', color='#B45309')

    # Short cut from P1(3,3) straight to Q(4,3) and then shared suffix Q -> T(5,4)
    ax1.annotate("", xy=(4, 3), xytext=(3, 3),
                 arrowprops=dict(arrowstyle="->", color="#10B981", lw=3.5), zorder=6)
    ax1.annotate("", xy=(5, 4), xytext=(4, 3),
                 arrowprops=dict(arrowstyle="->", color="#2563EB", lw=3.2), zorder=6)
    ax1.scatter([5], [4], color='#2563EB', s=110, marker='*', zorder=8)
    ax1.text(5.15, 4.1, "终点 T", fontsize=8.8, fontweight='bold', color='#1E40AF')

    ax1.text(1.85, 5.12,
             "P2 到 Q 至少费 (4-3)=1 步(实际费 3 步)；P1 直走至 Q 至多 1 步！\n"
             "把 Q->T 剪接给 P1 ==> P1 走得绝不比 P2 慢！",
             fontsize=8.0, color='#1E293B', ha='center', zorder=9,
             bbox=dict(boxstyle='round,pad=0.25', facecolor='#FEF3C7', edgecolor='#F59E0B', alpha=0.97))

    # Right subplot: Dual-Layer Explanation (Intuition + Rigorous Induction)
    ax2.axis('off')
    ax2.set_title("定理 2：从直观几何支配到严密数学归纳证明",
                  fontsize=10.8, fontweight='bold', color='#0F172A', pad=8)

    proof_boxes = [
        ("【几何直觉：同车道赛车追赶】",
         "同车道 k 上，后车 P2 想通向右下角终点，必跨过前车 P1 画出的 L 型门框；\n"
         "P1 站在门框顶点截胡接力跑，P2 永远无法反超！",
         "#FFFBEB", "#F59E0B"),
        ("【形式化证明 1：L-边界剪接引理 (Cut-and-Paste)】",
         "P2(x2,y2) -> (N,M) 必交 L={(x,y1):x>=x1} U {(x1,y):y>=y1} 于 Q。\n"
         "代价 cost(P2->Q) >= |k_Q - k| = max(x_q-x1, y_q-y1) >= cost(P1->Q)，剪接必不劣。",
         "#EFF6FF", "#3B82F6"),
        ("【形式化证明 2：对编辑步数 D 数学归纳法 P(D)】",
         "基：D=0 时 V_0[0] 沿 k=0 滑到底唯一最优。\n"
         "步：D 步最远点必由 D-1 步 k±1 最远点转移并免费滑到底得出，故全局最优！",
         "#ECFDF5", "#10B981")
    ]
    y_curr = 0.94
    for title, body, bg, border in proof_boxes:
        ax2.text(0.01, y_curr, f"{title}\n{body}", fontsize=8.85, color='#1E293B', va='top',
                 linespacing=1.40,
                 bbox=dict(boxstyle='round,pad=0.38', facecolor=bg, edgecolor=border, lw=1.35))
        y_curr -= 0.325

    plt.tight_layout()
    path = os.path.join(ASSETS_DIR, "fig_proof_cut_and_paste.png")
    plt.savefig(path, dpi=240)
    plt.close(fig)
    return path


# ==============================================================================
# FIGURE 10: Real Empirical Benchmark (Loaded from benchmark_results.json)
# ==============================================================================
def make_fig_benchmark():
    bench_path = os.path.join(ASSETS_DIR, "benchmark_results.json")
    with open(bench_path, "r", encoding="utf-8") as f:
        bench = json.load(f)

    n_lines = np.array(bench["sizes"])
    time_2d_dp = np.array(bench["time_ms"]["dp"])
    time_myers_basic = np.array(bench["time_ms"]["myers"])
    time_myers_bi = np.array(bench["time_ms"]["bi"])

    mem_2d_dp_mb = np.array(bench["mem_kb"]["dp"]) / 1024.0
    mem_myers_basic_mb = np.array(bench["mem_kb"]["myers"]) / 1024.0
    mem_myers_bi_mb = np.array(bench["mem_kb"]["bi"]) / 1024.0

    speedup_max = int(round(time_2d_dp[-1] / time_myers_basic[-1]))
    mem_save_max = int(round(bench["mem_kb"]["dp"][-1] / bench["mem_kb"]["bi"][-1]))

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12.2, 3.85), dpi=240)
    fig.patch.set_facecolor('#FFFFFF')

    ax1.plot(n_lines, time_2d_dp, 'o--', color='#EF4444', lw=2.4, markersize=6, label='课本二维 DP Θ(NM)')
    ax1.plot(n_lines, time_myers_bi, 's-', color='#10B981', lw=2.4, markersize=6, label='Myers 双向分治版 O((N+M)D)')
    ax1.plot(n_lines, time_myers_basic, '^-', color='#2563EB', lw=2.4, markersize=6, label='Myers 基础版 O((N+M)D)')
    ax1.set_yscale('log')
    ax1.set_xticks(n_lines)
    ax1.set_yticks([0.05, 0.5, 5.0, 50.0, 500.0])
    ax1.set_yticklabels(['0.05 ms', '0.5 ms', '5 ms', '50 ms', '500 ms'])
    ax1.minorticks_off()
    ax1.set_title("本机 Python 实测耗时 (对数坐标，固定日常改动 D=12 处)",
                  fontsize=10.6, fontweight='bold', color='#0F172A')
    ax1.set_xlabel("序列长度 N = M (行/字符)", fontsize=9.2)
    ax1.set_ylabel("墙钟运行耗时 (ms, 对数刻度)", fontsize=9.2)
    ax1.grid(True, alpha=0.25)
    ax1.legend(loc='upper left', fontsize=8.0)
    ax1.annotate(f"N={n_lines[-1]} 实测：\n二维 DP: {time_2d_dp[-1]:.1f} ms\nMyers: {time_myers_basic[-1]:.2f} ms\n提速 {speedup_max:,} 倍!",
                 xy=(n_lines[-1], time_myers_basic[-1]), xytext=(n_lines[1] + 200, 1.5),
                 arrowprops=dict(arrowstyle="->", color="#2563EB", lw=1.8),
                 fontsize=8.6, fontweight='bold', color='#1E40AF',
                 bbox=dict(boxstyle='round,pad=0.3', facecolor='#EFF6FF', edgecolor='#3B82F6'))

    ax2.plot(n_lines, mem_2d_dp_mb, 'o--', color='#EF4444', lw=2.4, markersize=6, label='课本二维 DP: O(NM) 内存爆炸')
    ax2.plot(n_lines, mem_myers_basic_mb, '^-', color='#F59E0B', lw=2.4, markersize=6, label='Myers 基础版 (存快照回溯): O(D^2)')
    ax2.plot(n_lines, mem_myers_bi_mb, 's-', color='#10B981', lw=2.6, markersize=6, label='Myers 双向分治版: 严格线性 O(N+M)')
    ax2.set_yscale('log')
    ax2.set_xticks(n_lines)
    ax2.set_yticks([0.01, 0.1, 1.0, 10.0, 100.0])
    ax2.set_yticklabels(['0.01 MB', '0.1 MB', '1 MB', '10 MB', '100 MB'])
    ax2.minorticks_off()
    ax2.set_title("本机 tracemalloc 峰值内存实测 (对数坐标，大重构 D≈0.35N)",
                  fontsize=10.6, fontweight='bold', color='#0F172A')
    ax2.set_xlabel("序列长度 N = M (行/字符)", fontsize=9.2)
    ax2.set_ylabel("峰值堆内存占用 (MB, 对数刻度)", fontsize=9.2)
    ax2.grid(True, alpha=0.25)
    ax2.legend(loc='upper left', fontsize=8.0)
    ax2.annotate(f"N={n_lines[-1]} 实测：\n二维 DP: {mem_2d_dp_mb[-1]:.1f} MB (45,233 KB)\n双向分治: {mem_myers_bi_mb[-1]:.2f} MB (206 KB)\n内存节省 {mem_save_max} 倍!",
                 xy=(n_lines[-1], mem_myers_bi_mb[-1]), xytext=(n_lines[1] + 50, 0.45),
                 arrowprops=dict(arrowstyle="->", color="#059669", lw=1.8),
                 fontsize=8.6, fontweight='bold', color='#065F46',
                 bbox=dict(boxstyle='round,pad=0.3', facecolor='#ECFDF5', edgecolor='#10B981'))

    plt.tight_layout()
    path = os.path.join(ASSETS_DIR, "fig_real_benchmark.png")
    plt.savefig(path, dpi=240)
    plt.close(fig)
    return path


# ==============================================================================
# PPTX BUILDER HELPERS
# ==============================================================================
C_NAVY = RGBColor(15, 23, 42)       # #0F172A
C_SLATE = RGBColor(51, 65, 85)      # #334155
C_MUTED = RGBColor(100, 116, 139)   # #64748B
C_BLUE = RGBColor(37, 99, 235)      # #2563EB
C_LIGHT_BG = RGBColor(248, 250, 252)# #F8FAFC
C_WHITE = RGBColor(255, 255, 255)
C_AMBER = RGBColor(217, 119, 6)     # #D97706
C_GREEN = RGBColor(5, 150, 105)     # #059669
C_RED = RGBColor(220, 38, 38)       # #DC2626
C_CARD_BORDER = RGBColor(203, 213, 225)


def set_slide_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_header(slide, badge_text, title_text, subtitle_text, page_num, total_pages=15):
    set_slide_bg(slide, C_LIGHT_BG)
    if badge_text.startswith("第二部分"):
        accent_fg = RGBColor(109, 40, 217)
        pill_bg = RGBColor(237, 233, 254)
    elif badge_text.startswith("附录"):
        accent_fg = C_SLATE
        pill_bg = RGBColor(226, 232, 240)
    else:
        accent_fg = C_BLUE
        pill_bg = RGBColor(219, 234, 254)

    # Top accent bar
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12))
    bar.fill.solid()
    bar.fill.fore_color.rgb = accent_fg
    bar.line.fill.background()

    # Badge pill (wide enough for Part 1 / Part 2 section labels)
    pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(0.28), Inches(3.45), Inches(0.32))
    pill.fill.solid()
    pill.fill.fore_color.rgb = pill_bg
    pill.line.fill.background()
    tf_p = pill.text_frame
    tf_p.word_wrap = False
    p = tf_p.paragraphs[0]
    p.text = badge_text
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = accent_fg
    p.alignment = PP_ALIGN.CENTER

    # Page counter
    tb_pg = slide.shapes.add_textbox(Inches(11.5), Inches(0.25), Inches(1.3), Inches(0.35))
    p_pg = tb_pg.text_frame.paragraphs[0]
    p_pg.text = f"{page_num:02d} / {total_pages:02d}"
    p_pg.font.size = Pt(11)
    p_pg.font.bold = True
    p_pg.font.color.rgb = C_MUTED
    p_pg.alignment = PP_ALIGN.RIGHT

    # Title
    tb_t = slide.shapes.add_textbox(Inches(0.55), Inches(0.62), Inches(12.2), Inches(0.55))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True
    p_t = tf_t.paragraphs[0]
    p_t.text = title_text
    p_t.font.size = Pt(24)
    p_t.font.bold = True
    p_t.font.color.rgb = C_NAVY

    # Subtitle
    if subtitle_text:
        tb_s = slide.shapes.add_textbox(Inches(0.55), Inches(1.12), Inches(12.2), Inches(0.35))
        tf_s = tb_s.text_frame
        tf_s.word_wrap = True
        p_s = tf_s.paragraphs[0]
        p_s.text = subtitle_text
        p_s.font.size = Pt(13)
        p_s.font.color.rgb = C_SLATE


def add_card(slide, left, top, width, height, title, bullets,
             bg_color=C_WHITE, border_color=C_CARD_BORDER, title_color=C_NAVY, font_size=13):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.adjustments[0] = 0.03
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(1.2)

    tb = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.12), Inches(width - 0.36), Inches(height - 0.24))
    tf = tb.text_frame
    tf.word_wrap = True

    if title:
        p0 = tf.paragraphs[0]
        p0.text = title
        p0.font.size = Pt(font_size + 2)
        p0.font.bold = True
        p0.font.color.rgb = title_color
        p0.space_after = Pt(6)

    for idx, b in enumerate(bullets):
        p = tf.paragraphs[0] if (idx == 0 and not title) else tf.add_paragraph()
        p.text = b
        p.font.size = Pt(font_size)
        p.font.color.rgb = C_SLATE
        p.space_after = Pt(5)
    return card


def add_notes(slide, notes_text):
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text


def build_presentation():
    print("Generating high-DPI figures & animations...")
    f1 = make_fig1_pain_points()
    f2 = make_fig2_edit_graph()
    f3 = make_fig3_k_diagonals()
    f4 = make_fig4_state_transition()
    f_proof = make_fig_proof()
    f5a, f5b, f5_gif = make_fig5_and_gif()
    f_race_static, f_race_gif = make_race_gif()
    f6 = make_fig6_bidirectional()
    f_bench = make_fig_benchmark()
    f7 = make_fig7_worst_case()

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # ==========================================================================
    # SLIDE 1: Cover Slide
    # ==========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s1, C_NAVY)

    # Decorative accent stripe
    stripe = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.75), Inches(1.1), Inches(0.14), Inches(2.6))
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = RGBColor(59, 130, 246)
    stripe.line.fill.background()

    tb1 = s1.shapes.add_textbox(Inches(1.1), Inches(1.05), Inches(11.5), Inches(2.8))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p_tag = tf1.paragraphs[0]
    p_tag.text = "《算法设计与分析》课程专题汇报  |  Part 1 直观算法图解 · Part 2 理论证明与工程进阶"
    p_tag.font.size = Pt(14)
    p_tag.font.bold = True
    p_tag.font.color.rgb = RGBColor(96, 165, 250)
    p_tag.space_after = Pt(14)

    p_main = tf1.add_paragraph()
    p_main.text = "Git 底层是如何算代码差异的？\n—— Myers 差分算法深度解析"
    p_main.font.size = Pt(40)
    p_main.font.bold = True
    p_main.font.color.rgb = C_WHITE
    p_main.space_after = Pt(14)

    p_sub = tf1.add_paragraph()
    p_sub.text = "从课本 O(N×M) 最长公共子序列 (LCS) 到 O((N+M)D) 编辑图贪心搜索与线性空间分治"
    p_sub.font.size = Pt(17)
    p_sub.font.color.rgb = RGBColor(203, 213, 225)

    # Concise Core Insight Box on Cover (3 punchy bullets)
    hook_box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.75), Inches(4.35), Inches(11.8), Inches(2.35))
    hook_box.adjustments[0] = 0.04
    hook_box.fill.solid()
    hook_box.fill.fore_color.rgb = RGBColor(30, 41, 59)
    hook_box.line.color.rgb = RGBColor(245, 158, 11)
    hook_box.line.width = Pt(2.0)

    tb_h = s1.shapes.add_textbox(Inches(1.0), Inches(4.5), Inches(11.3), Inches(2.05))
    tf_h = tb_h.text_frame
    tf_h.word_wrap = True
    ph0 = tf_h.paragraphs[0]
    ph0.text = "🎯 核心问题与破局思路 (Core Problem & Key Insight)"
    ph0.font.size = Pt(15)
    ph0.font.bold = True
    ph0.font.color.rgb = RGBColor(251, 191, 36)
    ph0.space_after = Pt(8)

    cover_bullets = [
        "• 痛点：课本二维 DP 求 LCS 恒定计算 Θ(N·M) 格，万行文件哪怕只改 1 行也卡顿，且极易错配孤立 `{}`。",
        "• 破局：构造带“免费对角线滑梯”的编辑图，按最小编辑步数 D 与车道 k=x-y 贪心推进。",
        "• 效果：时间降至 O((N+M)D)（改动越少越快），结合双向分治将空间压至严格线性 O(N+M)。"
    ]
    for cb in cover_bullets:
        ph_b = tf_h.add_paragraph()
        ph_b.text = cb
        ph_b.font.size = Pt(14)
        ph_b.font.color.rgb = RGBColor(241, 245, 249)
        ph_b.space_after = Pt(6)

    add_notes(s1, (
        "【第 1 页演讲逐字稿（约 40 秒）】\n"
        "老师、同学们好！今天我给大家分享《Git 底层是如何算代码差异的？——Myers 差分算法》。\n"
        "本次汇报分为两部分：第一部分用全直观图解带大家走通算法核心流程；第二部分再深入它的数学证明、线性空间分治和工业界实验对比。\n"
        "先用一句话破题：课本上学过的二维动态规划求 LCS，哪怕比对两个仅改了 1 行的万行代码文件，也要死板算完 1 亿个格子；而 Git 底层的 Myers 算法把二维表格改造成带免费斜向滑梯的“编辑图”，改按编辑步数 D 向外搜索，让时间降维到 O((N+M)D)——改得越少跑得越快！"
    ))

    # ==========================================================================
    # PART 1 (SLIDE 2): Why Textbook LCS DP Fails in Real-World `git diff`
    # ==========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "第一部分 · 直观图解 | 01",
               "为什么课本上的二维 LCS 动态规划在 Git 里行不通？",
               "理论背景：从课本最长公共子序列 (LCS) 与编辑距离看真实工程瓶颈", 2)

    s2.shapes.add_picture(f1, Inches(0.55), Inches(1.55), width=Inches(12.2))

    add_card(s2, 0.55, 5.65, 12.2, 1.35,
             "💡 核心洞察：绝大多数日常代码提交都属于“局部小改动”（D ≪ N, M）",
             [
                 "• 课本 DP 像“盲目压路机”：无论两文件有多相似，恒定计算全部 N×M 格，复杂度 Θ(N·M)。",
                 "• Myers 是输出敏感算法：专顺相同代码对角线滑行，仅探索差异步数 D 周边窄带，复杂度 O((N+M)D)。"
             ],
             bg_color=RGBColor(239, 246, 255), border_color=C_BLUE, title_color=C_BLUE, font_size=12.5)

    add_notes(s2, (
        "【第 2 页演讲逐字稿（约 1 分钟）】\n"
        "我们先回顾课本《动态规划》章节必讲的 LCS。给定两个序列 A 和 B，课本用二维数组 dp[i][j] 填表。\n"
        "但直接把这个算法塞进 Git 会遇到两个致命问题：\n"
        "第一是性能灾难（看左图）：想象你在改毕业设计代码，两千多行里其实只改了十几处。课本的二维 DP 就像一台“盲目的压路机”，老老实实填满 N×M 全部格子；而 Myers 算法像一束激光手电筒，只沿着主对角线附近窄窄的一条带宽搜索，本机实测 2400 行文件提速超 2131 倍！\n"
        "第二是可读性灾难（看右图）：普通 LCS 存在很多等价解，回溯时常常把旧函数结尾的右花括号 `}` 和新函数结尾的 `}` 强行配对，导致 diff 支离破碎；而 Myers 的贪心滑行策略偏好连续匹配，天然产出整洁的变更块。"
    ))

    # ==========================================================================
    # PART 1 (SLIDE 3): Modeling as an Edit Graph
    # ==========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "第一部分 · 直观图解 | 02",
               "第一步转化：从“二维 DP 填表”到“带免费滑梯的编辑图 (Edit Graph)”",
               "以 Eugene Myers 1986 原版论文经典算例 A = \"ABCABBA\" (N=7), B = \"CBABAC\" (M=6) 为例", 3)

    s3.shapes.add_picture(f2, Inches(0.55), Inches(1.55), width=Inches(12.2))

    add_notes(s3, (
        "【第 3 页演讲逐字稿（约 1 分 15 秒）】\n"
        "怎么让算法学会跳过相同代码呢？Myers 的第一步转化，是把填表格变成走迷宫——构造【编辑图 Edit Graph】。\n"
        "大家看左图：横轴是旧串 A = ABCABBA，纵轴是新串 B = CBABAC。从左上角 (0,0) 走到右下角 (7,6)，一共有三种走法：\n"
        "1. 红色向右走一格：表示把 A 的当前字符删掉（-），编辑代价 +1；\n"
        "2. 绿色向下走一格：表示从 B 插入一个字符（+），编辑代价 +1；\n"
        "3. 黄色斜对角线（Snake）：只要 A[x] == B[y]，就可以沿斜线同时消掉两个相同字符，编辑代价为 0，就像免费滑梯！\n"
        "于是，原问题就转化成：从 (0,0) 到 (7,6)，怎样花最少的非对角线步数 D（即蹭到最多免费滑梯）抵达终点？"
    ))

    # ==========================================================================
    # PART 1 (SLIDE 4): Coordinate Transformation (x, y) -> (D, k = x - y)
    # ==========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "第一部分 · 直观图解 | 03",
               "第二步转化：歪头 45° 定义车道 k = x - y，二维坐标压成一维 V[k]",
               "核心几何直观：每条斜线 k 相当于一条斜向下的“高速车道”；同一车道只需记录冲得最远的横坐标 V[k]", 4)

    s4.shapes.add_picture(f3, Inches(0.55), Inches(1.55), width=Inches(12.2))

    add_notes(s4, (
        "【第 4 页演讲逐字稿（约 1 分 20 秒）】\n"
        "迷宫画好了，如果直接跑普通 BFS 最短路径，状态依然是二维坐标 (x, y)，内存还是很大。\n"
        "Myers 做了第二个天才变形：让我们歪头 45° 看这张网格图，给每条斜对角线编一个车道号 k = x - y！\n"
        "大家看左图的彩色虚线：正中间主对角线是 0 号车道 k=0，往右上方偏是 1、2 号车道，往左下方偏是 -1、-2 号车道。\n"
        "这个坐标系有三个神仙性质（看右边）：\n"
        "第一，顺着免费滑梯斜走时，x 和 y 同时 +1，所以 x - y 保持不变——也就是说“坐滑梯永远不换车道”！只有花门票向右走（删除）才会往右换道到 k+1，向下走（插入）才会往左换道到 k-1。\n"
        "第二，奇偶性引理：从 0 号车道出发，每花 1 步门票就必须左移或右移一条车道，所以花了 D 张门票后，你所在车道 k 的奇偶性必然和 D 相同！第 D 轮有一半的车道根本不可能有人，直接跳过！\n"
        "第三，二维压一维：在同一条车道 k 上，y = x - k 是被锁死的，谁的横坐标 x 越大，纵坐标 y 也一定越大，就离终点越近！所以我们不必记所有点，每条车道只保留“车头”的最大横坐标 x，记为一维数组 V[k]！"
    ))

    # ==========================================================================
    # PART 1 (SLIDE 5): State Transition Grid + 7-Line Core Logic
    # ==========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "第一部分 · 直观图解 | 04",
               "核心递推：先择优换道迈 1 步 (DP)，再顺对角线滑到底 (Greedy)",
               "每一轮只做两件事：从相邻车道择优换道迈出 1 步（代价 +1），再顺匹配对角线免费滑到底（代价 0）", 5)

    s5.shapes.add_picture(f4, Inches(0.55), Inches(1.55), width=Inches(6.2))

    add_card(s5, 0.55, 6.12, 6.2, 0.91,
             "",
             [
                 "🚗 核心转移：先从左右邻居车道挑 x 更大的一侧跨 1 步（绿下 / 红右），再顺蓝斜线滑到底，更新 V[k]。"
             ],
             bg_color=RGBColor(254, 243, 199), border_color=C_AMBER, title_color=C_AMBER, font_size=11.5)

    # Right top card: Concise Mathematical recurrence
    add_card(s5, 6.95, 1.55, 5.8, 2.35,
             "📐 状态转移规则（第 D 轮，车道 k）",
             [
                 "1. 迈出第 D 步（从 D-1 轮相邻车道二选一）：",
                 "   • 从 k-1 右移（删除）：x = V[k-1] + 1",
                 "   • 从 k+1 下移（插入）：x = V[k+1]",
                 "   • 择优落点：x = max(V[k-1] + 1, V[k+1])  (边界 k=±D 单向)",
                 "2. 顺对角线免费滑到底（Greedy Snake）：",
                 "   • while A[x] == B[x - k]: x += 1，最后写回 V[k] = x"
             ],
             bg_color=C_WHITE, border_color=C_BLUE, title_color=C_BLUE, font_size=12.0)

    # Right bottom card: Strictly <= 10 lines core snippet (7 lines!)
    code_card = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.95), Inches(4.08), Inches(5.8), Inches(2.95))
    code_card.adjustments[0] = 0.04
    code_card.fill.solid()
    code_card.fill.fore_color.rgb = C_NAVY
    code_card.line.color.rgb = RGBColor(59, 130, 246)
    code_card.line.width = Pt(1.5)

    tb_c = s5.shapes.add_textbox(Inches(7.10), Inches(4.16), Inches(5.5), Inches(2.80))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True
    pc0 = tf_c.paragraphs[0]
    pc0.text = "⚡ Myers 核心状态转移伪代码（原地滚动更新一维数组 V[k]）"
    pc0.font.size = Pt(12)
    pc0.font.bold = True
    pc0.font.color.rgb = RGBColor(251, 191, 36)
    pc0.space_after = Pt(5)

    colored_lines = [
        ("for d in 0 .. N + M:             # 1. 最少修改步数 D", RGBColor(52, 211, 153)),
        ("  for k in -d, -d+2, ..., +d:    # 2. 同奇偶车道 k", RGBColor(147, 197, 253)),
        ("    if k==-d or (k!=d and V[k-1]<V[k+1]):", RGBColor(226, 232, 240)),
        ("      x = V[k+1]                 # 3. k+1 向下插入", RGBColor(96, 165, 250)),
        ("    else: x = V[k-1] + 1         # 4. k-1 向右删除", RGBColor(248, 113, 113)),
        ("    while x<N and x-k<M and A[x]==B[x-k]: x+=1 # 5.免费滑", RGBColor(251, 191, 36)),
        ("    V[k]=x; if x>=N and x-k>=M: return d  # 6-7. 到终点!", RGBColor(52, 211, 153)),
    ]
    for line_txt, line_rgb in colored_lines:
        p_code = tf_c.add_paragraph()
        p_code.text = line_txt
        p_code.font.name = "Menlo"
        p_code.font.size = Pt(9.8)
        p_code.font.bold = True
        p_code.font.color.rgb = line_rgb
        p_code.space_after = Pt(2.5)

    add_notes(s5, (
        "【第 5 页演讲逐字稿（约 1 分钟）】\n"
        "Myers 每一步只有两个动作——“先换道，再坐滑梯”：\n"
        "大家看左图，第 D 轮计算目标车道 k：只能从左方 k-1 向右跨一格（x = V[k-1]+1）或上方 k+1 向下跨一格（x = V[k+1]），择优挑选 x 更大的落脚点，随后只要字符 A[x] == B[y] 就顺斜线免费滑到底，把终点横坐标写回 V[k]。\n"
        "右下角仅 7 行伪代码就是全部状态转移：第 D 轮只写奇偶性与 D 相同的 k，只读奇偶性相反的 k±1，读写零冲突，直接在单个一维数组 V[k] 原地覆盖即可。"
    ))

    # ==========================================================================
    # PART 1 (SLIDE 6): Step-by-Step Visual Trace Part A (D = 0 -> 1 -> 2)
    # ==========================================================================
    s6a = prs.slides.add_slide(blank_layout)
    add_header(s6a, "第一部分 · 直观图解 | 05",
               "实例推演（上）：D = 0 → 1 → 2 起步探索与首次触发免费滑梯",
               "以 A = \"ABCABBA\"、B = \"CBABAC\" 为例：放大看清每轮如何先跨 1 步，再沿斜线滑行并写回 V[k]", 6)

    s6a.shapes.add_picture(f5a, Inches(0.55), Inches(1.55), width=Inches(12.2))

    add_card(s6a, 0.55, 5.65, 12.2, 1.40,
             "🔍 前 3 轮推演要点（注意奇偶交替）",
             [
                 "• D = 0（偶车道）：起点 (0,0) 对应 A[0]='A' ≠ B[0]='C'，无滑梯可走，记录 V[0] = 0。",
                 "• D = 1（奇车道）：向右删 A 到 (1,0) 得 V[+1] = 1；向下插 C 到 (0,1) 得 V[-1] = 0。",
                 "• D = 2（偶车道）：从 (1,0) 右移至 (2,0) 触发 C==C 免费滑到 (3,1)（V[+2]=3）；下路从 (0,1) 下移至 (0,2) 连滑两格到达 (2,4)（V[-2]=2）！"
             ],
             bg_color=RGBColor(239, 246, 255), border_color=C_BLUE, title_color=C_BLUE, font_size=11.5)

    add_notes(s6a, (
        "【第 6 页演讲逐字稿（约 1 分钟）】\n"
        "前面规则都讲清了，我们现在放大网格，带大家把 ABCABBA 和 CBABAC 这组经典例子从 D=0 跑到 D=2：\n"
        "- 看左图 D=0：站在左上角 (0,0)，因为第一个字母 A 和 C 不相等，滑梯没触发，只能留在原地 V[0]=0。\n"
        "- 看中图 D=1：必须花 1 步代价换道。向右删掉 A 停在 (1,0)（k=+1 车道），向下插一个 C 停在 (0,1)（k=-1 车道），正好都没有滑梯。\n"
        "- 看右图 D=2：精彩开始了！上路从 (1,0) 再向右走一格到 (2,0)，立刻撞上 C==C 的对角线滑梯，免费滑到 (3,1)，所以 V[+2] 直接跳到 3；下路更厉害，向下到 (0,2) 后撞上 A、B 两格连滑，直抵 (2,4)！"
    ))

    # ==========================================================================
    # PART 1 (SLIDE 7): Step-by-Step Visual Trace Part B (D = 3 -> 4 -> 5) + GIF
    # ==========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "第一部分 · 直观图解 | 06",
               "实例推演（下）：D = 3 → 4 → 5 长滑梯冲刺终点与动态回放",
               "左上：D = 3..5 放大轨迹  |  右上：同步波前与 V[k] 状态条动画  |  左下：一维数组 V[k] 完整演进记录", 7)

    s6.shapes.add_picture(f5b, Inches(0.45), Inches(1.55), width=Inches(7.75))
    s6.shapes.add_picture(f5_gif, Inches(8.35), Inches(1.55), width=Inches(4.5))

    v_snapshot_lines = []
    for h in HISTORY:
        d_val = h['d']
        parity = "偶" if d_val % 2 == 0 else "奇"
        items = [f"V[{e['k']:+d}]={e['ex']}" for e in h['edges']]
        suffix = "  ★ 抵达 (7,6) 停!" if d_val == OPTIMAL_D else ""
        v_snapshot_lines.append(f"• D={d_val}({parity}): " + ", ".join(items) + suffix)

    add_card(s6, 0.45, 5.32, 7.75, 1.75,
             "📋 一维数组 V[k] 演进记录（仅覆盖同奇偶车道，读写永不冲突）",
             v_snapshot_lines,
             bg_color=RGBColor(239, 246, 255), border_color=C_BLUE, title_color=C_BLUE, font_size=10.2)

    add_card(s6, 8.35, 5.32, 4.5, 1.75,
             "🎯 推演关键节点解析（D = 3 → 5）",
             [
                 "• D=3：从 (3,1) 向下至 (3,2)，连滑两格 A、B 直冲 (5,4)（V[+1]=5）！",
                 "• D=4：k=+2 与 k=+4 并驾冲至最右列 x=7（抵达 (7,5)）。",
                 "• D=5：从 (7,5) 向下踩中 (7,6)，立即终止，最小距离 D=5！"
             ],
             bg_color=RGBColor(254, 243, 199), border_color=C_AMBER, title_color=C_AMBER, font_size=10.8)

    add_notes(s6, (
        "【第 7 页演讲逐字稿（约 1 分钟）】\n"
        "接着看后半程 D=3 到 D=5，右下角 GIF 也同步展示了底部一维数组 V[k] 的跳变：\n"
        "- D=3 时，从刚才的 (3,1) 向下踩一格到 (3,2)，直接触发连续两格长滑梯（A==A, B==B），瞬间推进到 (5,4)；\n"
        "- D=4 时，多条车头已经摸到了最右边界 x=7 的 (7,5)；\n"
        "- D=5 时，从 (7,5) 往下迈最后一步精准命中终点 (7,6)，搜索立刻终止！左下角列出了 6 轮下来一维数组 V[k] 的全部更新数值，整个 7×6 的网格有一大片空白区域压根看都没看。"
    ))

    # ==========================================================================
    # PART 1 (SLIDE 8): Search Space Comparison Animation (Textbook DP vs Myers)
    # ==========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "第一部分 · 直观图解 | 07",
               "搜索空间对比：课本二维 DP“压路机全表扫描” vs. Myers“激光束聚焦”",
               "直观理解为什么时间复杂度能从 Θ(N×M) 降为 O((N+M)D)", 8)

    s7.shapes.add_picture(f_race_gif, Inches(0.55), Inches(1.55), width=Inches(7.6))

    add_card(s7, 0.55, 5.35, 7.60, 1.70,
             "🔦 几何对比：为什么 Myers 比传统二维 DP 快得多？",
             [
                 "• 课本 DP（左）：像压路机一样必须逐行填满 N×M 全表，即使几乎没改动也不例外。",
                 "• Myers（右）：像激光手电筒仅沿主对角线 ±D 窄带推进，已早早命中终点！"
             ],
             bg_color=RGBColor(239, 246, 255), border_color=C_BLUE, title_color=C_BLUE, font_size=11.5)

    add_card(s7, 8.35, 1.55, 4.45, 5.50,
             "📊 时间复杂度 O((N+M)D) 的几何摊还证明",
             [
                 "核心问题：双层 for 嵌套 while 循环，为何总复杂度不是 O(N³)？",
                 "1. 外层循环：精确执行 D + 1 轮。",
                 "2. 换道常数步：第 d 轮查 d+1 条车道，累计 O(D²) 次。",
                 "3. while 滑梯累计总步数：",
                 "   同一车道 k 上，横坐标 x 只增不减！",
                 "   全场至多活跃 2D+1 条车道，",
                 "   每条最长不超过 N+M 格，",
                 "   因此全部 while 累计 ≤ (2D+1)(N+M) 步。",
                 "4. 总时间复杂度：",
                 "   T = O(D² + (N+M)D) = O((N+M)D) 🏆",
                 "   👉 日常小改动 (D ≪ N) 时近似线性 O(N)！"
             ],
             bg_color=C_WHITE, border_color=C_GREEN, title_color=C_GREEN, font_size=11.3)

    add_notes(s7, (
        "【第 8 页演讲逐字稿（约 1 分钟）】\n"
        "第一部分最后看这张同屏竞速动画，能最直观感受到差距：红色二维 DP 像一台压路机死板填表；蓝色 Myers 像手电筒聚焦在主对角线附近，DP 才扫一半 Myers 就收工了。\n"
        "顺便回答一个大家看 7 行代码时常见的疑问：双层 for 里面还有个 while 循环，总时间怎么不是立方级？\n"
        "用右边的几何摊还一看就通：在任何一条车道 k 上，车头 x 只会往右下推进、绝不回头！全场总共只碰过 2D+1 条车道，每条车道封顶长度 N+M，所以哪怕套在双层循环里，所有 while 加起来也最多走 O((N+M)D) 步。至此，基础版 Myers 的原理与推演全部讲完，接下来进入第二部分：它的严格数学证明与线性空间黑科技。"
    ))

    # ==========================================================================
    # PART 2 (SLIDE 9): Rigorous Mathematical Correctness Proof (Cut-and-Paste + Induction)
    # ==========================================================================
    s_proof = prs.slides.add_slide(blank_layout)
    add_header(s_proof, "第二部分 · 理论与进阶 | 08",
               "正确性数学证明：为什么“同车道只留最远点 + 贪心滑到底”必得全局最优？",
               "直观同车道追赶模型 + 严格形式化证明：贪心支配引理 (Cut-and-Paste) 与数学归纳法", 9)

    s_proof.shapes.add_picture(f_proof, Inches(0.55), Inches(1.55), width=Inches(12.2))

    add_card(s_proof, 0.55, 5.60, 12.2, 1.40,
             "🏆 核心定理总结：局部贪心滑行何以保证全局最优？",
             [
                 "• 几何直观（同车道追赶）：同一车道 k 上，领先车 P₁(x₁,y₁) 已封锁落后车 P₂(x₂,y₂) 通向终点的必经 L 型门框，故每条车道只留第一名。",
                 "• 形式化定理（剪接引理 + 归纳法）：从 P₂ 去往 (N,M) 必跨过 L={x=x₁ 或 y=y₁} 于 Q，由 d(P₂→Q) ≥ d(P₁→Q) 剪接必不劣；对 D 归纳即证 V_D[k] 最优！"
             ],
             bg_color=RGBColor(239, 246, 255), border_color=C_BLUE, title_color=C_BLUE, font_size=11.5)

    add_notes(s_proof, (
        "【第 9 页演讲逐字稿（约 1 分 15 秒）】\n"
        "进入第二部分，首先从理论上严密闭环一个核心问题：为什么每次遇到相等字符无脑滑到底、同车道只留最远点，一定不会错过全局最优解？\n"
        "看左图，用课本贪心章节经典的【剪接反证法（Cut-and-Paste）】：假设同在车道 k 上有领先点 P1(3,3) 和落后点 P2(1,1)。\n"
        "从 P2 出发通往终点 T 的任何路径，必然穿过 P1 引出的黄色 L 型门框边界于某点 Q(4,3)。从 P2 到 Q 至少需要耗费 |k_Q - k| 步以上代价，绝不比 P1 直接跨向 Q 更少！把 Q 之后的路径剪下来接给 P1，新路径总步数一定不劣于原路径。再配合对步数 D 的数学归纳法 P(D)，就严格证明了算法的全局最优性。"
    ))

    # ==========================================================================
    # PART 2 (SLIDE 10): Linear Space O(N+M) Divide-and-Conquer + Bidirectional Search
    # ==========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "第二部分 · 理论与进阶 | 09",
               "空间极限优化：只留一维 V[k] 怎么回溯路径？—— 双向搜索找“中间蛇”",
               "Myers 1986 线性空间变体：双向波前搜索 + Hirschberg 分治，将空间从 O(D²) 压至严格线性 O(N+M)", 10)

    s8.shapes.add_picture(f6, Inches(0.55), Inches(1.55), width=Inches(12.2))

    add_card(s8, 0.55, 5.60, 5.95, 1.40,
             "🚇 分治递推与时间复杂度的几何级数证明",
             [
                 "• 正向/反向各推进 D/2 步在中间蛇碰撞切分，历史状态丢弃复用。",
                 "• T(N,D) ≤ c(N+M)(D/2) + T₁ + T₂ ⇒ 1/2 + 1/4 + 1/8 + … = 1 ⇒ 仍为 O((N+M)D)！"
             ],
             bg_color=RGBColor(254, 242, 242), border_color=C_RED, title_color=C_RED, font_size=11.0)

    add_card(s8, 6.75, 5.60, 6.0, 1.40,
             "🏆 双向分治版（Git 默认策略）：空间骤降至 O(N+M)",
             [
                 "• 全程仅维护正向 V_fwd[k] 与反向 V_bwd[k] 两个一维数组。",
                 "• 递归树栈深仅 O(log D)，不牺牲渐近时间实现 O(D²) ➔ O(N+M) 压缩！"
             ],
             bg_color=RGBColor(236, 253, 245), border_color=C_GREEN, title_color=C_GREEN, font_size=11.0)

    add_notes(s8, (
        "【第 10 页演讲逐字稿（约 1 分 15 秒）】\n"
        "单向跑 Myers 虽然快，但如果只覆写一个一维数组 V[k] 就会丢失回溯路径；如果每轮拍照存档，大重构时空间又会膨胀回 O(D²)。\n"
        "Myers 在论文第 4 节借用 Hirschberg 分治给出绝妙解法——两头挖隧道：\n"
        "一边从 (0,0) 正挖 D/2 步，另一边从 (N,M) 倒挖 D/2 步，两队在半空相撞锁定【中间蛇 (Middle Snake)】，随后擦掉旧数组，递归求解左上和右下两个半大子网格！看左下角等比级数 1/2+1/4+1/8+...=1，总时间依然是 O((N+M)D)，而空间严格压到了线性 O(N+M)！"
    ))

    # ==========================================================================
    # PART 2 (SLIDE 11): Horizontal Comparison Table (MANDATORY)
    # ==========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "第二部分 · 理论与进阶 | 10",
               "算法横向对比：课本基础算法 vs. Myers 差分家族",
               "时间复杂度、空间复杂度、退化边界 (Worst Case) 与工程策略一览表", 11)

    rows, cols = 5, 6
    table_shape = s9.shapes.add_table(rows, cols, Inches(0.55), Inches(1.56), Inches(12.2), Inches(4.32))
    table = table_shape.table

    col_widths = [Inches(2.0), Inches(1.85), Inches(2.05), Inches(1.45), Inches(2.45), Inches(2.4)]
    for i, w in enumerate(col_widths):
        table.columns[i].width = w
    table.rows[0].height = Inches(0.56)
    for r_i in range(1, rows):
        table.rows[r_i].height = Inches(0.94)

    headers = [
        "算法名称",
        "核心算法思想",
        "时间复杂度\n(最优 / 平均 / 最坏)",
        "空间复杂度",
        "退化边界情况\n(Worst Case 触发条件)",
        "适用场景与缺陷"
    ]

    data = [
        [
            "课本标准二维 DP\n(Wagner-Fischer)",
            "二维网格动态规划\n(LCS / 编辑距离)",
            "最优: Θ(N·M)\n平均: Θ(N·M)\n最坏: Θ(N·M)",
            "O(N·M) 回溯路径\n(仅求长度可滚动至 O(min(N,M)))",
            "无退化，因为永远处于最坏情况！\n哪怕两文件完全相同 (D=0) 也要算满 N×M 次",
            "仅适合短字符串教学演示；\n千行以上代码文件耗时与内存双爆"
        ],
        [
            "Hunt-McIlroy 算法\n(早期 Unix diff)",
            "稀疏匹配点对 +\n最长递增子序列 (LIS)",
            "最优: O(N log N)\n平均: O((N+R) log N)\n最坏: O(N·M log N)",
            "O(N + R)\n(R 为匹配点对总数)",
            "当文件含大量高频重复行（如空行、`}`）时，匹配点对数 R 退化至 N·M 级",
            "文本行少且重复率低时快；\n遇到大量括号/空行时急剧退化"
        ],
        [
            "Myers 算法\n(基础贪心 DP 版)",
            "编辑图最短路 +\n对角线贪心 + DP",
            "最优: O(N + M) (D=0)\n平均: O((N+M)D)\n最坏: O((N+M)²)",
            "O(D²) 存快照回溯\n(仅求 D 需 O(D))",
            "两文件几乎完全不同 (D ≈ N+M) 时，时间退化为 O((N+M)²)，空间 O((N+M)²)",
            "小改动极快 (D ≪ N)；\n但大段重写时回溯内存开销大"
        ],
        [
            "★ Myers 线性空间版\n(Git 底层默认引擎)",
            "双向 Myers 搜索 +\nHirschberg 分治法",
            "最优: O(N + M)\n平均: O((N+M)D)\n最坏: O((N+M)²)",
            "严格线性\nO(N + M) 🏆",
            "1. 两文件完全不同 (D → N+M)\n2. 大量重复 `{}` 形成“假对角线”误导贪心",
            "Git 默认核心算法；\n工程上配合截断阈值与 Patience 锚点规避退化"
        ]
    ]

    for c_idx, h_text in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.size = Pt(11.0)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.CENTER

    for r_idx, row_data in enumerate(data, start=1):
        is_highlight = (r_idx == 4)
        row_bg = RGBColor(239, 246, 255) if is_highlight else (C_WHITE if r_idx % 2 == 1 else RGBColor(241, 245, 249))
        for c_idx, val in enumerate(row_data):
            cell = table.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = row_bg
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.size = Pt(10.0)
            p.font.bold = (c_idx == 0 or is_highlight)
            p.font.color.rgb = C_BLUE if (is_highlight and c_idx in (0, 3)) else C_NAVY

    add_card(s9, 0.55, 6.18, 12.2, 0.88,
             "",
             [
                 "📌 核心结论：Myers 将输出规模参数 D 引入复杂度 O((N+M)D)，在日常代码差异极小 (D ≪ N) 的场景下对 Θ(N·M) 形成数量级优势！"
             ],
             bg_color=RGBColor(254, 243, 199), border_color=C_AMBER, title_color=C_NAVY, font_size=12.0)

    add_notes(s9, (
        "【第 11 页演讲逐字稿（约 1 分钟）】\n"
        "这张表把课本基础算法和工程演进版本放在一起横向对比：课本二维 DP 没有最优情况，永远 Θ(N·M)；"
        "Hunt-McIlroy 在大量空行/括号下匹配点 R 易退化；Myers 线性空间版则同时拿下最佳情况 O(N+M)、常规小改动 O((N+M)D) 和严格线性空间 O(N+M)。"
        "当然，如果两份代码彻底面目全非（D ≈ N+M），它的最坏情况也会回到 O((N+M)²)。下面看实测跑分。"
    ))

    # ==========================================================================
    # PART 2 (SLIDE 12): Real Empirical Benchmark (Time ms & Peak Memory KB/MB)
    # ==========================================================================
    s_bench = prs.slides.add_slide(blank_layout)
    add_header(s_bench, "第二部分 · 理论与进阶 | 11",
               "实验性能与内存实测：本机 2,131 倍提速与 219 倍内存压缩",
               "Mac 本机 Python tracemalloc 实测（N = 300 ~ 2,400 行） | 左：日常小改 (D=12) | 右：大重构峰值内存 (D≈0.35N)", 12)

    s_bench.shapes.add_picture(f_bench, Inches(0.55), Inches(1.55), width=Inches(12.2))

    add_card(s_bench, 0.55, 5.60, 12.2, 1.40,
             "📈 本机实测结论（N = 2,400 行对数坐标对比）",
             [
                 "• 运行耗时（左图，D=12）：二维 DP 耗时 275.3 ms，Myers 基础版仅 0.13 ms（实测提速 2,131×），双向版 0.18 ms。",
                 "• 峰值内存（右图，D≈0.35N）：二维 DP 占 45,233 KB (44.2 MB)，基础 Myers 占 25,198 KB (24.6 MB)，双向分治版仅 206.4 KB (0.20 MB，压缩 219 倍)！"
             ],
             bg_color=RGBColor(236, 253, 245), border_color=C_GREEN, title_color=C_GREEN, font_size=11.8)

    add_notes(s_bench, (
        "【第 12 页演讲逐字稿（约 50 秒）】\n"
        "这是我在本机上用 Python time 和 tracemalloc 对三种实现从 300 到 2400 行的实测数据：\n"
        "2400 行仅改 12 处时，二维 DP 要 275.3 毫秒，Myers 只需 0.13 毫秒——快了 2131 倍；\n"
        "而在大段重构时，双向分治版内存占用从 44.2 MB 直降到 206.4 KB，压缩 219 倍，验证了前面所有的复杂度推导。"
    ))

    # ==========================================================================
    # PART 2 (SLIDE 13): Worst Case & Engineering Evolution in Git (Patience / Histogram)
    # ==========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, "第二部分 · 理论与进阶 | 12",
               "退化边界 (Worst Case) 与 Git 工业界防线：启发式截断与唯一行锚点",
               "从复杂度兜底截断 (Cutoff) 到 Patience / Histogram Diff 锚点分治", 13)

    s10.shapes.add_picture(f7, Inches(0.55), Inches(1.55), width=Inches(12.2))

    add_card(s10, 0.55, 5.60, 12.2, 1.40,
             "🛡️ 工业界工程防护双保险（git/xdiff）",
             [
                 "• 唯一行锚点（Patience / Histogram）：先抽取全篇仅出现 1 次的函数头求 LIS 当定海神针切分网格，避开连续 `{}` 迷魂阵，再分段跑 Myers。",
                 "• 防卡死截断（xdl_split）：整文件重写 D → N+M 时触发启发式阈值提前切分，用极微小非最优性换取稳态响应。"
             ],
             bg_color=RGBColor(239, 246, 255), border_color=C_BLUE, title_color=C_BLUE, font_size=11.8)

    add_notes(s10, (
        "【第 13 页演讲逐字稿（约 50 秒）】\n"
        "实际源码库里还要防两个极端问题：第一，整个文件被删光重写（D ≈ N+M）时，Git 用 `too_expensive` 阈值做启发式截断防卡死；"
        "第二，代码里有大量重复空行和 `{}` 容易诱发错位滑行，因此 Git 提供了 `--patience` 和 `--histogram` 模式，先抓仅出现一次的核心行作为 LIS 锚点切分网格，再分别跑 Myers。"
    ))

    # ==========================================================================
    # PART 2 (SLIDE 14): Takeaways & Summary
    # ==========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    add_header(s11, "第二部分 · 理论与进阶 | 13",
               "总结与升华：Myers 差分算法蕴含的三大算法设计思想",
               "从坐标变换降维、输出敏感性设计到多算法范式协同的方法论启示", 14)

    add_card(s11, 0.55, 1.65, 3.85, 3.55,
             "1️⃣ 坐标系转换降维\n(x, y) ➔ (D, k)",
             [
                 "• 按 (i, j) 填表空间必为 N×M。",
                 "• 把目标步数 D 拿来做外层循环，沿车道 k=x-y 只记最大 x = V[k]。",
                 "• 设计启示：DP 状态过大时，尝试交换“待求目标值”与“状态下标”。"
             ],
             bg_color=C_WHITE, border_color=C_BLUE, title_color=C_BLUE, font_size=12.0)

    add_card(s11, 4.72, 1.65, 3.85, 3.55,
             "2️⃣ 输出敏感性设计\n压路机 ➔ 手电筒",
             [
                 "• 日常源码差异 D 远小于文件规模 N。",
                 "• 依靠对角线滑梯 (Snake) 跳过相同块，使耗时随差异 D 增长。",
                 "• 设计启示：利用真实数据局部性特征实现平均性能跃迁。"
             ],
             bg_color=C_WHITE, border_color=C_GREEN, title_color=C_GREEN, font_size=12.0)

    add_card(s11, 8.9, 1.65, 3.85, 3.55,
             "3️⃣ 多范式协同交响\nDP + 贪心 + 双向分治",
             [
                 "• 择优换道：动态规划最优子结构",
                 "• 对角滑行：贪心选择性质（剪接保证）",
                 "• 路径回溯：双向分治将空间压至 O(N+M)",
                 "• 设计启示：工业级核心往往是多种课本范式的组合交响。"
             ],
             bg_color=C_WHITE, border_color=C_AMBER, title_color=C_AMBER, font_size=12.0)

    add_card(s11, 0.55, 5.45, 12.2, 1.55,
             "🎓 参考文献与交互式演示系统索引 (References & Live Demo)",
             [
                 "• Eugene W. Myers, \"An O(ND) Difference Algorithm and Its Variations\", Algorithmica (1986), Vol. 1, pp. 251–266.",
                 "• D. S. Hirschberg, \"A Linear Space Algorithm for Computing Maximal Common Subsequences\", CACM (1975).",
                 "• Live Sandbox: `interactive_myers_visualizer.html`（内置 3 组算例与实时双向会师波前动画）—— 感谢聆听，欢迎交流！"
             ],
             bg_color=C_NAVY, border_color=C_BLUE, title_color=RGBColor(251, 191, 36), font_size=11.8)

    ref_tb = s11.shapes[-1]
    for p in ref_tb.text_frame.paragraphs[1:]:
        p.font.color.rgb = RGBColor(226, 232, 240)

    add_notes(s11, (
        "【第 14 页演讲逐字稿（约 45 秒）】\n"
        "最后总结今天分享的三大算法思维启发：坐标系转换降维、输出敏感性设计，以及 DP+贪心+双向分治的多范式组合。\n"
        "现场如想体验任意字符串推演或双向波前会师过程，可以切到配套的 `interactive_myers_visualizer.html` 演示网页。感谢老师和同学们的聆听，欢迎交流提问！"
    ))

    # ==========================================================================
    # APPENDIX (SLIDE 15): Further Discussion & Implementation Details (Q&A)
    # ==========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    add_header(s14, "附录 | 延伸探讨与工程细节 (Q&A)",
               "延伸探讨：关于贪心支配性、双向坐标对齐与行哈希的三个关键问题",
               "Further Discussion & Implementation Details in Git (xdiff)", 15)

    add_card(s14, 0.55, 1.6, 12.2, 1.65,
             "❓ 探讨问题 1：为什么同一对角线上只留最远点 P₁、顺着对角线贪心滑到底绝不会漏掉最优解？",
             [
                 "• 严格证明（L-边界剪接引理）：在同一车道 k 上，领先点 P₁(x₁,y₁) 与落后点 P₂(x₂,y₂) 满足 x₁-x₂ = y₁-y₂ = Δ > 0。",
                 "• 从 P₂ 到右下角终点 (N,M) 的路径必跨 L={x=x₁ 或 y=y₁} 于点 Q，d(P₂→Q) ≥ d(P₁→Q)，将 Q 后路径剪接到 P₁ 上必不劣。"
             ],
             bg_color=C_WHITE, border_color=C_BLUE, title_color=C_BLUE, font_size=11.5)

    add_card(s14, 0.55, 3.42, 12.2, 1.75,
             "❓ 探讨问题 2：双向搜索寻找中间蛇时，正反向坐标系怎么对齐？奇偶性如何判定碰撞？",
             [
                 "• 坐标对齐：正向车道定义 k = x - y（起点 (0,0)）；反向车道定义 bk = (N-x) - (M-y) = Δ - k，其中 Δ = N - M。",
                 "• 会师碰撞判定：总差异为奇数时在正向第 d 轮检查 x + V_bwd[Δ - k] ≥ N；为偶数时在反向第 d 轮检查并锁定 Middle Snake。"
             ],
             bg_color=C_WHITE, border_color=C_GREEN, title_color=C_GREEN, font_size=11.5)

    add_card(s14, 0.55, 5.34, 12.2, 1.65,
             "❓ 探讨问题 3：Git 跑 Myers 时如果某一行代码有 200 个字符，每次比较 A[x]==B[y] 是 O(L) 吗？",
             [
                 "• 哈希整数化预处理（git/xdiff/xprepare.c）：进入 Myers 前先对两文件每行计算哈希并分桶，将整行文本映射为 32 位整数 ID。",
                 "• 因此核心内层滑行循环 `while A[x] == B[y]` 实际执行的只是 CPU 单周期 O(1) 整数相等比较 (`ha1[x] == ha2[y]`)！"
             ],
             bg_color=C_WHITE, border_color=C_AMBER, title_color=C_AMBER, font_size=11.5)

    add_notes(s14, (
        "【第 15 页：延伸探讨 Q&A 备用页】\n"
        "用于回应互动提问中的三个核心细节：L-边界剪接证明、双向搜索正反向对角线偏移 Δ=N-M 的奇偶碰撞判定，以及 Git 行级预哈希 O(1) 比较优化。"
    ))

    pptx_path = os.path.join(OUT_DIR, "Myers_Diff_Algorithm_Presentation.pptx")
    prs.save(pptx_path)
    print(f"Saved PPTX to: {pptx_path}")


if __name__ == "__main__":
    build_presentation()
