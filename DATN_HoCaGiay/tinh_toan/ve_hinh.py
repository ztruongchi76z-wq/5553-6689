# -*- coding: utf-8 -*-
"""Kiểu hình vẽ thống nhất cho toàn bộ đồ án (font Liberation Serif ~ Times New Roman)."""
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.tri as mtri
import numpy as np

HINH_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'hinh')
os.makedirs(HINH_DIR, exist_ok=True)

plt.rcParams.update({
    'font.family': 'Liberation Serif', 'font.size': 10, 'axes.titlesize': 11, 'axes.labelsize': 10,
    'legend.fontsize': 9, 'xtick.labelsize': 9, 'ytick.labelsize': 9, 'figure.dpi': 110,
    'savefig.dpi': 220, 'axes.grid': True, 'grid.alpha': 0.3, 'grid.linestyle': '--',
})

MAU_VL = {1: '#d9b38c', 2: '#f2e394', 3: '#c2a378', 4: '#f2e394', 5: '#9e9e9e', 6: '#cfa57a',
          7: '#b5c99a', 8: '#9fb3c8', 9: '#8a8fb0', 10: '#e07a5f', 11: '#e9c46a', 12: '#bfbfbf',
          21: '#d4a373', 22: '#a3b18a', 23: '#8d99ae', 24: '#e76f51', 0: '#ff0000'}


def luu(fig, ten):
    p = os.path.join(HINH_DIR, ten)
    fig.savefig(p, bbox_inches='tight')
    plt.close(fig)
    return p


def ve_vat_lieu(ax, xy, tri, mat, ten_vl, xlim=None, ylim=None, nhan=True):
    from matplotlib.collections import PolyCollection
    polys = xy[tri]
    cols = [MAU_VL.get(m, '#dddddd') for m in mat]
    ax.add_collection(PolyCollection(polys, facecolors=cols, edgecolors='none', linewidths=0))
    if nhan:
        import matplotlib.patches as mpatches
        hs = [mpatches.Patch(color=MAU_VL.get(m, '#ddd'), label=ten_vl[m]) for m in sorted(set(mat))]
        ax.legend(handles=hs, loc='lower right', fontsize=7.5, ncol=2, framealpha=0.9)
    if xlim:
        ax.set_xlim(*xlim)
    if ylim:
        ax.set_ylim(*ylim)
    ax.set_aspect('equal')


def ve_luoi(ax, xy, tri, lw=0.15):
    ax.triplot(mtri.Triangulation(xy[:, 0], xy[:, 1], tri), lw=lw, color='k', alpha=0.5)
