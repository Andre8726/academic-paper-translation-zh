#!/usr/bin/env python3
"""Render a LaTeX/mathtext equation to a tight PNG for insertion into DOCX."""
from __future__ import annotations

import argparse
import os
import tempfile
from pathlib import Path

# Avoid writing matplotlib cache to ~/.matplotlib, which may be read-only.
os.environ.setdefault("MPLCONFIGDIR", str(Path(tempfile.gettempdir()) / "dsh-mpl-cache"))

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--tex", required=True, help="mathtext, normally wrapped in $...$")
    ap.add_argument("--out", required=True, help="output .png")
    ap.add_argument("--fontsize", type=float, default=17)
    ap.add_argument("--dpi", type=int, default=400)
    args = ap.parse_args()

    matplotlib.rcParams["mathtext.fontset"] = "stix"
    matplotlib.rcParams["font.family"] = "STIXGeneral"

    fig = plt.figure(figsize=(0.01, 0.01))
    text = fig.text(0, 0, args.tex, fontsize=args.fontsize, color="black")
    fig.canvas.draw()
    bb = text.get_window_extent(renderer=fig.canvas.get_renderer())
    w, h = bb.width / fig.dpi, bb.height / fig.dpi
    fig.set_size_inches(w + 0.08, h + 0.08)
    text.set_position((0.04 / (w + 0.08), 0.04 / (h + 0.08)))

    out = Path(args.out).expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(str(out), dpi=args.dpi, facecolor="white",
                bbox_inches="tight", pad_inches=0.03)
    plt.close(fig)
    print(out)


if __name__ == "__main__":
    main()
