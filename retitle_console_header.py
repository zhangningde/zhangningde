#!/usr/bin/env python3
"""Replace console center title 5G+IMS AI远程协助系统 -> AI+IMS智能融合调度指挥系统."""
from __future__ import annotations
import argparse
from pathlib import Path
import numpy as np
from PIL import Image, ImageDraw, ImageFont

NEW_TITLE = "AI+IMS智能融合调度指挥系统"
FONT_PATH = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"


def replace_center_title(im: Image.Image, min_wr=0.14, max_wr=0.72):
    im = im.convert("RGBA")
    W, H = im.size
    y1 = min(80, max(36, int(H * 0.14)))
    cx0, cx1 = int(W * 0.12), int(W * 0.88)
    band = np.array(im.crop((cx0, 0, cx1, y1))).astype(np.float32)[:, :, :3]
    lum = band.mean(axis=2)
    text_mask = lum > 155
    if text_mask.sum() < 60:
        text_mask = lum > 125
    if text_mask.sum() < 40:
        return im, "no-text"
    ys, xs = np.where(text_mask)
    ty0, ty1 = max(0, int(ys.min()) - 3), min(y1, int(ys.max()) + 5)
    tx0, tx1 = cx0 + int(xs.min()), cx0 + int(xs.max())
    mid = (tx0 + tx1) / 2
    wr = (tx1 - tx0) / W
    if abs(mid - W / 2) > W * 0.15:
        return im, f"not-centered({wr:.3f})"
    if wr < min_wr or wr > max_wr:
        return im, f"width({wr:.3f})"
    arr = np.array(im)
    sample = arr[ty0:ty1, max(0, tx0 - 40) : max(1, tx0), :3].reshape(-1, 3)
    dark = sample[sample.mean(axis=1) < 90] if len(sample) else sample
    bg = tuple(int(v) for v in (np.median(dark, axis=0) if len(dark) else (16, 40, 76)))
    out = Image.fromarray(arr)
    draw = ImageDraw.Draw(out)
    box_h = max(14, ty1 - ty0)
    font_size = max(11, min(int(box_h * 0.92), int(W * 0.022)))
    font = ImageFont.truetype(FONT_PATH, font_size)
    while font.getlength(NEW_TITLE) > W * 0.72 and font_size > 10:
        font_size -= 1
        font = ImageFont.truetype(FONT_PATH, font_size)
    tw = font.getlength(NEW_TITLE)
    ntx0 = max(int(W * 0.08), int((W - tw) / 2 - 18))
    ntx1 = min(int(W * 0.92), int((W + tw) / 2 + 18))
    draw.rectangle([min(tx0 - 6, ntx0), ty0 - 1, max(tx1 + 6, ntx1), ty1 + 1], fill=bg + (255,))
    draw.text(((W - tw) / 2, ty0 + max(0, (box_h - font_size) / 2 - 1)), NEW_TITLE, font=font, fill=(245, 250, 255, 255))
    return out, f"replaced({wr:.3f})"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("paths", nargs="+", type=Path)
    ap.add_argument("-o", "--outdir", type=Path, default=None)
    args = ap.parse_args()
    for p in args.paths:
        out, mode = replace_center_title(Image.open(p))
        dest = (args.outdir / p.name) if args.outdir else p
        if args.outdir:
            args.outdir.mkdir(parents=True, exist_ok=True)
        if mode.startswith("replaced"):
            out.convert("RGB").save(dest)
        print(p.name, mode, "->", dest)


if __name__ == "__main__":
    main()
