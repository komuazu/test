#!/usr/bin/env python3
"""Generate a seamless-looping, transparent falling-sakura-petal overlay.

Outputs a PNG sequence (RGBA) that make_music_video.sh encodes into a looping
overlay clip. Every petal's motion is periodic in t/T, so the clip loops with
no visible seam when repeated over a longer track.
"""
import argparse
import math
import os

import numpy as np
from PIL import Image, ImageFilter


def petal_sprite(size, color, alpha, blur):
    """Draw one cherry-blossom petal: narrow at the base, notched at the tip.

    `blur` softens the sprite so distant petals sit back and the nearest ones
    read as out-of-focus foreground bokeh.
    """
    h = max(8, int(size))
    w = max(6, int(size * 0.72))          # petals are taller than they are wide
    # Normalized petal space: u across (-1..1), t along (0 = base, 1 = tip).
    row, col = np.mgrid[0:h, 0:w]
    u = (col / (w - 1)) * 2 - 1
    t = 1 - row / (h - 1)                 # image rows run downward

    # Half-width profile: pinched to a point at the base, broad at the tip.
    half = np.power(np.clip(t, 0, 1), 0.60) * \
        np.power(np.clip(1 - t ** 12, 0, 1), 0.22)
    half = half / max(half.max(), 1e-6)

    # The tip dips inward in the middle — the notch that reads as "sakura".
    # Keep it shallow; a deep notch turns the petal into a heart.
    tip = 1 - 0.07 * np.exp(-((u / 0.32) ** 2))

    inside = (np.abs(u) <= half) & (t <= tip)
    # Soften the rim so the edge is not aliased before the blur.
    edge = np.clip((half - np.abs(u)) * w * 0.5, 0, 1)
    a = (inside * edge * alpha).astype(np.uint8)

    rgba = np.zeros((h, w, 4), dtype=np.uint8)
    rgba[..., 0], rgba[..., 1], rgba[..., 2] = color
    rgba[..., 3] = a
    img = Image.fromarray(rgba, "RGBA")
    if blur > 0:
        img = img.filter(ImageFilter.GaussianBlur(blur))
    return img


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--width", type=int, default=1080)
    ap.add_argument("--height", type=int, default=1920)
    ap.add_argument("--fps", type=int, default=30)
    ap.add_argument("--seconds", type=float, default=10.0, help="loop length")
    ap.add_argument("--count", type=int, default=70, help="number of petals")
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--seed", type=int, default=7)
    args = ap.parse_args()

    os.makedirs(args.outdir, exist_ok=True)
    rng = np.random.default_rng(args.seed)
    W, H, FPS, T = args.width, args.height, args.fps, args.seconds
    frames = int(round(T * FPS))
    margin = 120  # vertical wrap margin so petals enter/exit smoothly

    # Soft pink tones, warm to cool, for a little color variety.
    tones = [(0xF7, 0xC5, 0xD6), (0xF6, 0xE0, 0xE7), (0xEC, 0xA9, 0xC4),
             (0xFA, 0xD8, 0xC8)]

    petals = []
    for _ in range(args.count):
        depth = rng.uniform(0.0, 1.0)                # 0 = far, 1 = near
        size = int(14 + depth * depth * 78)          # near petals much larger
        # Distant petals are faint; the very nearest go soft and translucent
        # so they read as foreground bokeh rather than stickers.
        base_alpha = int(55 + 120 * depth - 55 * max(0.0, depth - 0.75) / 0.25)
        blur = 0.4 + 3.6 * max(0.0, depth - 0.72) / 0.28 + (1 - depth) * 0.8
        color = tones[rng.integers(0, len(tones))]
        petals.append(
            dict(
                x0=rng.uniform(-60, W + 60),
                phase=rng.uniform(0, 1),             # vertical phase (loop)
                fall=rng.uniform(0.45, 1.25) * (0.5 + depth),
                sway_amp=rng.uniform(20, 70) * (0.4 + depth),
                sway_freq=rng.integers(1, 4),        # integer -> periodic
                sway_phase=rng.uniform(0, 2 * math.pi),
                spin=rng.integers(-2, 3),            # integer turns per T
                spin_phase=rng.uniform(0, 360),
                sprite=petal_sprite(size, color, base_alpha, blur),
            )
        )

    for f in range(frames):
        t = f / frames  # 0..1, periodic
        canvas = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        for p in petals:
            prog = (p["phase"] + t * p["fall"]) % 1.0
            y = prog * (H + 2 * margin) - margin
            x = p["x0"] + p["sway_amp"] * math.sin(
                2 * math.pi * (p["sway_freq"] * t) + p["sway_phase"]
            )
            angle = p["spin_phase"] + 360.0 * p["spin"] * t
            spr = p["sprite"].rotate(angle, resample=Image.BICUBIC, expand=True)
            canvas.alpha_composite(spr, (int(x - spr.width / 2), int(y - spr.height / 2)))
        canvas.save(os.path.join(args.outdir, f"petal_{f:04d}.png"))

    print(f"wrote {frames} petal frames to {args.outdir}")


if __name__ == "__main__":
    main()
