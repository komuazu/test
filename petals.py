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
from PIL import Image, ImageDraw


def petal_sprite(size, color, alpha):
    """Draw a single soft cherry-blossom petal on an RGBA tile."""
    s = size
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    pad = max(1, s // 8)
    # Elongated soft blossom: a rounded petal a little taller than it is wide.
    d.ellipse([pad, pad // 2, s - pad, s - pad // 2], fill=color + (alpha,))
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

    # Two soft pink tones for depth.
    tones = [(0xF7, 0xC5, 0xD6), (0xF3, 0xD9, 0xE4), (0xEC, 0xA9, 0xC4)]

    petals = []
    for _ in range(args.count):
        depth = rng.uniform(0.35, 1.0)               # far -> near
        size = int(18 + depth * 46)                  # px
        base_alpha = int(70 + depth * 150)
        color = tones[rng.integers(0, len(tones))]
        petals.append(
            dict(
                x0=rng.uniform(-40, W + 40),
                phase=rng.uniform(0, 1),             # vertical phase (loop)
                fall=rng.uniform(0.8, 1.6),          # loops per T (fall speed)
                sway_amp=rng.uniform(18, 60) * depth,
                sway_freq=rng.integers(1, 4),        # integer -> periodic
                sway_phase=rng.uniform(0, 2 * math.pi),
                spin=rng.integers(-2, 3),            # integer turns per T
                spin_phase=rng.uniform(0, 360),
                sprite=petal_sprite(size, color, base_alpha),
                size=size,
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
