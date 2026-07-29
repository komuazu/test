#!/usr/bin/env python3
"""Build a multi-scene music video from stills + an audio track.

Each scene gets its own Ken Burns move; scenes are joined with crossfades whose
timing is chosen to land on the song's structure. The joined chain is graded,
grained and vignetted once, then a looping sakura-petal overlay is composited on
top (at full strength over the spring scenes, softened over the city scenes).

Edit SCENES below to re-cut the video; run with --preview N to render only the
first N seconds while you dial the look in.
"""
import argparse
import os
import shlex
import subprocess
import sys

W, H, FPS = 1080, 1920, 30
XF = 1.2                      # crossfade duration, seconds
SRC_SCALE = 1.5               # working resolution multiplier for Ken Burns
HERE = os.path.dirname(os.path.abspath(__file__))
PHOTOS = os.path.join(HERE, "photos")
FONT = "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf"

# (file, seconds_visible, motion, transition_into_next)
# Cuts land on the song's sections: quiet intro, build, chorus 1, verse 2,
# chorus 2, turn, final chorus, outro.
SCENES = [
    # Scenes 1 and 2 share a composition, so they bloom through white rather
    # than dissolving — a straight dissolve reads as a double exposure.
    ("img01_a224b8434f19.webp", 33, "push",  "fadewhite"),  # 0:00 quiet intro
    ("img08_c135c593635d.webp", 17, "up",    "fade"),       # 0:33 build
    ("img02_84067a3bff92.webp", 20, "pull",  "dissolve"),   # 0:50 chorus 1
    ("img06_5410be68cd3a.webp", 33, "panL",  "fade"),       # 1:10 verse 2
    ("scene_closeup.jpg",       27, "push",  "dissolve"),   # 1:43 chorus 2
    ("img07_6561d165fa09.webp", 18, "panR",  "fadeblack"),  # 2:10 turn
    ("img05_a3c97fc05f1a.webp", 25, "pull",  "fade"),       # 2:28 final chorus
    ("img01_a224b8434f19.webp", 10, "down",  None),         # 2:53 outro bookend
]

# Scene indices that keep the petals at full strength (the sakura scenes).
SPRING = {0, 1, 7}


def motion_filter(kind, frames):
    """Return the zoompan z/x/y expressions for a Ken Burns move."""
    f = max(frames - 1, 1)
    cx, cy = "iw/2-(iw/zoom/2)", "ih/2-(ih/zoom/2)"
    if kind == "push":
        return f"1.02+0.13*on/{f}", cx, cy
    if kind == "pull":
        return f"1.15-0.13*on/{f}", cx, cy
    if kind == "panL":
        return "1.10", f"(iw-iw/zoom)*(0.65-0.30*on/{f})", cy
    if kind == "panR":
        return "1.10", f"(iw-iw/zoom)*(0.35+0.30*on/{f})", cy
    if kind == "up":
        return f"1.06+0.06*on/{f}", cx, f"(ih-ih/zoom)*(0.60-0.22*on/{f})"
    if kind == "down":
        return f"1.06+0.06*on/{f}", cx, f"(ih-ih/zoom)*(0.30+0.22*on/{f})"
    raise ValueError(f"unknown motion: {kind}")


def build(audio, out, petals, preview=None, endcard="Beside The Door"):
    scenes = list(SCENES)
    total = sum(s[1] for s in scenes)

    if preview:                      # trim the scene list down for a fast look
        kept, acc = [], 0
        for f, d, m, t in scenes:
            if acc >= preview:
                break
            kept.append((f, min(d, preview - acc), m, t))
            acc += d
        scenes, total = kept, min(preview, total)

    n = len(scenes)
    sw, sh = int(W * SRC_SCALE), int(H * SRC_SCALE)

    inputs, chains = [], []
    for i, (fname, vis, motion, _) in enumerate(scenes):
        # Each scene runs XF longer than it is visible so it has material to
        # crossfade into the next one.
        clip = vis + (XF if i < n - 1 else 0)
        frames = int(round(clip * FPS))
        z, x, y = motion_filter(motion, frames)
        inputs += ["-loop", "1", "-framerate", str(FPS), "-t", f"{clip:.3f}",
                   "-i", os.path.join(PHOTOS, fname)]
        chains.append(
            f"[{i}:v]scale={sw}:{sh}:force_original_aspect_ratio=increase,"
            f"crop={sw}:{sh},"
            f"zoompan=z='{z}':x='{x}':y='{y}':d=1:s={W}x{H}:fps={FPS},"
            f"setsar=1,format=yuv420p[s{i}]"
        )

    ai, pi = n, n + 1              # audio and petal input indices
    inputs += ["-i", audio]
    inputs += ["-stream_loop", "-1", "-t", f"{total:.3f}", "-i", petals]

    # Crossfade the scenes together, tracking the running duration so each
    # xfade starts XF before the current chain ends.
    cur, dur = "[s0]", scenes[0][1] + (XF if n > 1 else 0)
    for i in range(1, n):
        trans = scenes[i - 1][3] or "fade"
        off = dur - XF
        nxt = f"[x{i}]"
        chains.append(
            f"{cur}[s{i}]xfade=transition={trans}:duration={XF}:"
            f"offset={off:.3f}{nxt}"
        )
        cur = nxt
        dur = dur + scenes[i][1] + (XF if i < n - 1 else 0) - XF

    # One grade for the whole film, then the petals.
    chains.append(
        f"{cur}eq=contrast=1.05:saturation=1.08:brightness=0.006:gamma=0.99,"
        f"vignette=PI/5,noise=alls=6:allf=t+u[graded]"
    )
    chains.append(f"[{pi}:v]format=rgba,split=2[pf][pd]")
    chains.append(f"[pd]colorchannelmixer=aa=0.42[psoft]")

    # Petals at full strength over the spring scenes, softened over the city.
    spring_ranges, city_ranges, t = [], [], 0.0
    for i, (_, vis, _, _) in enumerate(scenes):
        (spring_ranges if i in SPRING else city_ranges).append((t, t + vis))
        t += vis

    def enable(ranges):
        if not ranges:
            return "0"
        return "+".join(f"between(t,{a:.2f},{b:.2f})" for a, b in ranges)

    chains.append(
        f"[graded][pf]overlay=0:0:enable='{enable(spring_ranges)}'[v1]"
    )
    chains.append(
        f"[v1][psoft]overlay=0:0:enable='{enable(city_ranges)}'[v2]"
    )

    vf = (f"[v2]fade=t=in:st=0:d=2,"
          f"fade=t=out:st={total - 2.5:.2f}:d=2.5")
    if endcard and not preview:
        esc = endcard.replace(":", r"\:").replace("'", r"\'")
        start = total - 9
        vf += (f",drawtext=fontfile='{FONT}':text='{esc}':fontcolor=white@0.9:"
               f"fontsize=64:x=(w-text_w)/2:y=h*0.46:"
               f"shadowcolor=black@0.45:shadowx=2:shadowy=2:"
               f"alpha='if(lt(t,{start}),0,"
               f"if(lt(t,{start + 1.5}),(t-{start})/1.5,"
               f"if(lt(t,{total - 3.0}),1,"
               f"max(0,1-(t-{total - 3.0})/2))))'")
    chains.append(vf + "[v]")
    chains.append(
        f"[{ai}:a]afade=t=in:st=0:d=1.5,"
        f"afade=t=out:st={total - 4:.2f}:d=4[a]"
    )

    cmd = (["ffmpeg", "-y", "-loglevel", "error", "-stats"] + inputs +
           ["-filter_complex", ";".join(chains),
            "-map", "[v]", "-map", "[a]", "-t", f"{total:.3f}",
            "-c:v", "libx264", "-preset", "fast", "-crf", "20",
            "-pix_fmt", "yuv420p", "-r", str(FPS),
            "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart", out])
    print("scenes:", n, " total:", f"{total:.1f}s")
    print(" ".join(shlex.quote(c) for c in cmd)[:400], "...\n")
    return subprocess.call(cmd)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--audio", required=True)
    ap.add_argument("--petals", required=True, help="looping petal overlay clip")
    ap.add_argument("--out", required=True)
    ap.add_argument("--preview", type=float, default=None,
                    help="render only the first N seconds")
    args = ap.parse_args()
    sys.exit(build(args.audio, args.out, args.petals, args.preview))


if __name__ == "__main__":
    main()
