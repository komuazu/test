# Music Video Maker

Turn a single still image + an audio track into a cinematic vertical
(1080×1920) music video: a slow Ken Burns push-in, drifting sakura petals, a
warm dusk color grade, subtle film grain, a vignette, and gentle fade in/out.
Your audio becomes the full soundtrack.

## Requirements

- `ffmpeg` / `ffprobe`
- Python 3 with `pillow` and `numpy` (`pip install pillow numpy`)
- A Japanese-capable font at
  `/usr/share/fonts/truetype/fonts-japanese-gothic.ttf` (only needed when you
  pass a title)

## Usage

```bash
./make_music_video.sh IMAGE AUDIO OUTPUT.mp4 [DURATION_SECONDS] [TITLE]
```

| Argument   | Meaning                                                       |
|------------|---------------------------------------------------------------|
| `IMAGE`    | Source photo — portrait orientation works best                |
| `AUDIO`    | mp3/wav/… — used in full as the soundtrack                     |
| `OUTPUT`   | Output `.mp4` path                                            |
| `DURATION` | *(optional)* cap length in seconds (default: full audio)      |
| `TITLE`    | *(optional)* overlay text, Japanese OK; omit for no title     |

### Examples

```bash
# Full-length video from your photo + song
./make_music_video.sh photo.jpg song.mp3 out.mp4

# 15-second preview with a title card
./make_music_video.sh photo.jpg song.mp3 preview.mp4 15 "遠くの港町"
```

## How it works

1. `petals.py` renders a 10-second, seamlessly looping transparent overlay of
   falling cherry-blossom petals (every petal's motion is periodic, so the loop
   has no seam when repeated over a longer track).
2. `make_music_video.sh` applies a Ken Burns zoom/drift to the still, grades and
   grains it, composites the looping petals, adds fades, and muxes in the audio.

## Tuning

- **Petals:** `petals.py` flags (`--count`, `--seconds`) or the `tones` list for
  petal color.
- **Motion / grade:** the `KB` (Ken Burns) and `GRADE` filter strings near the
  top of `make_music_video.sh`.
- **Quality/size:** `-crf` in `make_music_video.sh` (lower = higher quality,
  bigger file).
