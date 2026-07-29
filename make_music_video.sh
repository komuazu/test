#!/usr/bin/env bash
#
# make_music_video.sh — turn a still image + an audio track into a cinematic
# vertical music video (Ken Burns push-in, drifting sakura petals, warm dusk
# grade, film grain, vignette, gentle fades).
#
#   ./make_music_video.sh IMAGE AUDIO OUTPUT.mp4 [DURATION_SECONDS] [TITLE]
#
#   IMAGE     source photo (portrait works best)
#   AUDIO     mp3/wav/etc — becomes the full soundtrack
#   OUTPUT    output .mp4 path
#   DURATION  optional cap in seconds (default: full audio length)
#   TITLE     optional overlay text (Japanese OK); omit for none
#
set -euo pipefail

IMG="${1:?image path required}"
AUD="${2:?audio path required}"
OUT="${3:?output path required}"
DUR_OVERRIDE="${4:-}"
TITLE="${5:-}"

W=1080; H=1920; FPS=30
FONT="/usr/share/fonts/truetype/fonts-japanese-gothic.ttf"
HERE="$(cd "$(dirname "$0")" && pwd)"

for f in "$IMG" "$AUD"; do
  [ -f "$f" ] || { echo "ERROR: not found: $f" >&2; exit 1; }
done

# --- duration ---------------------------------------------------------------
ALEN=$(ffprobe -v error -show_entries format=duration \
        -of default=noprint_wrappers=1:nokey=1 "$AUD")
DUR="$ALEN"
if [ -n "$DUR_OVERRIDE" ]; then DUR="$DUR_OVERRIDE"; fi
DUR=$(printf '%.3f' "$DUR")
echo ">> duration: ${DUR}s  (audio ${ALEN}s)"

# --- petal overlay (seamless 10s loop) --------------------------------------
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT
echo ">> generating sakura petals..."
python3 "$HERE/petals.py" --width "$W" --height "$H" --fps "$FPS" \
        --seconds 10 --count 70 --outdir "$TMP/petals" >/dev/null
ffmpeg -y -loglevel error -framerate "$FPS" -i "$TMP/petals/petal_%04d.png" \
       -c:v qtrle "$TMP/petals.mov"

# --- fade timing ------------------------------------------------------------
FOUT_START=$(awk "BEGIN{printf \"%.3f\", $DUR-2.0}")
AOUT_START=$(awk "BEGIN{printf \"%.3f\", $DUR-3.0}")

# Ken Burns: slow push-in from 1.00 to ~1.15 with a gentle downward drift.
KB="scale=${W}*2:${H}*2:force_original_aspect_ratio=increase,crop=${W}*2:${H}*2,\
zoompan=z='min(1.0+0.00006*on,1.15)':d=1:x='iw/2-(iw/zoom/2)':\
y='ih/2-(ih/zoom/2)+sin(on/700)*60':s=${W}x${H}:fps=${FPS},setsar=1"

# Warm dusk grade + vignette + subtle grain.
GRADE="eq=contrast=1.06:saturation=1.12:brightness=0.008:gamma=0.98,\
vignette=PI/5,noise=alls=7:allf=t+u"

# Optional title, fading in at 1s and out at 6s.
TITLE_FILTER=""
if [ -n "$TITLE" ]; then
  ESC=$(printf '%s' "$TITLE" | sed "s/:/\\\\:/g; s/'/\\\\'/g")
  TITLE_FILTER=",drawtext=fontfile='${FONT}':text='${ESC}':fontcolor=white@0.92:\
fontsize=76:x=(w-text_w)/2:y=h*0.12:shadowcolor=black@0.5:shadowx=2:shadowy=2:\
alpha='if(lt(t,1),0,if(lt(t,2),(t-1),if(lt(t,5),1,if(lt(t,6),1-(t-5),0))))'"
fi

echo ">> rendering video..."
ffmpeg -y -loglevel error -stats \
  -loop 1 -framerate "$FPS" -t "$DUR" -i "$IMG" \
  -i "$AUD" \
  -stream_loop -1 -t "$DUR" -i "$TMP/petals.mov" \
  -filter_complex "\
    [0:v]${KB},${GRADE}[base]; \
    [base][2:v]overlay=0:0:format=auto[mix]; \
    [mix]fade=t=in:st=0:d=1.5,fade=t=out:st=${FOUT_START}:d=2${TITLE_FILTER}[v]; \
    [1:a]afade=t=in:st=0:d=1.2,afade=t=out:st=${AOUT_START}:d=3[a]" \
  -map "[v]" -map "[a]" -t "$DUR" \
  -c:v libx264 -preset medium -crf 19 -pix_fmt yuv420p -r "$FPS" \
  -c:a aac -b:a 192k -movflags +faststart \
  "$OUT"

echo ">> done: $OUT"
ffprobe -v error -show_entries format=duration:stream=width,height \
        -of default=noprint_wrappers=1 "$OUT" | sed 's/^/   /'
