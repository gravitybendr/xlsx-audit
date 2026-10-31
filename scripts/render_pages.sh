#!/usr/bin/env bash
# Render every HTML page in a folder to a PNG next to it, with headless Chromium.
# Runs inside the mcr.microsoft.com/playwright image; scripts/make_images.py starts it.
#
# Each page measures its #frame element after load and writes the size to
# <body data-size="WxH">. A first pass reads that size from the DOM, the second
# takes the screenshot at exactly that size. <html data-scale="N"> sets the
# device pixel ratio (default 2). The background is transparent unless the page
# paints one.
set -euo pipefail
dir=$(realpath "$1")
chrome=$(ls -d /ms-playwright/chromium_headless_shell-*/chrome-headless-shell-linux64/chrome-headless-shell | head -n1)
flags=(--no-sandbox --disable-gpu --hide-scrollbars --font-render-hinting=none)
for page in "$dir"/*.html; do
  out="${page%.html}.png"
  scale=$(grep -o 'data-scale="[0-9.]*"' "$page" | head -n1 | grep -o '[0-9.]*' || true)
  scale=${scale:-2}
  size=$("$chrome" "${flags[@]}" --window-size=2600,2600 --virtual-time-budget=3000 --dump-dom "file://$page" 2>/dev/null \
    | grep -o 'data-size="[0-9]*x[0-9]*"' | head -n1 | grep -o '[0-9]*x[0-9]*')
  w=${size%x*}; h=${size#*x}
  "$chrome" "${flags[@]}" --force-device-scale-factor="$scale" --default-background-color=00000000 \
    --window-size="$w,$h" --virtual-time-budget=3000 --screenshot="$out" "file://$page" >/dev/null 2>&1
  echo "$(basename "$out") ${w}x${h} @${scale}x"
done
