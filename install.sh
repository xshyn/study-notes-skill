#!/usr/bin/env bash
# Install the study-notes skill for one or more AI agents.
#   curl -fsSL https://raw.githubusercontent.com/xshyn/study-notes-skill/main/install.sh | bash
#   ./install.sh --agent claude|codex|cursor|gemini|opencode|all   [--project]   [--dir PATH]   [--no-deps]
set -euo pipefail

REPO="https://github.com/xshyn/study-notes-skill"
AGENT="claude"; SCOPE="user"; CUSTOM=""; DEPS=1
while [ $# -gt 0 ]; do
  case "$1" in
    --agent) AGENT="$2"; shift 2;;
    --project) SCOPE="project"; shift;;
    --dir) CUSTOM="$2"; shift 2;;
    --no-deps) DEPS=0; shift;;
    -h|--help) sed -n '2,5p' "$0"; exit 0;;
    *) echo "unknown option $1"; exit 1;;
  esac
done

SRC_DIR="$(cd "$(dirname "${BASH_SOURCE[0]:-$0}")" 2>/dev/null && pwd || true)"
if [ ! -f "$SRC_DIR/skills/study-notes/SKILL.md" ]; then
  TMP="$(mktemp -d)"; echo "→ downloading $REPO"
  git clone --depth 1 "$REPO" "$TMP/repo" >/dev/null 2>&1
  SRC_DIR="$TMP/repo"
fi
SKILL="$SRC_DIR/skills/study-notes"

target_for() {
  local base; [ "$SCOPE" = "project" ] && base="$PWD" || base="$HOME"
  case "$1" in
    claude)   echo "$base/.claude/skills/study-notes";;
    codex)    echo "$base/.codex/skills/study-notes";;
    cursor)   echo "$base/.cursor/skills/study-notes";;
    gemini)   echo "$base/.gemini/skills/study-notes";;
    opencode) echo "$base/.config/opencode/skills/study-notes";;
    agents)   echo "$base/.agents/skills/study-notes";;
  esac
}

if [ -n "$CUSTOM" ]; then TARGETS=("$CUSTOM/study-notes")
elif [ "$AGENT" = "all" ]; then TARGETS=(); for a in claude codex cursor gemini opencode agents; do TARGETS+=("$(target_for $a)"); done
else TARGETS=("$(target_for "$AGENT")"); fi

for T in "${TARGETS[@]}"; do
  mkdir -p "$(dirname "$T")"; rm -rf "$T"
  cp -R "$SKILL" "$T"; rm -rf "$T/evals"
  echo "✓ installed → $T"
done

if [ "$DEPS" = 1 ]; then
  echo "→ installing Python dependencies"
  python3 -m pip install -q -r "$SKILL/requirements.txt" 2>/dev/null \
    || python3 -m pip install -q --user -r "$SKILL/requirements.txt" 2>/dev/null \
    || python3 -m pip install -q --break-system-packages -r "$SKILL/requirements.txt" \
    || echo "! pip install failed — run: pip install -r $SKILL/requirements.txt"
fi
command -v soffice >/dev/null || echo "ℹ optional: install LibreOffice for visual previews (render_preview.py)"
command -v mmdc >/dev/null || echo "ℹ optional: npm i -g @mermaid-js/mermaid-cli to render diagrams"
echo "Done. Ask your agent: \"make study notes from <url or file>\""
