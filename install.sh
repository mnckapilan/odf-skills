#!/usr/bin/env bash
set -euo pipefail

REPO="mnckapilan/odf-skills"
SKILLS_DIR="${CLAUDE_SKILLS_DIR:-$HOME/.claude/skills}"
ALL_SKILLS=(ods odt)

usage() {
    echo "Usage: install.sh [skill ...]"
    echo "       install.sh ods odt    # install specific skills"
    echo "       install.sh            # install all skills"
    echo ""
    echo "Available skills: ${ALL_SKILLS[*]}"
    exit 1
}

# Default to all skills if none specified
if [ $# -eq 0 ]; then
    INSTALL=("${ALL_SKILLS[@]}")
else
    INSTALL=("$@")
fi

# Validate skill names
for skill in "${INSTALL[@]}"; do
    valid=false
    for s in "${ALL_SKILLS[@]}"; do
        [ "$skill" = "$s" ] && valid=true && break
    done
    if [ "$valid" = false ]; then
        echo "error: unknown skill '$skill'. Available: ${ALL_SKILLS[*]}" >&2
        usage
    fi
done

# Download repo archive into a temp dir
TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT

echo "Downloading odf-skills..."
curl -fsSL "https://github.com/$REPO/archive/refs/heads/main.tar.gz" \
    | tar -xz -C "$TMP" --strip-components=1

mkdir -p "$SKILLS_DIR"

for skill in "${INSTALL[@]}"; do
    dest="$SKILLS_DIR/$skill"
    if [ -d "$dest" ]; then
        echo "Updating  $skill → $dest"
    else
        echo "Installing $skill → $dest"
    fi
    rm -rf "$dest"
    # Install only SKILL.md and scripts/ — tests and evals are not needed at runtime
    mkdir -p "$dest/scripts"
    cp "$TMP/$skill/SKILL.md"         "$dest/SKILL.md"
    cp "$TMP/$skill/scripts/$skill.py" "$dest/scripts/$skill.py"
done

echo ""
echo "Done. Invoke with /${INSTALL[*]// //} in your agent."
