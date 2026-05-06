#!/bin/bash
# Build PDF from LaTeX source
# Usage: build_pdf.sh <input.tex> [output_dir]

set -euo pipefail

INPUT="$1"
OUTPUT_DIR="${2:-.}"

if [ ! -f "$INPUT" ]; then
  echo "Error: File not found: $INPUT" >&2
  exit 1
fi

BASENAME=$(basename "$INPUT" .tex)

# Detect LaTeX engine
if command -v lualatex &>/dev/null; then
  ENGINE="lualatex"
elif command -v xelatex &>/dev/null; then
  ENGINE="xelatex"
elif command -v pdflatex &>/dev/null; then
  ENGINE="pdflatex"
else
  echo "Error: No LaTeX engine found. Install one of: mactex, basictex, tinytex" >&2
  echo "  brew install --cask mactex    # full installation" >&2
  echo "  brew install --cask basictex  # minimal installation" >&2
  exit 1
fi

echo "Using engine: $ENGINE"

# Run twice for references/TOC
cd "$(dirname "$INPUT")"
TEXFILE=$(basename "$INPUT")

$ENGINE -interaction=nonstopmode -output-directory="$OUTPUT_DIR" "$TEXFILE"

# Check if bibtex/biber needed
if grep -q '\\bibliography\|\\addbibresource' "$TEXFILE"; then
  if command -v biber &>/dev/null; then
    biber "$OUTPUT_DIR/$BASENAME"
  elif command -v bibtex &>/dev/null; then
    bibtex "$OUTPUT_DIR/$BASENAME"
  fi
  $ENGINE -interaction=nonstopmode -output-directory="$OUTPUT_DIR" "$TEXFILE"
fi

# Second pass for cross-references
$ENGINE -interaction=nonstopmode -output-directory="$OUTPUT_DIR" "$TEXFILE"

echo "✅ Generated: $OUTPUT_DIR/$BASENAME.pdf"
