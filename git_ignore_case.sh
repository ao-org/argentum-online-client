#!/bin/bash

set -e

process_file() {
    local file=$1
    ORIGFILE=$(mktemp)
    PATCHFILE=$(mktemp)
    PROCESSEDFILE=$(mktemp)

    # Extract the original version of the file from git
    git show :"$file" > "$ORIGFILE"
    
    # Normalize the case for both the original and current files to avoid case-only differences
    awk '{print tolower($0)}' "$ORIGFILE" > "$PROCESSEDFILE"
    awk '{print tolower($0)}' "$file" > "$file.lower"

    # Generate diff excluding case changes
    diff --strip-trailing-cr "$PROCESSEDFILE" "$file.lower" > "$PATCHFILE" || true

    # If there are meaningful changes, apply them to the original file
    if [ -s "$PATCHFILE" ]; then
        patch -s "$ORIGFILE" < "$PATCHFILE"
        cp "$ORIGFILE" "$file"
    fi

    # Clean up
    rm "$ORIGFILE" "$PATCHFILE" "$PROCESSEDFILE" "$file.lower"
    unix2dos --quiet "$file"
}

# Process .frm, .bas, .cls, and .Dsr files
for file in $(git status --porcelain | grep -E "^[ AM]" | grep -E "\.(frm|bas|cls|Dsr)$"); do
    process_file "$file"
done

# Process .vbp files
for file in $(git status --porcelain | grep -E "^[ AM]" | grep -E "\.vbp$"); do
    process_file "$file"
done
