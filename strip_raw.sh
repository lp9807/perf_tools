#!/bin/bash

# Check if a filename was provided
if [ -z "$1" ]; then
    echo "Error: Please provide an input filename."
    echo "Usage: $0 <filename_raw.ext>"
    exit 1
fi

INPUT_FILE="$1"

# Generate output name by removing the exact string "_raw"
OUTPUT_FILE="${INPUT_FILE/_raw/}"

# Inform the user
echo "Input:  $INPUT_FILE"
echo "Output: $OUTPUT_FILE"

# Your processing command goes here (e.g., mv, cp, or awk)
awk '/min,/ {print; exit}' $INPUT_FILE > $OUTPUT_FILE
cat $INPUT_FILE | grep '^[0-9]' >> $OUTPUT_FILE
echo "Processing Done."