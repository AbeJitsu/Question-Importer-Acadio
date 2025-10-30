#!/usr/bin/env python3
"""
Question CSV Transformer

What: Transforms question CSV format to intermediate format, then converts to XLSX
Why: Standardizes question format for consistent import processing
How: Reads source CSV, restructures columns, exports intermediate CSV, then converts to XLSX
"""

import csv
import pandas as pd
import subprocess
import sys
import re
from pathlib import Path

# Configuration - Easy to modify defaults
INPUT_CSV = "input/PREP-AL 4th Ed Question Excel Database 9-30-25.xlsx - Sheet1.csv"
OUTPUT_NAME = "output/upload_ready_questions_std"

def extract_source_prefix(filename):
    """
    Auto-detect source prefix from filename

    What: Extracts identifier prefix from input filename with type suffix
    Why: Automatically generates unique source IDs for different question sets
    How: Uses regex to extract pattern, then adds -ITP or -STD suffix based on filename

    Example: "PREP-AL 4th Ed Instructor..." -> "PREP-AL-ITP"
             "PREP-AL 4th Ed..." -> "PREP-AL-STD"
    """
    # Remove path and get just filename
    base_name = Path(filename).stem

    # Try to extract pattern like "PREP-AL" or "PREP-FL"
    match = re.match(r'^([A-Z]+-[A-Z]+)', base_name)
    if match:
        base_prefix = match.group(1)
    else:
        # Fallback: take first word
        base_prefix = base_name.split()[0] if base_name else "UNKNOWN"

    # Add suffix based on file type to ensure unique IDs
    if "Instructor" in base_name:
        return f"{base_prefix}-ITP"  # Instructor Test Pack
    else:
        return f"{base_prefix}-STD"  # Standard edition

    return base_prefix

def convert_answer_letter_to_number(letter):
    """
    Convert answer letter (a/b/c/d) to number (1/2/3/4)

    What: Maps letter-based answers to numeric format
    Why: Standardizes answer format for converter tool
    How: Simple case-insensitive mapping with fallback
    """
    letter_map = {
        'a': '1',
        'b': '2',
        'c': '3',
        'd': '4'
    }
    return letter_map.get(letter.lower().strip(), letter)

def transform_questions_csv():
    """
    Transform question CSV to intermediate format

    What: Reads source CSV and extracts question data
    Why: Converts various CSV formats to standardized intermediate format
    How: Parses rows, extracts fields, handles explanations, generates source IDs

    Input format:
    - Book Name, Question #, Question Stem, Answer A, Answer B, Answer C, Answer D,
      Correct Answer, Correct Answer Explanation

    Output format:
    - Question, Choice 1, Choice 2, Choice 3, Choice 4, Correct Answer, Source, Explanation
    """

    questions = []
    source_prefix = extract_source_prefix(INPUT_CSV)

    with open(INPUT_CSV, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)

        # Skip first 3 rows (title, blank, header)
        next(reader)  # Title row
        next(reader)  # Blank row
        next(reader)  # Header row

        for row in reader:
            # Skip empty rows
            if not row or len(row) < 8:
                continue

            # Extract question data by column index
            question_stem = row[2].strip() if len(row) > 2 else ''
            if not question_stem:
                continue

            answer_a = row[3].strip() if len(row) > 3 else ''
            answer_b = row[4].strip() if len(row) > 4 else ''
            answer_c = row[5].strip() if len(row) > 5 else ''
            answer_d = row[6].strip() if len(row) > 6 else ''
            correct_answer = row[7].strip() if len(row) > 7 else ''
            explanation = row[8].strip() if len(row) > 8 else ''
            question_number = row[1].strip() if len(row) > 1 else ''

            # Convert letter-based correct answer to number
            correct_answer = convert_answer_letter_to_number(correct_answer)

            # Generate Source ID (e.g., "PREP-AL-1.1")
            source_id = f"{source_prefix}-{question_number}"

            # Create intermediate format row
            intermediate_row = {
                'Question': question_stem,
                'Choice 1': answer_a,
                'Choice 2': answer_b,
                'Choice 3': answer_c,
                'Choice 4': answer_d,
                'Correct Answer': correct_answer,
                'Source': source_id,
                'Explanation': explanation
            }

            questions.append(intermediate_row)

    return questions

def run_xlsx_converter(intermediate_csv, output_xlsx):
    """
    Run the CSV to XLSX converter as subprocess

    What: Executes csv_to_xlsx_converter.py to generate final XLSX
    Why: Separates concerns - transformation vs formatting
    How: Uses subprocess to call converter with proper arguments
    """
    try:
        result = subprocess.run(
            ['python3', 'csv_to_xlsx_converter.py', intermediate_csv, output_xlsx],
            check=True,
            capture_output=True,
            text=True
        )
        print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Error running converter:")
        print(e.stderr)
        return False

def main():
    """
    Main transformation and conversion pipeline

    What: Orchestrates the two-step transformation process
    Why: Provides end-to-end conversion from source CSV to upload-ready XLSX
    How: Transforms to intermediate CSV, then converts to final XLSX format
    """

    print("=" * 60)
    print("Question CSV Transformation Pipeline")
    print("=" * 60)

    # Step 1: Transform to intermediate CSV
    print(f"\n📖 Reading: {INPUT_CSV}")
    questions = transform_questions_csv()

    print(f"✅ Found: {len(questions)} questions")

    # Create DataFrame and save intermediate CSV
    df = pd.DataFrame(questions)
    intermediate_csv = f"{OUTPUT_NAME}_intermediate.csv"
    df.to_csv(intermediate_csv, index=False)

    print(f"💾 Saved intermediate: {intermediate_csv}")

    # Step 2: Convert to XLSX
    output_xlsx = f"{OUTPUT_NAME}.xlsx"
    print(f"\n🔄 Converting to XLSX...")

    success = run_xlsx_converter(intermediate_csv, output_xlsx)

    if success:
        print(f"\n✅ Pipeline complete!")
        print(f"   Total questions: {len(questions)}")
        print(f"   Intermediate CSV: {intermediate_csv}")
        print(f"   Final XLSX: {output_xlsx}")
    else:
        print(f"\n❌ Pipeline failed during XLSX conversion")
        print(f"   Intermediate CSV saved: {intermediate_csv}")
        sys.exit(1)

    print("=" * 60)

if __name__ == "__main__":
    main()
