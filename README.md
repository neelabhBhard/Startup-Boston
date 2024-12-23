# Startup Boston Event Data Analysis
This Python script analyzes attendee data from Startup Boston Week 2024, processing information about participants, sessions, and various demographic details.

## Features

- Removes whitespace from string columns
- Creates floor number assignments
- Calculates unique attendee counts per session and floor
- Analyzes attendee distribution by:
  - Company size
  - Job function
  - Geographic location
  - Funding stage
  - Current residence

## Key Functions

### get_unique_cols(col_name)
Extracts unique values from a column, removing empty strings and NaN values.

### get_freq_of_questions(col_name)
Analyzes frequency of responses to questions like "How did you hear about this event?"

### group_count_2d_col(cols, col_to_process)
Processes two-dimensional column data and returns frequency counts.

## Output

Generates an Excel file with multiple sheets containing:
- Attendance summary by session
- Attendee breakdown
- Marketing channel effectiveness
- Attendance motivation analysis
- Demographics distribution

## Requirements

- pandas
- openpyxl

The analysis results will be saved to 'bostonStartUp_attendance.xlsx' with multiple worksheets.
