# Study Schedule Manager

This Python project helps you manage your study topics and track your review schedule using a simple CLI and an automated Excel output.

## Features

- **CLI-based Topic Management:** Add new study topics and update their review status.
- **Automated Scheduling:** Calculates review dates based on a set of ranges (e.g., 0, 7, 20 days).
- **Excel Export:** Automatically generates and updates an Excel spreadsheet (`study_schedule.xlsx`) on Google Drive with color-coded review statuses.
- **Summary Sheet:** Provides an overview of overdue, today's, tomorrow's, and the day after tomorrow's study tasks.
- **Local Data Storage:** Uses local CSV files to store subject data (these are excluded from git to maintain data privacy).

## Project Structure

- `data_input.py`: The main script for user interaction and data processing.
- `*.csv`: Subject-specific data files (automatically created as needed).
- `study_schedule.xlsx`: The final output spreadsheet.

## Requirements

- Python 3.x
- `openpyxl` library

## Setup

1. **Clone the repository:**
   ```bash
   git clone <your-repo-url>
   cd study_schedule
   ```

2. **Set up a virtual environment (optional but recommended):**
   ```bash
   python -m venv venv
   .\venv\Scripts\activate  # Windows
   source venv/bin/activate # Linux/macOS
   ```

3. **Install dependencies:**
   ```bash
   pip install openpyxl
   ```

4. **Configuration:**
   Update the `FILE_PATH` in `data_input.py` to point to your desired Excel file location (e.g., your Google Drive folder).

## Usage

Run the main script:
```bash
python data_input.py
```
Follow the on-screen prompts to select a subject, add topics, or update reviews.
