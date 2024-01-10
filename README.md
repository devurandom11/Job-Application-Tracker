# Job Application Tracker

## Description

Job Application Tracker is a Python CLI tool designed to help me track my job applications. It allows users to input job application details such as job title, website, date applied, company name, location, and job role, and saves this information in an Excel spreadsheet. Additionally, it calculates the days since each application was submitted and visually indicates the urgency for follow-up actions. I'm sure there are many other tools out there that do the same thing, but I wanted to create my own to practice my Python skills.

## Features

- Input validation for website URLs and date formats.
- Ability to specify a custom Excel file for storing job application details.
- Automatically creates a new Excel file if none exists or if the existing file doesn't match the expected format.
- Safe handling of Excel files to prevent data loss.
- Automatically calculates the days since each job application was submitted.
- Color-coded indication in the Excel file for follow-up actions based on the time elapsed since application submission.

## Requirements

- Python 3.x
- `openpyxl` module for handling Excel files.
- `dateutil` module for date parsing.

## Installation

1. Clone this repository:

    ```bash
    git clone https://github.com/devurandom11/Job-Application-Tracker.git
    ```

2. Navigate to the cloned directory.

3. Install the required modules:

    ```bash
    pip install openpyxl python-dateutil
    ```

## Usage

1. Run the script:

    ```bash
    python job_application_tracker.py
    ```

2. Optionally, specify a custom Excel file to store the job details:

    ```bash
    python job_application_tracker.py -f path/to/yourfile.xlsx
    ```

3. Follow the on-screen prompts to enter your job application details.

4. The entered details will be saved in the specified Excel file. If the file does not exist, it will be created. If the file exists but does not match the expected format, you will be prompted to create a new file.

## File Format

The Excel file used to store job application details contains the following columns:

- Job Title
- Website
- Date Applied
- Company Name
- Location
- Job Role
- Date Since Submission (automatically calculated)
- Followed Up? (for manual user input to indicate whether follow-up actions have been taken)

## Additional Functionality

- The "Date Since Submission" column uses an Excel formula to calculate the number of days since the application date.
- This column is color-coded: Red if less than 7 days, Yellow if between 7 and 14 days, and Green if 14 days or more have passed since the application date.
- The "Followed Up?" column is provided for users to manually mark whether they have followed up on the application.
