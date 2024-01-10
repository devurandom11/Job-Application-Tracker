import openpyxl
from openpyxl import Workbook
import os
import argparse


def check_for_spreadsheet(filename):
    expected_headers = [
        "Job Title",
        "Website",
        "Date Applied",
        "Company",
        "Location",
        "Job Role",
    ]

    try:
        wb = openpyxl.load_workbook(filename)
        sheet = wb.active
        actual_headers = [cell.value for cell in sheet[1]]

        return actual_headers == expected_headers
    except Exception as e:
        print(f"Error reading file: {e}")
        return False


def save_to_excel(job_details, filename):
    if not os.path.exists(filename):
        wb = Workbook()
        sheet = wb.active
        sheet.append(
            ["Job Title", "Website", "Date Applied", "Company", "Location", "Job Role"]
        )
    else:
        wb = openpyxl.load_workbook(filename)
        sheet = wb.active

    sheet.append(job_details)
    wb.save(filename)


def main():
    parser = argparse.ArgumentParser(description="Job Application Tracker")
    parser.add_argument("-f", "--file", help="File to save job details to")
    args = parser.parse_args()

    filename = args.file
    while True:
        if os.path.exists(filename) and not check_for_spreadsheet(filename):
            print("File exists but does not contain correct headers")
            user_choice = input(
                "Would you like to create a new worksheet in the same directory? (y/N): "
            ).lower()
            if user_choice == "y":
                path, ext = os.path.splitext(filename)
                filename = f"{path}_new{ext}"
                break
            elif user_choice == "n":
                print("Exiting")
                exit()
            else:
                print("Invalid choice")

    print("Job Application Tracker")

    job_title = input("Enter job title: ")
    website = input("Enter website: ")
    date_applied = input("Enter date applied: ")
    company = input("Enter company: ")
    location = input("Enter location: ")
    job_role = input("Enter job role: ")

    job_details = [
        job_title,
        website,
        date_applied,
        company,
        location,
        job_role,
    ]
    save_to_excel(job_details, filename)

    print("Job details saved successfully")


if __name__ == "__main__":
    main()
