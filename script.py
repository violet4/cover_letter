import os
import csv
import pandas as pd
from jinja2 import Environment, FileSystemLoader
import argparse
import datetime

def generate_cover_letter(data, template):
    cover_letter = template.render(**data)
    return cover_letter

def validate_data(data):
    # Add data validation logic here, if needed
    required_fields = ['name', 'job_title', 'company_name']
    for field in required_fields:
        if field not in data:
            raise ValueError(f"Missing required field: '{field}'")

def process_row(row, template, output_folder):
    try:
        validate_data(row)
        # Convert all keys to lowercase in the data dictionary
        lowercase_data = {key.lower(): value for key, value in row.items()}
        cover_letter = generate_cover_letter(lowercase_data, template)

        output_file = os.path.join(output_folder, f'{row["name"]}_cover_letter.txt')
        with open(output_file, 'w') as file:
            file.write(cover_letter)

    except ValueError as e:
        print(f"Data validation error for row: {row}. {e}")

def create_output_folder():
    # Create the output folder based on today's date and a unique identifier
    today_date = datetime.date.today().strftime('%Y-%m-%d')
    unique_identifier = 1

    while True:
        default_output_folder = f'output_{today_date}_{unique_identifier}'
        if not os.path.exists(default_output_folder):
            os.makedirs(default_output_folder)
            break
        unique_identifier += 1

    return default_output_folder

def main(input_file, output_folder, template_file):
    # If no output folder is provided, create the default output folder
    if output_folder is None:
        output_folder = create_output_folder()
    else:
        # Create the output folder if it doesn't exist
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

    # Load the template only once
    env = Environment(loader=FileSystemLoader('templates'))
    template = env.get_template(template_file)

    # Check the file extension and read data accordingly
    if input_file.lower().endswith('.csv'):
        with open(input_file, 'r', newline='') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                process_row(row, template, output_folder)

    elif input_file.lower().endswith(('.xls', '.xlsx')):
        df = pd.read_excel(input_file)
        for _, row in df.iterrows():
            process_row(row.to_dict(), template, output_folder)

    else:
        print("Unsupported file format. Please provide a CSV or Excel file.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate personalized cover letters from a CSV or Excel file.")
    parser.add_argument("input_file", help="Path to the input CSV or Excel file.")
    parser.add_argument("--output_folder", help="Path to the output folder where the cover letters will be saved.")
    parser.add_argument("--template_file", default="cover_letter_template.txt", help="Path to the template file for the cover letter content.")
    args = parser.parse_args()

    main(args.input_file, args.output_folder, args.template_file)
