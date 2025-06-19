# --- VCF (Virtual Contact File) Creator ---
#
# Description:
# This script reads contact information from an Excel file and converts it into a
# universal .vcf file. This VCF file can then be easily imported into any
# standard contacts application on Android or iOS.
#
# The main purpose of this utility is to prepare contacts for the main whatsapp.py bot.

import pandas as pd
import re
import sys

# --- 1. CONFIGURATION ---
# All user-configurable settings are grouped here for easy access.

# --- File and Column Settings ---
# These names must exactly match your Excel file and its column headers.
INPUT_EXCEL_FILE = "phone_list.xlsx"
OUTPUT_VCF_FILE = "contacts.vcf"

NAME_COLUMN = 'ADI SOYADI'
COMPANY_COLUMN = 'ŞİRKET ADI'
PHONE_COLUMN = 'TELEFON NO'


def read_contacts(file_path: str):
    """
    Reads contact data from the specified Excel file into a pandas DataFrame.

    Args:
        file_path: The path to the input Excel file.

    Returns:
        A pandas DataFrame containing the contact data, or None if an error occurs.
    """
    print(f"Reading data from '{file_path}'...")
    try:
        df = pd.read_excel(file_path, engine="openpyxl")
        # Replace all empty (NaN) cells with an empty string to prevent errors.
        df = df.fillna("")
        print(f"Successfully found {len(df)} rows to process.")
        return df
    except FileNotFoundError:
        print(f"Error: The input file '{file_path}' was not found. Please check the file name and location.")
        return None
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return None


def generate_vcf_card(full_name: str, company: str, phone: str) -> str:
    """
    Generates the text for a single contact in the VCF 3.0 format.

    Args:
        full_name: The contact's full name.
        company: The contact's company name.
        phone: A cleaned phone number (digits only).

    Returns:
        A string representing a single VCF card.
    """
    # Create a display name, including the company in parentheses if it exists.
    # This is what will be shown in the contact list.
    display_name = f"{full_name} ({company})" if company else full_name

    # VCF cards are structured plain text. Each card begins with BEGIN:VCARD
    # and ends with END:VCARD.
    vcf_card_lines = [
        "BEGIN:VCARD",
        "VERSION:3.0",
        # FN (Formatted Name): The primary display name for the contact.
        f"FN:{display_name}",
        # N (Structured Name): Provides components of the name. We place the
        # full display name in the 'Given Name' part for maximum compatibility.
        f"N:;{display_name};;;",
        # TEL (Telephone): The contact's phone number.
        f"TEL;TYPE=CELL:{phone}",
        "END:VCARD"
    ]
    return "\n".join(vcf_card_lines) + "\n"


def process_contacts(contacts_df: pd.DataFrame, output_file: str):
    """
    Processes the DataFrame, generates VCF cards for unique contacts,
    and writes them to the output file.
    """
    # Using a set is a highly efficient way to track phone numbers we've
    # already processed, ensuring no duplicate contacts are created.
    processed_phones = set()

    print(f"Writing contacts to '{output_file}'...")
    with open(output_file, "w", encoding="utf-8") as f:
        # Iterate over each row in the DataFrame.
        for _, row in contacts_df.iterrows():
            # Extract data, ensuring it's a string and stripping whitespace.
            full_name = str(row.get(NAME_COLUMN, "")).strip()
            company = str(row.get(COMPANY_COLUMN, "")).strip()
            phone = str(row.get(PHONE_COLUMN, "")).strip()

            # Skip rows that are missing a name and a phone number.
            if not full_name and not phone:
                continue

            # Clean the phone number by removing all non-digit characters.
            cleaned_phone = re.sub(r'\D', '', phone)

            # Skip if the phone number is empty after cleaning or if it's a duplicate.
            if not cleaned_phone or cleaned_phone in processed_phones:
                continue

            # Generate the VCF card text and write it to the file.
            vcf_entry = generate_vcf_card(full_name, company, cleaned_phone)
            f.write(vcf_entry)

            # Add the processed number to our set to prevent duplicates.
            processed_phones.add(cleaned_phone)

    print("\nProcessing complete.")
    print(f"A total of {len(processed_phones)} unique contacts have been written to '{output_file}'.")


def main():
    """
    Main function to orchestrate the VCF creation process.
    """
    contacts_df = read_contacts(INPUT_EXCEL_FILE)

    # Only proceed if the data was loaded successfully.
    if contacts_df is not None:
        process_contacts(contacts_df, OUTPUT_VCF_FILE)
    else:
        print("Halting execution due to an error reading the Excel file.")
        sys.exit(1) # Exit with an error code.


# --- 3. SCRIPT EXECUTION ---
# This standard Python block ensures that the main() function is called
# only when the script is executed directly.
if __name__ == "__main__":
    main()
