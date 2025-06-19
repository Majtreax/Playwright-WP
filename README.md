# WhatsApp Automation Suite
* This project contains a suite of Python scripts designed to automate sending personalized WhatsApp messages with PDF attachments. 
* It includes a primary bot for sending messages and a utility script to first convert an Excel list into a VCF (Virtual Contact File) for easy import into your phone.
## Project Components
1.  **`whatsapp.py` (The Main Bot)**
    * Automates sending personalized messages with unique PDF attachments.
    * Reads contact details, debt information, and PDF paths from an Excel file.
    * Uses a safe "Act then Verify" logic to ensure messages are sent to the correct person.
    * Creates timestamped log files for each run.
2.  **`vcf.py` (VCF Creator Utility)**
    * Reads an Excel file containing names, company names, and phone numbers.
    * Generates a standard `contacts.vcf` file.
    * This VCF file can be easily imported into your Android or iOS phone.
## Project Structure
`Your project folder should be set up like this.`
* |-- requirements.txt
* |-- whatsapp.py
* |-- list.xlsx
* |-- /logs/               (created automatically by `whatsapp.py`)
* |-- /firefox_user_data/  (created automatically after first login)
## Setup and Installation
`Follow these steps to get the bot running on a new machine.`
1. Python
* Ensure you have Python 3.8 or newer installed. You can download it from [python.org](https://www.python.org/).
2. Install Libraries
* Open your terminal or command prompt in the project folder and install all the necessary Python libraries at once using the `requirements.txt` file:
```bash
pip install -r requirements.txt
```
3. Install Playwright Browsers
* After installing the libraries, you must download the necessary browser files for Playwright. 
```bash
playwright install
```
## How to Use: The Complete Workflow
`Follow these steps in order for a smooth and reliable process.`
* Prepare the main `list.xlsx` Excel file.
* Your primary Excel file must contain all the information for the `whatsapp.py` bot.
* `ADI SOYADI`
* `ŞİRKET ADI`
* `TELEFON NO`
* `TOPLAM BORÇ`
* `PDF_PATH`
* Important: The `PDF_PATH` for each row must be the complete and correct file path to the PDF on your computer 
(e.g., C:\Users\YourUser\Documents\invoices\123.pdf).
## [Optional] Create the VCF Contact File
`This step prepares your contacts for import into your phone.`
* Ensure the `phone_list.xlsx` contains required columns for the `vcf.py` script.
* `ADI SOYADI`
* `ŞİRKET ADI`
* `TELEFON NO`
* Run the script:
```bash
python vcf.py
```
* Import Contacts to Your Phone
* Transfer the `contacts.vcf` file to your phone and import it into your phone's Contacts or Address Book application. This will save all your recipients as contacts.
## First-Time Login for WhatsApp Bot
`You only need to do this once.`
* **In whatsapp.py, change the setting `HEADLESS_MODE = True` to `HEADLESS_MODE = False`.**
* **Run the script. A Firefox browser window will open and navigate to WhatsApp Web.**
* **Scan the QR code using the WhatsApp app on your phone.**
* **Wait for your chats to load completely. You can then stop the script.**
* **Change `HEADLESS_MODE` back to `True` for all future runs.**
## Run the Main Bot to Send Messages
* With everything configured and your contacts imported, simply run the main script:
```bash
python whatsapp.py
```
* The bot will start processing the rows from your `list.xlsx` file.
## Disclaimer
* This script is intended for personal and legitimate use cases. 
* Automating interactions with web services may be against their terms of service. Please use this script responsibly. 
* The user interface of WhatsApp Web can change at any time, which may require updating the selectors in the script.
