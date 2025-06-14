# WhatsApp Personalized PDF Sender

This Python script automates the process of sending personalized messages with unique PDF attachments to a list of contacts via WhatsApp Web. It reads contact information, PDF file paths, and personalization data from an Excel file, and then uses Playwright to control a browser, mimicking human behavior to reduce the risk of detection.

This tool is ideal for businesses and organizations that need to send customized documents, such as invoices, reports, or certificates, to a large number of recipients efficiently and reliably.

## Key Features

* **Excel Integration:** Reads contact details (Phone Number, Name, Company Name) and corresponding PDF file paths from an `.xlsx` file.
* **Personalized Messaging:** Uses a message template to dynamically insert the recipient's name and company name into each message.
* **Automated PDF Attachment:** Searches for, selects, and attaches the correct PDF file for each contact.
* **Human-like Automation:** Utilizes Playwright with persistent browser sessions and robust, "patient" interaction logic (randomized delays, stable selectors) to mimic human behavior.
* **Advanced Anti-Bot Measures:** Includes browser arguments and context options designed to make the automation harder to detect.
* **Success Logging:** Creates a `sending_log.txt` file to keep a timestamped record of every successfully sent message, making it easy to track progress and resume if interrupted.
* **Robust Error Handling:** If an error occurs with one contact, the script logs the error, resets the UI gracefully, and continues to the next person on the list.
* **Language Support:** Selectors are configured for the English version of WhatsApp Web, with comments provided for the Turkish versions.

## How It Works

The script follows a logical, step-by-step process for each contact in the Excel file:

1.  **Launch Browser:** Opens a Firefox browser with a persistent context to save your WhatsApp login session.
2.  **Manual Login:** The script waits for you to manually log into WhatsApp Web by scanning the QR code.
3.  **Loop Through Contacts:** It iterates through each row of the `final_list.xlsx` file.
4.  **Search & Select:** It types the contact's phone number into the WhatsApp search bar and clicks the first result.
5.  **Attach Document:** It clicks the "Attach" button, selects the "Document" option, and programmatically uploads the correct PDF.
6.  **Add Caption & Send:** It types the personalized message into the caption box and clicks the "Send" button.
7.  **Log Success:** Upon a successful send, it writes a timestamped entry to `sending_log.txt`.
8.  **Wait:** It pauses for a safe, randomized interval before processing the next contact.
9.  **Reset on Error:** If any step fails, it logs the error, resets the UI, and moves on to the next contact.

## Configuration

All user-configurable parameters are located at the top of the `whatsapp.py` script:

* `EXCEL_FILE_PATH`
* `PHONE_NUMBER_COLUMN`
* `PDF_PATH_COLUMN`
* `NAME_COLUMN`
* `COMPANY_NAME_COLUMN`
* `MESSAGE_TEMPLATE`
* `DELAY_BETWEEN_MESSAGES_BASE`
* `DELAY_BETWEEN_MESSAGES_RANDOM`

## Prerequisites

Before running the script, you need to have Python and the following libraries installed.

```bash
pip install pandas
pip install openpyxl
pip install playwright
```

After installing the libraries, you must install the necessary browser files for Playwright:
```bash
playwright install
```

## How to Run

1.  Place your Excel file (e.g., `final_list.xlsx`) and the folder containing all your PDFs in the same directory as the script.
2.  Ensure the column names and file paths in the **Configuration** section of the script match your setup.
3.  Open a terminal or command prompt in the project directory.
4.  Run the script using the following command:
    ```bash
    python whatsapp.py
    ```
5.  The first time you run it, a Firefox window will open. Scan the WhatsApp Web QR code with your phone to log in. The script will save your session for future runs.

