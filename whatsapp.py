from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError, Page, Locator
from datetime import datetime
import pandas as pd
import random
import time
import os

# --- Configuration ---
EXCEL_FILE_PATH = 'example_list.xlsx'
COMPANY_NAME_COLUMN = 'ŞİRKET ADI'
PHONE_NUMBER_COLUMN = 'TELEFON'
NAME_COLUMN = 'ADI SOYADI'
PDF_PATH_COLUMN = 'PDF_PATH'
LOG_FILE_PATH = 'logs.txt'
MESSAGE_TEMPLATE = "Sn. {name}, {company_name} ...[DESIRED MESSAGE]..."

# --- Security & Anti-Bot Configuration ---
# A longer, more random delay is much safer for large batches.
DELAY_BETWEEN_MESSAGES_BASE = 30  # seconds
DELAY_BETWEEN_MESSAGES_RANDOM = 25 # seconds (final delay will be 30 + 0-25s)

# --- Selectors (English with Turkish comments for reference) ---
SELECTORS = {
    "search_box": 'div[aria-label="Search input textbox"]',  # Turkish: 'div[aria-label="Arama metni giriş alanı"]'
    "first_contact_result": "//*[@id='pane-side']/div/div/div/div[2]/div",
    "attach_button": 'button[title="Attach"]',               # Turkish: 'button[title="Ekle"]'
    "document_menu_item": 'text="Document"',                 # Turkish: 'text="Belge"'
    "caption_box": 'div[aria-label="Add a caption"]',        # Turkish: 'div[aria-label="Başlık ekleyin"]'
    "send_button": 'div[aria-label="Send"]',                 # Turkish: 'div[aria-label="Gönder"]'
}

# --- Advanced Anti-Bot Detection Settings ---
# These arguments help the browser appear less like an automated tool.
BROWSER_ARGS = [
    '--no-sandbox', '--disable-setuid-sandbox', '--disable-infobars',
    '--disable-dev-shm-usage', '--disable-blink-features=AutomationControlled',
    '--disable-gpu', '--no-zygote', '--window-size=1280,1600', '--lang=en-US',
]

# These options set a realistic browser environment.
CONTEXT_OPTIONS = {
    'ignore_https_errors': True,
    'user_agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/115.0',
    'extra_http_headers': { 'Accept-Language': 'en-US,en;q=0.9' },
    'locale': 'en-US',
    'timezone_id': 'Europe/Istanbul',
}

class WhatsAppBot:
    """
    Final refactored class for WhatsApp automation with enhanced security and stability.
    """
    def __init__(self, config):
        self.config = config
        self.page: Page = None
        self.context = None
        self.playwright = None

    def _log_success(self, name, company, phone):
        """Appends a record of a successful send to the log file."""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"{timestamp} - SUCCESS - Name: {name}, Company: {company}, Phone: {phone}\n"
        try:
            with open(self.config['log_file'], 'a', encoding='utf-8') as f:
                f.write(log_entry)
        except Exception as e:
            print(f"Could not write to log file: {e}")

    def _patient_click(self, locator: Locator, timeout=15000):
        """A robust click function with human-like delays."""
        locator.wait_for(state='visible', timeout=timeout)
        time.sleep(random.uniform(0.7, 1.5))
        locator.click()
        time.sleep(random.uniform(0.4, 0.8))

    def _load_data(self):
        """Loads and validates data from the Excel file."""
        print(f"Reading data from '{self.config['excel_path']}'...")
        try:
            df = pd.read_excel(
                self.config['excel_path'], 
                dtype={
                    self.config['phone_col']: str, 
                    self.config['name_col']: str, 
                    self.config['company_col']: str
                }
            )
            required_cols = [self.config['phone_col'], self.config['pdf_col'], self.config['name_col'], self.config['company_col']]
            df.dropna(subset=required_cols, inplace=True)
            df = df[df[self.config['phone_col']].str.strip() != '']
            df = df[df[self.config['pdf_col']].str.strip() != '']
            if df.empty:
                print("No valid rows found. Please check your Excel file.")
                return None
            print(f"Found {len(df)} rows to process.")
            return df
        except FileNotFoundError:
            print(f"Error: The file '{self.config['excel_path']}' was not found.")
            return None
        except Exception as e:
            print(f"Failed to load Excel file: {e}")
            return None

    def _launch_browser(self):
        """Launches the browser with anti-bot settings."""
        self.playwright = sync_playwright().start()
        self.context = self.playwright.firefox.launch_persistent_context(
            user_data_dir="firefox_user_data",
            headless=False,
            slow_mo=200,
            args=BROWSER_ARGS,
            **CONTEXT_OPTIONS
        )
        self.page = self.context.pages[0]

    def _login(self):
        """Navigates to WhatsApp and waits for a reliable login indicator."""
        print("\nOpening WhatsApp Web...")
        self.page.goto("https://web.whatsapp.com/")
        print("Please scan the QR code if you are not logged in.")
        print("The script will wait for you to log in...")
        try:
            # Using the search box as the login indicator, as it's the first key interactive element.
            self.page.locator(SELECTORS["search_box"]).wait_for(timeout=120000)
            print("Login successful! Starting automation.")
            return True
        except PlaywrightTimeoutError:
            print("Login timeout. Please try again.")
            return False

    def _reset_ui(self):
        """Resets the UI by clearing the search box."""
        print("Attempting to reset UI by clearing the search box...")
        try:
            search_box = self.page.locator(SELECTORS["search_box"])
            search_box.fill("")
            time.sleep(1)
        except Exception as e:
            print(f"Could not reset UI state gracefully: {e}")

    def run(self):
        """The main execution method for the bot."""
        data = self._load_data()
        if data is None: return
        self._launch_browser()

        try:
            if not self._login(): return

            for index, row in data.iterrows():
                phone_number = str(row[self.config['phone_col']]).strip()
                pdf_path = str(row[self.config['pdf_col']]).strip()
                name = str(row[self.config['name_col']]).strip()
                company_name = str(row[self.config['company_col']]).strip()
                
                print(f"\n--- Processing Row {index + 1}/{len(data)}: {name} ({company_name}) ---")

                if not os.path.exists(pdf_path):
                    print(f"PDF not found: {pdf_path}. Skipping.")
                    continue

                try:
                    # Step 1: Search and select the contact
                    print(f"Searching for {phone_number}...")
                    search_box = self.page.locator(SELECTORS["search_box"])
                    self._patient_click(search_box)
                    search_box.fill(phone_number)
                    
                    print("Selecting contact...")
                    contact_result = self.page.locator(SELECTORS["first_contact_result"]).first
                    self._patient_click(contact_result)
                    print("Contact found and selected.")

                    # Step 2: Attach the document
                    print("Opening attachment menu...")
                    attach_btn = self.page.locator(SELECTORS["attach_button"])
                    self._patient_click(attach_btn)
                    print("Attach button clicked.")

                    # Step 3: Choose the "Document" option and upload the file
                    print("Choosing document...")
                    with self.page.expect_file_chooser() as fc_info:
                        document_btn = self.page.locator(SELECTORS["document_menu_item"])
                        self._patient_click(document_btn, timeout=10000)
                    
                    file_chooser = fc_info.value
                    file_chooser.set_files(pdf_path)

                    # Step 4: Add caption and send
                    send_btn = self.page.locator(SELECTORS["send_button"])
                    
                    personalized_message = self.config['message_template'].format(name=name, company_name=company_name)
                    
                    print(f"Adding caption: '{personalized_message}'")
                    caption_box = self.page.locator(SELECTORS["caption_box"])
                    self._patient_click(caption_box)
                    caption_box.fill(personalized_message)

                    self._patient_click(send_btn)
                    print(f"✅ Sent PDF to {name} ({phone_number})")
                    
                    # Log the successful send to the text file
                    self._log_success(name, company_name, phone_number)

                    # Step 5: Wait between messages with a safer, randomized delay
                    wait_time = self.config['delay_base'] + random.uniform(0, self.config['delay_random'])
                    print(f"Waiting {wait_time:.2f} seconds before next message...")
                    time.sleep(wait_time)

                except Exception as e:
                    print(f"❌ Error processing {phone_number}: {e}")
                    self._reset_ui()
                    continue

            print("\n✅ All messages processed.")

        finally:
            print("Closing browser context...")
            if self.context: self.context.close()
            if self.playwright: self.playwright.stop()

if __name__ == "__main__":
    bot_config = {
        "excel_path": EXCEL_FILE_PATH,
        "company_col": COMPANY_NAME_COLUMN,
        "phone_col": PHONE_NUMBER_COLUMN,
        "name_col": NAME_COLUMN,
        "pdf_col": PDF_PATH_COLUMN,
        "log_file": LOG_FILE_PATH,
        "message_template": MESSAGE_TEMPLATE,
        "delay_base": DELAY_BETWEEN_MESSAGES_BASE,
        "delay_random": DELAY_BETWEEN_MESSAGES_RANDOM,
    }
    bot = WhatsAppBot(bot_config)
    bot.run()
