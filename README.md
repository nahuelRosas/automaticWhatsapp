# WhatsApp Message Automation with Selenium and Python

## Introduction
This Python script automates the sending of WhatsApp messages using Selenium WebDriver. It interacts with WhatsApp Web to send reminders about insurance expiration dates to recipients stored in an Excel file.

## Dependencies
- **Python Libraries:**
  - `re`: Regular expression operations.
  - `pandas`: Data manipulation and analysis.
  - `datetime`: Date and time handling.
  - `selenium`: Browser automation.
  - `chromedriver_autoinstaller`: Automatic installation of ChromeDriver.
  - `time`, `os`: Standard Python libraries for time delays and operating system operations.
  - `colorama`: Terminal text colorization.

## Setup and Initialization
The script begins by importing necessary libraries and initializing Colorama for terminal output colorization. It defines a message template dynamically and sets up functions for Chromedriver installation and WhatsApp session login status.

## Functions Overview
1. **`setup_chromedriver()`**: Checks and installs the latest Chromedriver.
2. **`is_logged_in(driver)`**: Checks if WhatsApp session is logged in.
3. **`clean_to_bmp(text)`**: Cleans text to include only Basic Multilingual Plane (BMP) characters.
4. **`send_whatsapp_message(driver, recipient, df, file_path, chat_type='contact')`**: Sends WhatsApp messages based on data from Excel, handling group or individual chats.
5. **`load_excel_data(file_path, sheet_name='Sheet1')`**: Loads data from an Excel file into a Pandas DataFrame.
6. **`verbose_excel_data(file_path, sheet_name='Sheet1')`**: Displays loaded Excel data.
7. **`preview_whatsapp_message(df)`**: Previews messages to be sent based on upcoming expiration dates.
8. **`display_pending_messages(df, today)`**: Displays pending messages to be sent today.

## Main Execution Flow
- **Initialization**: Installs Chromedriver, loads Excel data, and sets up a WebDriver session with WhatsApp Web.
- **Menu Interaction**: Allows the user to view data, verify pending messages, send messages, preview messages, or exit.
- **Message Sending**: Checks expiration dates and sends reminders to WhatsApp recipients based on predefined templates.

## Error Handling
- The script includes error handling for various scenarios such as missing keys in Excel data, element not found exceptions, and general errors during message sending.

## Conclusion
This script demonstrates automated messaging using Selenium and Python, facilitating efficient communication of insurance expiration reminders via WhatsApp.
