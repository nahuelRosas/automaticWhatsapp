### WhatsApp Reminder Messaging Automation

#### Overview
This project automates sending WhatsApp reminders for vehicle insurance expiration using Selenium and Chrome WebDriver. It reads recipient data from an Excel sheet, generates personalized messages based on a template, and sends messages through WhatsApp Web.

#### Requirements
- Python 3.7+
- Chrome Browser
- Chromedriver
- Libraries:
  - pandas
  - requests
  - selenium
  - colorama
  - tqdm

#### Installation Instructions
1. **Clone the Repository**
   ```bash
   git clone <https://github.com/nahuelRosas/automaticWhatsapp>
   cd automaticWhatsapp
   ```

2. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Setup Chromedriver**
   - Chromedriver is automatically set up using `chromedriver_autoinstaller` within the script.

#### Usage
1. **Prepare Data**
   - Place recipient details in `data.xlsx`. If the file doesn't exist, a template is created automatically.

2. **Run the Script**
   ```bash
   python whatsapp_reminder.py
   ```

3. **Script Options**
   - **View Data:** Displays data loaded from `data.xlsx`.
   - **Verify Messages:** Checks pending messages to be sent.
   - **Send Messages:** Sends pending WhatsApp reminders.
   - **Preview Messages:** Displays a preview of WhatsApp messages.

#### Detailed Functionality
- **`setup_chromedriver()`**: Downloads and installs Chrome and Chromedriver if necessary.
- **`load_excel_data(file_path)`**: Loads data from `data.xlsx` into a pandas DataFrame.
- **`send_whatsapp_message(...)`**: Sends a personalized WhatsApp message to recipients.
- **`preview_whatsapp_message(df)`**: Previews WhatsApp messages before sending.
- **Error Handling**: Handles errors gracefully with colored output using `colorama`.

#### Notes
- Ensure Chrome and Chromedriver are compatible with your operating system.
- Customize the message template (`template`) in the script as needed.
- Review Chrome options (`chrome_options`) for WebDriver customization.

#### Contributors
- Nahuel Rosas ([GitHub](https://github.com/nahuelRosas))

#### License
This project is licensed under the [MIT License](LICENSE).
