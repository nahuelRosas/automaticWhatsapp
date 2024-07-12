import platform
import re
import subprocess
import sys
import pandas as pd
from datetime import datetime, timedelta
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import chromedriver_autoinstaller
import time
import os
from colorama import init, Fore
from tqdm import tqdm
from selenium.webdriver.remote.webdriver import WebDriver

# Initialize colorama
init(autoreset=True)

# Dynamically generate template and message
template: str = (
    "_¡Hola {name}!_\n\n"
    "_*Recordatorio Importante*_\n\n"
    "_Queremos recordarte que tu seguro para el vehículo {model_vehicle} con patente {vehicle_license} vence el {expire_date:%d/%m/%Y}._\n\n"
    "_Desde ya, muchas gracias por confiar en nosotros._\n\n"
    "_No olvides que puedes pagar de manera sencilla y segura *por este medio.*_\n\n"
    "_Si ya realizaste el pago, *ignora este mensaje*._\n"
    "_Para dejar de recibir estos recordatorios, simplemente envíanos un mensaje con la palabra *cancelar*._"
)


def setup_chromedriver() -> None:
    """
    Sets up the Chromedriver by downloading and installing Chrome if necessary,
    and then installing the Chromedriver.

    Raises:
        Exception: If there is an error setting up Chromedriver.

    Returns:
        None
    """
    try:
        download_and_install_chrome()
        chromedriver_autoinstaller.install()
        print(f"{Fore.GREEN}Chromedriver is up to date.")
    except Exception as e:
        print(f"{Fore.RED}Error setting up Chromedriver: {str(e)}")
        exit()


def is_chrome_installed() -> bool:
    """
    Checks if Google Chrome is installed on the system.

    Returns:
        bool: True if Chrome is installed, False otherwise.
    """
    system: str = platform.system()

    if system == "Windows":
        chrome_path: str = os.path.join(
            os.getenv("PROGRAMFILES", "C:\\Program Files"),
            "Google",
            "Chrome",
            "Application",
            "chrome.exe",
        )
        return os.path.exists(chrome_path)
    elif system == "Linux":
        try:
            result: subprocess.CompletedProcess = subprocess.run(
                ["which", "google-chrome"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            return result.returncode == 0
        except Exception:
            return False
    else:
        return False


def download_file_with_progress(url: str, dest_path: str) -> None:
    """
    Downloads a file from the given URL and saves it to the specified destination path.

    Args:
        url (str): The URL of the file to download.
        dest_path (str): The path where the downloaded file will be saved.

    Raises:
        Exception: If the downloading process fails.
    """
    response = requests.get(url, stream=True)
    total_size: int = int(response.headers.get("content-length", 0))
    block_size: int = 1024  # 1 KB
    progress_bar: tqdm = tqdm(
        total=total_size, unit="iB", unit_scale=True, ncols=80, desc="Downloading"
    )

    with open(dest_path, "wb") as file:
        for data in response.iter_content(block_size):
            progress_bar.update(len(data))
            file.write(data)

    progress_bar.close()
    if total_size != 0 and progress_bar.n != total_size:
        raise Exception("Error: Downloading failed.")


def simulate_installation_progress(description: str, steps: int = 10) -> None:
    """
    Simulates the progress of an installation process.

    Args:
        description (str): The description of the installation process.
        steps (int, optional): The number of steps in the installation process. Defaults to 10.
    """
    progress_bar: tqdm = tqdm(total=steps, ncols=80, desc=description)
    for _ in range(steps):
        subprocess.run(["sleep", "1"], check=True)
        progress_bar.update(1)
    progress_bar.close()


def download_and_install_chrome() -> None:
    """
    Downloads and installs Google Chrome based on the operating system.

    Raises:
        Exception: If the operating system is not supported.

    Returns:
        None
    """
    system: str = platform.system()

    try:
        if is_chrome_installed():
            print(Fore.GREEN + "Google Chrome is already installed and usable.")
            return

        if system == "Windows":
            print(Fore.CYAN + "Detected Windows OS.")
            chrome_url: str = (
                "https://dl.google.com/chrome/install/latest/chrome_installer.exe"
            )
            installer_path: str = os.path.join(os.getcwd(), "chrome_installer.exe")
            print(Fore.CYAN + "Downloading Chrome installer for Windows...")
            download_file_with_progress(chrome_url, installer_path)

            print(Fore.CYAN + "Installing Chrome on Windows...")
            subprocess.run([installer_path], check=True)
            os.remove(installer_path)
            print(Fore.GREEN + "Chrome installed successfully on Windows.")

        elif system == "Linux":
            print(Fore.CYAN + "Detected Linux OS.")
            with open("/etc/os-release") as f:
                os_release_info: str = f.read()

            if "Fedora" in os_release_info or "Nobara" in os_release_info:
                print(Fore.CYAN + "Detected Fedora or Nobara Linux.")
                chrome_url: str = (
                    "https://dl.google.com/linux/direct/google-chrome-stable_current_x86_64.rpm"
                )
                installer_path: str = os.path.join(
                    os.getcwd(), "google-chrome-stable_current_x86_64.rpm"
                )
                print(Fore.CYAN + "Downloading Chrome installer for Fedora/Nobara...")
                download_file_with_progress(chrome_url, installer_path)

                print(Fore.CYAN + "Installing Chrome on Fedora/Nobara...")
                simulate_installation_progress("Installing Chrome")
                subprocess.run(
                    ["sudo", "dnf", "install", "-y", installer_path], check=True
                )
                os.remove(installer_path)
                print(Fore.GREEN + "Chrome installed successfully on Fedora/Nobara.")
            else:
                raise Exception(
                    "This script currently supports only Fedora and Nobara Linux."
                )
        else:
            raise Exception(
                "This script currently supports only Windows and Fedora/Nobara Linux."
            )
    except Exception as e:
        print(Fore.RED + f"Error: {str(e)}")


def is_logged_in(driver: WebDriver) -> bool:
    """
    Checks if the WhatsApp session is logged in.

    Args:
        driver (webdriver): The WebDriver instance.

    Returns:
        bool: True if the session is logged in, False otherwise.
    """
    max_attempts: int = 24  # 2 minutes (30 seconds x 24 attempts)
    attempt: int = 1

    while attempt <= max_attempts:
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
                )
            )
            return True
        except TimeoutException:
            print(
                f"{Fore.YELLOW}Attempt {attempt}: Timeout - WhatsApp session not logged in yet. Retrying..."
            )
            attempt += 1
            time.sleep(5)  # Wait for 5 seconds before retrying

    print(f"{Fore.RED}Max attempts reached. WhatsApp session not logged in.")
    return False


def clean_to_bmp(text: str) -> str:
    """
    Cleans the given text by removing any characters that are not within the BMP (Basic Multilingual Plane) range.

    Args:
        text (str): The text to be cleaned.

    Returns:
        str: The cleaned text.

    """
    return re.sub(r"[^\U00000000-\U0000FFFF]", "", text)


def send_whatsapp_message(
    driver: WebDriver,
    recipient: str,
    df: pd.DataFrame,
    file_path: str,
    chat_type: str = "contact",
) -> None:
    """
    Sends a WhatsApp message to the specified recipient using the provided driver and DataFrame.

    Args:
        driver (WebDriver): The WebDriver instance for controlling the browser.
        recipient (str): The recipient's WhatsApp number or group name.
        df (pd.DataFrame): The DataFrame containing the message data.
        file_path (str): The file path to save the updated DataFrame.
        chat_type (str, optional): The type of chat to send the message to. Defaults to 'contact'.

    Raises:
        NoSuchElementException: If an element is not found on the page.

    Returns:
        None
    """
    try:
        if not is_logged_in(driver):
            print(f"{Fore.RED}WhatsApp session not logged in. Please log in and retry.")
            return

        # Check if message has already been sent or has an error
        current_status = df.loc[df["whatsapp_number"] == recipient, "status"].iloc[0]
        if current_status.lower() == "send" or current_status.lower() == "error":
            print(
                f"{Fore.YELLOW}Message for {recipient} already processed with status: {current_status}"
            )
            return

        # Open WhatsApp Web if not already open
        if driver.current_url != "https://web.whatsapp.com/":
            driver.get("https://web.whatsapp.com/")

        # Wait for WhatsApp Web to load
        wait = WebDriverWait(driver, 30)
        wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
            )
        )

        # Select recipient (group or contact)
        if chat_type == "group":
            group = wait.until(
                EC.element_to_be_clickable((By.XPATH, f"//span[@title='{recipient}']"))
            )
            group.click()
        else:  # Assume 'contact' as default
            chat_input = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
                )
            )
            chat_input.clear()
            chat_input.send_keys(recipient)
            chat_input.send_keys(Keys.ENTER)

        # Wait for chat to load
        time.sleep(3)  # Adjust time as needed

        message_fields = {
            "name": df.loc[df["whatsapp_number"] == recipient, "name"].iloc[0].upper(),
            "vehicle_license": df.loc[
                df["whatsapp_number"] == recipient, "vehicle_license"
            ]
            .iloc[0]
            .upper(),
            "model_vehicle": df.loc[df["whatsapp_number"] == recipient, "model_vehicle"]
            .iloc[0]
            .upper(),
            "expire_date": df.loc[
                df["whatsapp_number"] == recipient, "expire_date"
            ].iloc[0],
        }

        message = template.format(**message_fields)

        # Clean message to BMP characters only
        clean_message = clean_to_bmp(message)

        # Locate message input field and send message
        message_input = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "/html/body/div[1]/div/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p",
                )
            )
        )

        # Clear any existing text in message input field
        message_input.clear()

        # Send new message with multiline support
        lines = clean_message.splitlines()
        for line in lines:
            message_input.send_keys(line)
            message_input.send_keys(
                Keys.SHIFT + Keys.ENTER
            )  # Insert new line without sending message

        # Send final message
        message_input.send_keys(Keys.ENTER)
        print(f"{Fore.GREEN}Message sent to {recipient}")

        # Update status and timestamp in DataFrame
        df.loc[df["whatsapp_number"] == recipient, "status"] = "Correct"
        df.loc[df["whatsapp_number"] == recipient, "timestep"] = datetime.today().date()

        # Save DataFrame with updated status
        df.to_excel(file_path, index=False)

    except NoSuchElementException as e:
        print(f"{Fore.RED}Error: Element not found - {str(e)}")
        # Update status to 'Error' in DataFrame
        df.loc[df["whatsapp_number"] == recipient, "status"] = "Error"
        df.loc[df["whatsapp_number"] == recipient, "timestep"] = datetime.today().date()
        df.to_excel(file_path, index=False)  # Save DataFrame with updated status
    except Exception as e:
        print(f"{Fore.RED}Error sending message to {recipient}: {str(e)}")
        # Update status to 'Error' in DataFrame
        df.loc[df["whatsapp_number"] == recipient, "status"] = "Error"
        df.loc[df["whatsapp_number"] == recipient, "timestep"] = datetime.today().date()
        df.to_excel(file_path, index=False)  # Save DataFrame with updated status


def load_excel_data(file_path: str, sheet_name: str = "Sheet1") -> pd.DataFrame | None:
    """
    Load data from an Excel file and return it as a pandas DataFrame.

    Args:
        file_path (str): The path to the Excel file.
        sheet_name (str, optional): The name of the sheet to load data from. Defaults to 'Sheet1'.

    Returns:
        pandas.DataFrame or None: The loaded data as a DataFrame, or None if there was an error.

    Raises:
        FileNotFoundError: If the specified Excel file does not exist.

    """
    try:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        except FileNotFoundError:
            df = None

        if df is None:
            data = {
                "surname": ["Rosas"],
                "name": ["Nahuel"],
                "whatsapp_number": ["3517885067"],
                "expire_date": [pd.Timestamp.today() + pd.Timedelta(days=2)],
                "vehicle_license": ["AE325CB"],
                "model_vehicle": ["cronos"],
                "status": [""],
                "timestep": [pd.Timestamp.today()],
            }
            df = pd.DataFrame(data)
            df.to_excel(file_path, index=False)
            print(
                f"{Fore.GREEN}File '{file_path}' created successfully as a testing template."
            )

        # Convert 'status' column to string and fill NaN values with an empty string
        if "status" in df.columns:
            df["status"] = df["status"].astype(str).fillna("")
        else:
            df["status"] = ""

        print(f"{Fore.GREEN}Excel data loaded successfully.")
        return df

    except Exception as e:
        print(f"{Fore.RED}Error reading or creating Excel file: {str(e)}")
        return None


def verbose_excel_data(file_path: str) -> None:
    """
    Display the content of an Excel file in a verbose manner.

    Args:
        file_path (str): The path to the Excel file.

    Raises:
        Exception: If there is an error displaying the Excel content.

    """
    try:
        with pd.ExcelFile(file_path) as xls:
            for sheet in xls.sheet_names:
                print(f"{Fore.BLUE}Sheet: {sheet}")
                df = pd.read_excel(xls, sheet_name=sheet)
                print(df.to_string(index=False))
                print("=" * 80)
    except Exception as e:
        print(f"{Fore.RED}Error displaying Excel content: {str(e)}")


def preview_whatsapp_message(df):
    """
    Preview the WhatsApp message for each recipient in the given DataFrame.

    Args:
        df (pandas.DataFrame): The DataFrame containing the WhatsApp message data.

    Raises:
        KeyError: If a required key is missing in the Excel data.
        Exception: If an error occurs while previewing the message.

    Returns:
        None
    """
    try:
        for index, row in df.iterrows():

            expire_date = (
                datetime.strptime(row["expire_date"], "%d/%m/%Y")
                if isinstance(row["expire_date"], str)
                else row["expire_date"]
            )
            recipient = row["whatsapp_number"]
            send_date = expire_date - timedelta(days=2)

            if datetime.now().date() == send_date.date():
                message_fields = {
                    "name": df.loc[df["whatsapp_number"] == recipient, "name"]
                    .iloc[0]
                    .upper(),
                    "vehicle_license": df.loc[
                        df["whatsapp_number"] == recipient, "vehicle_license"
                    ]
                    .iloc[0]
                    .upper(),
                    "model_vehicle": df.loc[
                        df["whatsapp_number"] == recipient, "model_vehicle"
                    ]
                    .iloc[0]
                    .upper(),
                    "expire_date": df.loc[
                        df["whatsapp_number"] == recipient, "expire_date"
                    ].iloc[0],
                }

                message = template.format(**message_fields)
                clean_message = clean_to_bmp(message)
                clean_message = clean_to_bmp(message)

                print(
                    f"{Fore.BLUE}Preview message for {row['whatsapp_number']}:\n{clean_message}\n"
                )

    except KeyError as e:
        print(f"{Fore.RED}Error: Missing key in Excel data: {str(e)}")
    except Exception as e:
        print(f"{Fore.RED}Error previewing message: {str(e)}")


def display_pending_messages(df, today):
    """
    Display pending messages to be sent.

    Args:
        df (pandas.DataFrame): The DataFrame containing the messages.
        today (datetime.date): The current date.

    Returns:
        None

    Raises:
        None
    """
    try:
        pending_df = df[
            (df["expire_date"].apply(lambda x: x.date()) - timedelta(days=2) == today)
            & (~df["status"].str.lower().isin(["correct", "error"]))
        ]
        if not pending_df.empty:
            print(f"{Fore.YELLOW}Pending messages to be sent:")
            print(pending_df)
        else:
            print(
                f"{Fore.GREEN}No messages need to be sent today or all messages already sent."
            )
    except Exception as e:
        print(f"{Fore.RED}Error processing data: {str(e)}")


if __name__ == "__main__":
    setup_chromedriver()  # Setup Chromedriver automatically

    file_path = "data.xlsx"
    df = load_excel_data(file_path)

    if df is not None:
        # Set up user data directory for Chrome
        user_data_dir = os.path.join(os.getcwd(), "chrome_user_data")
        if not os.path.exists(user_data_dir):
            os.makedirs(user_data_dir)

        # Start WebDriver with user data directory
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument(f"--user-data-dir={user_data_dir}")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-plugins")
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-infobars")
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_argument("--disable-popup-blocking")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option("useAutomationExtension", False)
        driver = webdriver.Chrome(options=chrome_options)
        driver.get("https://web.whatsapp.com/")

        try:
            if not is_logged_in(driver):
                print(f"{Fore.CYAN}Please scan the QR code to log in to WhatsApp Web.")

            today = datetime.today().date()

            while True:
                print(f"\n{Fore.CYAN}Menu:")
                print(f"{Fore.CYAN}1. View the table of users to send messages")
                print(f"{Fore.CYAN}2. Verify messages")
                print(f"{Fore.CYAN}3. Send messages")
                print(f"{Fore.CYAN}4. Preview messages")
                print(f"{Fore.CYAN}5. Exit")

                choice = input("Enter your choice (1-5): ")

                if choice == "1":
                    if df is not None:
                        verbose_excel_data(file_path)
                    else:
                        print(
                            f"{Fore.RED}Excel data is not loaded. Please load the data first."
                        )

                elif choice == "2":
                    if df is not None:
                        pending_df = df[
                            ~df["status"].str.lower().isin(["correct", "error"])
                        ]
                        if not pending_df.empty:
                            print(
                                f"{Fore.CYAN}Reloading the Excel file to check for changes..."
                            )
                            df = load_excel_data(file_path)
                            if df is not None:
                                display_pending_messages(df, today)
                            else:
                                print(f"{Fore.RED}Failed to reload the Excel file.")
                        else:
                            print(f"{Fore.GREEN}No pending messages to verify.")
                    else:
                        print(
                            f"{Fore.RED}Excel data is not loaded. Please load the data first."
                        )

                elif choice == "3":
                    if not is_logged_in(driver):
                        print(
                            f"{Fore.RED}WhatsApp session not logged in. Please log in to WhatsApp Web."
                        )

                    if df is not None:
                        pending_df = df[
                            ~df["status"].str.lower().isin(["correct", "error"])
                        ]
                        if not pending_df.empty:
                            for index, row in pending_df.iterrows():
                                try:
                                    expire_date = (
                                        datetime.strptime(
                                            row["expire_date"], "%d/%m/%Y"
                                        )
                                        if isinstance(row["expire_date"], str)
                                        else row["expire_date"]
                                    )
                                    send_date = expire_date - timedelta(days=2)

                                    if datetime.now().date() == send_date.date():
                                        send_whatsapp_message(
                                            driver,
                                            row["whatsapp_number"],
                                            df,
                                            file_path,
                                        )

                                except KeyError as e:
                                    print(
                                        f"{Fore.RED}Error: Missing key in Excel data: {str(e)}"
                                    )
                                except Exception as e:
                                    print(f"{Fore.RED}Error processing data: {str(e)}")
                        else:
                            print(f"{Fore.GREEN}No pending messages to send.")
                    else:
                        print(
                            f"{Fore.RED}Excel data is not loaded. Please load the data first."
                        )

                elif choice == "4":
                    if df is not None:
                        pending_df = df[
                            ~df["status"].str.lower().isin(["correct", "error"])
                        ]
                        if not pending_df.empty:
                            preview_whatsapp_message(df)
                        else:
                            print(f"{Fore.GREEN}No pending messages to preview.")
                    else:
                        print(
                            f"{Fore.RED}Excel data is not loaded. Please load the data first."
                        )

                elif choice == "5":
                    break
                else:
                    print(f"{Fore.RED}Invalid choice. Please try again.")

        finally:
            driver.quit()  # Ensure WebDriver is closed
            sys.exit(0)
            
    else:
        verbose_excel_data(file_path)
