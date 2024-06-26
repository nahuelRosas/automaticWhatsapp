import re
import pandas as pd
from datetime import datetime, timedelta
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

# Initialize colorama
init(autoreset=True)

# Generar dinámicamente la plantilla y el mensaje
template = (
    "_¡Hola {name}!_\n\n"
    "_*Recordatorio Importante*_\n\n"
    "_Queremos recordarte que tu seguro para el vehículo {model_vehicle} con patente {vehicle_license} vence el {expire_date:%d/%m/%Y}._\n\n"
    "_Desde ya, muchas gracias por confiar en nosotros._\n\n"
    "_No olvides que puedes pagar de manera sencilla y segura *por este medio.*_\n\n"
    "_Si ya realizaste el pago, *ignora este mensaje*._\n"
    "_Para dejar de recibir estos recordatorios, simplemente envíanos un mensaje con la palabra *cancelar*._"
)


# Function to ensure Chromedriver is up-to-date
def setup_chromedriver():
    try:
        chromedriver_autoinstaller.install()
        print(f"{Fore.GREEN}Chromedriver is up to date.")
    except Exception as e:
        print(f"{Fore.RED}Error setting up Chromedriver: {str(e)}")
        exit()

# Function to check if WhatsApp session is logged in
def is_logged_in(driver):
    max_attempts = 24  # 2 minutes (30 seconds x 24 attempts)
    attempt = 1

    while attempt <= max_attempts:
        try:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')))
            return True
        except TimeoutException:
            print(f"{Fore.YELLOW}Attempt {attempt}: Timeout - WhatsApp session not logged in yet. Retrying...")
            attempt += 1
            time.sleep(5)  # Wait for 5 seconds before retrying

    print(f"{Fore.RED}Max attempts reached. WhatsApp session not logged in.")
    return False

# Function to clean text to only BMP characters
def clean_to_bmp(text):
    return re.sub(r'[^\U00000000-\U0000FFFF]', '', text)

# Function to send WhatsApp message using Chrome
def send_whatsapp_message(driver, recipient, df, file_path, chat_type='contact'):
    try:
        if not is_logged_in(driver):
            print(f"{Fore.RED}WhatsApp session not logged in. Please log in and retry.")
            return
        
        # Verificar si el mensaje ya fue enviado o tiene un error
        current_status = df.loc[df['whatsapp_number'] == recipient, 'status'].iloc[0]
        if current_status.lower() == 'send' or current_status.lower() == 'error':
            print(f"{Fore.YELLOW}Message for {recipient} already processed with status: {current_status}")
            return
        
        # Abrir WhatsApp Web si no está abierto
        if driver.current_url != 'https://web.whatsapp.com/':
            driver.get('https://web.whatsapp.com/')
        
        # Esperar a que cargue WhatsApp Web
        wait = WebDriverWait(driver, 30)
        wait.until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')))
        
        # Seleccionar destinatario (grupo o contacto)
        if chat_type == 'group':
            group = wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[@title='{recipient}']")))
            group.click()
        else:  # Suponer 'contact' como valor por defecto
            chat_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')))
            chat_input.clear()
            chat_input.send_keys(recipient)
            chat_input.send_keys(Keys.ENTER)
        
        # Esperar a que cargue el chat
        time.sleep(3)  # Ajustar el tiempo según sea necesario
                
        message_fields = {
            'name': df.loc[df['whatsapp_number'] == recipient, 'name'].iloc[0].upper(),
            'vehicle_license': df.loc[df['whatsapp_number'] == recipient, 'vehicle_license'].iloc[0].upper(),
            'model_vehicle': df.loc[df['whatsapp_number'] == recipient, 'model_vehicle'].iloc[0].upper(),
            'expire_date': df.loc[df['whatsapp_number'] == recipient, 'expire_date'].iloc[0],
        }
        
        message = template.format(**message_fields)
        
        # Limpiar el mensaje para caracteres solo BMP
        clean_message = clean_to_bmp(message)
        
        # Localizar el campo de entrada del mensaje y enviar el mensaje
        message_input = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p')))
        
        # Limpiar cualquier texto existente en el campo de entrada del mensaje
        message_input.clear()
        
        # Enviar el nuevo mensaje con soporte multilínea
        lines = clean_message.splitlines()
        for line in lines:
            message_input.send_keys(line)
            message_input.send_keys(Keys.SHIFT + Keys.ENTER)  # Insertar nueva línea sin enviar el mensaje
        
        # Enviar el mensaje final
        message_input.send_keys(Keys.ENTER)
        print(f"{Fore.GREEN}Message sent to {recipient}")
        
        # Actualizar estado y tiempo en el DataFrame
        df.loc[df['whatsapp_number'] == recipient, 'status'] = 'Correct'
        df.loc[df['whatsapp_number'] == recipient, 'timestep'] = datetime.today().date()
        
        # Guardar DataFrame con el estado actualizado
        df.to_excel(file_path, index=False)
        
    except NoSuchElementException as e:
        print(f"{Fore.RED}Error: Element not found - {str(e)}")
        # Actualizar estado a 'Error' en el DataFrame
        df.loc[df['whatsapp_number'] == recipient, 'status'] = 'Error'
        df.loc[df['whatsapp_number'] == recipient, 'timestep'] = datetime.today().date()
        df.to_excel(file_path, index=False)  # Guardar DataFrame con estado actualizado
    except Exception as e:
        print(f"{Fore.RED}Error sending message to {recipient}: {str(e)}")
        # Actualizar estado a 'Error' en el DataFrame
        df.loc[df['whatsapp_number'] == recipient, 'status'] = 'Error'
        df.loc[df['whatsapp_number'] == recipient, 'timestep'] = datetime.today().date()
        df.to_excel(file_path, index=False)  # Guardar DataFrame con estado actualizado
        
# Load Excel data
def load_excel_data(file_path, sheet_name='Sheet1'):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Convert 'status' column to string and fill NaN values with an empty string
        df['status'] = df['status'].astype(str).fillna('')
        
        # Add 'status' column if not exists
        if 'status' not in df.columns:
            df['status'] = ''
        
        print(f"{Fore.GREEN}Excel data loaded successfully.")
        return df
    except FileNotFoundError:
        print(f"{Fore.RED}Error: File '{file_path}' not found.")
        return None
    except Exception as e:
        print(f"{Fore.RED}Error reading Excel file: {str(e)}")
        return None

# Display Excel data
def verbose_excel_data(file_path, sheet_name='Sheet1'):
    try:
        with pd.ExcelFile(file_path) as xls:
            for sheet in xls.sheet_names:
                print(f"{Fore.BLUE}Sheet: {sheet}")
                df = pd.read_excel(xls, sheet_name=sheet)
                print(df.to_string(index=False)) 
                print("=" * 80) 
    except Exception as e:
        print(f"{Fore.RED}Error displaying Excel content: {str(e)}")

# Preview WhatsApp message
def preview_whatsapp_message(df):
    try:
        for index, row in df.iterrows():
            
            expire_date = datetime.strptime(row['expire_date'], '%d/%m/%Y') if isinstance(row['expire_date'], str) else row['expire_date']
            recipient = row['whatsapp_number']
            send_date = expire_date - timedelta(days=2)

            if datetime.now().date() == send_date.date():
                message_fields = {
                        'name': df.loc[df['whatsapp_number'] == recipient, 'name'].iloc[0].upper(),
                        'vehicle_license': df.loc[df['whatsapp_number'] == recipient, 'vehicle_license'].iloc[0].upper(),
                        'model_vehicle': df.loc[df['whatsapp_number'] == recipient, 'model_vehicle'].iloc[0].upper(),
                        'expire_date': df.loc[df['whatsapp_number'] == recipient, 'expire_date'].iloc[0],
                    }
        
                message = template.format(**message_fields)
                clean_message = clean_to_bmp(message)
                clean_message = clean_to_bmp(message)
                
                print(f"{Fore.BLUE}Preview message for {row['whatsapp_number']}:\n{clean_message}\n")
    
    except KeyError as e:
        print(f"{Fore.RED}Error: Missing key in Excel data: {str(e)}")
    except Exception as e:
        print(f"{Fore.RED}Error previewing message: {str(e)}")


# Display the data that needs to be processed
def display_pending_messages(df, today):
    try:
        pending_df = df[(df['expire_date'].apply(lambda x: x.date()) - timedelta(days=2) == today) & (~df['status'].str.lower().isin(['send', 'error']))]
        if not pending_df.empty:
            print(f"{Fore.YELLOW}Pending messages to be sent:")
            print(pending_df)
        else:
            print(f"{Fore.GREEN}No messages need to be sent today or all messages already sent.")
    except Exception as e:
        print(f"{Fore.RED}Error processing data: {str(e)}")

# Main process
if __name__ == "__main__":
    setup_chromedriver()  # Setup Chromedriver automatically
    
    file_path = 'data.xlsx'
    df = load_excel_data(file_path)
    
    if df is not None:
        # Set up user data directory for Chrome
        user_data_dir = os.path.join(os.getcwd(), "chrome_user_data")
        if not os.path.exists(user_data_dir):
            os.makedirs(user_data_dir)
        
        # Start WebDriver with user data directory
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument(f"--user-data-dir={user_data_dir}")
        driver = webdriver.Chrome(options=chrome_options)
        driver.get('https://web.whatsapp.com/')
        
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
                
                if choice == '1':
                    if df is not None:
                        verbose_excel_data(file_path)
                    else:
                        print(f"{Fore.RED}Excel data is not loaded. Please load the data first.")
                
                elif choice == '2':
                    if df is not None:
                        pending_df = df[~df['status'].str.lower().isin(['correct', 'error'])]
                        if not pending_df.empty:
                            print(f"{Fore.CYAN}Reloading the Excel file to check for changes...")
                            df = load_excel_data(file_path)
                            if df is not None:
                                display_pending_messages(df, today)
                            else:
                                print(f"{Fore.RED}Failed to reload the Excel file.")
                        else:
                            print(f"{Fore.GREEN}No pending messages to verify.")
                    else:
                        print(f"{Fore.RED}Excel data is not loaded. Please load the data first.")
                
                elif choice == '3':
                    if not is_logged_in(driver):
                        print(f"{Fore.RED}WhatsApp session not logged in. Please log in to WhatsApp Web.")
                    
                    if df is not None:
                        pending_df = df[~df['status'].str.lower().isin(['correct', 'error'])]
                        if not pending_df.empty:
                            for index, row in pending_df.iterrows():
                                try:
                                    expire_date = datetime.strptime(row['expire_date'], '%d/%m/%Y') if isinstance(row['expire_date'], str) else row['expire_date']
                                    send_date = expire_date - timedelta(days=2)
                                    
                                    if  datetime.now().date() == send_date.date():
                                        send_whatsapp_message(driver, row['whatsapp_number'], df, file_path)
                                
                                except KeyError as e:
                                    print(f"{Fore.RED}Error: Missing key in Excel data: {str(e)}")
                                except Exception as e:
                                    print(f"{Fore.RED}Error processing data: {str(e)}")
                        else:
                            print(f"{Fore.GREEN}No pending messages to send.")
                    else:
                        print(f"{Fore.RED}Excel data is not loaded. Please load the data first.")
                
                elif choice == '4':
                    if df is not None:
                        pending_df = df[~df['status'].str.lower().isin(['correct', 'error'])]
                        if not pending_df.empty:
                            preview_whatsapp_message(df)
                        else:
                            print(f"{Fore.GREEN}No pending messages to preview.")
                    else:
                        print(f"{Fore.RED}Excel data is not loaded. Please load the data first.")
                
                elif choice == '5':
                    break
                else:
                    print(f"{Fore.RED}Invalid choice. Please try again.")
        
        finally:
            driver.quit()  # Ensure WebDriver is closed
    else:
        verbose_excel_data(file_path)
