#!/bin/bash

# Añadir directorio local bin al PATH si no está ya presente
if [[ ":$PATH:" != *":$HOME/.local/bin:"* ]]; then
    export PATH="$PATH:$HOME/.local/bin"
fi

# Function to install Google Chrome
install_chrome() {
    echo "Installing Google Chrome..."
    wget https://dl.google.com/linux/direct/google-chrome-stable_current_x86_64.rpm
    sudo dnf install -y ./google-chrome-stable_current_x86_64.rpm
    rm -f google-chrome-stable_current_x86_64.rpm
}

# Check and install Google Chrome if not installed
if ! command -v google-chrome &> /dev/null; then
    install_chrome
else
    echo "Google Chrome is already installed."
fi

# Install Python and pip (with sudo for dnf)
echo "Installing Python and pip..."
sudo dnf install -y python3 python3-pip

# Install common Python dependencies (with sudo for dnf)
echo "Installing common Python dependencies..."
sudo dnf install -y gcc openssl-devel bzip2-devel libffi-devel

# Install additional useful development tools (with sudo for dnf)
echo "Installing additional development tools..."
sudo dnf install -y git

# Install Python packages (without sudo)
echo "Installing Python packages..."
pip3 install --user pandas selenium chromedriver-autoinstaller openpyxl Jinja2 colorama

# Upgrade pip (without sudo)
echo "Upgrading pip..."
pip3 install --user --upgrade pip --no-warn-script-location

echo "Installation completed."

# Verify installed versions
python3 --version
pip3 --version

# Create Excel file if it does not exist
create_excel() {
    echo "Creating 'data.xlsx' file..."
    python3 - <<EOF
import pandas as pd

file_path = 'data.xlsx'

data = {
    'surname': ['Rosas'],
    'name': ['Nahuel'],
    'whatsapp_number': ['3517885067'],
    'expire_date': [pd.Timestamp('28/06/2024')],
    'vehicle_license': ['AE325CB'],
    'model_vehicle': ['cronos'],
    'status': [''],
    'timestep': [pd.Timestamp('26/06/2024')],
}

df = pd.DataFrame(data)
df.to_excel(file_path, index=False)
print(f"File '{file_path}' created successfully.")
EOF
}

if [ ! -f "data.xlsx" ]; then
    create_excel
else
    echo "'data.xlsx' file already exists."
fi
