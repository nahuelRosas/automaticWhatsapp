$localPath = "$env:USERPROFILE\AppData\Local\Microsoft\WindowsApps"
if (-not ([System.Environment]::GetEnvironmentVariable('PATH', 'User') -split ';' -contains $localPath)) {
    $newPath = [System.Environment]::GetEnvironmentVariable('PATH', 'User') + ";$localPath"
    [System.Environment]::SetEnvironmentVariable('PATH', $newPath, 'User')
}

function Install-Chrome {
    Write-Output "Installing Google Chrome..."
    Invoke-WebRequest -Uri "https://dl.google.com/chrome/install/latest/chrome_installer.exe" -OutFile "$env:TEMP\chrome_installer.exe"
    Start-Process -FilePath "$env:TEMP\chrome_installer.exe" -ArgumentList "/install" -Wait
    Remove-Item -Path "$env:TEMP\chrome_installer.exe" -Force
}

if (-not (Get-Command "chrome.exe" -ErrorAction SilentlyContinue)) {
    Install-Chrome
}
else {
    Write-Output "Google Chrome is already installed."
}

Write-Output "Installing Python and pip..."
choco install python3

Write-Output "Installing common Python dependencies..."
choco install gcc openssl bzip2 libffi

Write-Output "Installing additional development tools..."
choco install git

Write-Output "Installing Python packages..."
pip install pandas selenium openpyxl Jinja2 colorama

Write-Output "Upgrading pip..."
pip install --upgrade pip

Write-Output "Installation completed."

python --version
pip --version

function Create-Excel {
    Write-Output "Creating 'data.xlsx' file..."
    python - <<'EOF'
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

if (-not (Test-Path "data.xlsx")) {
    Create-Excel
}
else {
    Write-Output "'data.xlsx' file already exists."
}
