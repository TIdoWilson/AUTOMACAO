$PyExe = "C:\Users\ECD\AppData\Local\Programs\Python\Python311\python.exe"
$VenvName = "conjunto_processadores_ecd"
$VenvRoot = "W:\python\Venvs Exclusivos Servidor"
$VenvPath = Join-Path $VenvRoot $VenvName
$ScriptDir = "W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\Conjunto Processadores de ECDs"
$InfoTxt = Join-Path $ScriptDir "venv_exclusivo.txt"

New-Item -ItemType Directory -Force -Path $VenvRoot | Out-Null
& $PyExe -m venv $VenvPath
& "$VenvPath\Scripts\Activate.ps1"

python -m pip install --upgrade pip
python -m pip install pyautogui pywinauto pywin32 Pillow

@"
VENV: $VenvPath
SCRIPT: Conjunto Processadores de ECDs
"@ | Set-Content -Path $InfoTxt -Encoding UTF8
