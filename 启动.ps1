$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$appPython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $appPython)) {
    throw '请先创建 .venv 并安装 requirements.txt 中的依赖。'
}
& $appPython -m streamlit run app.py --server.address 127.0.0.1
