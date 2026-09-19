$ErrorActionPreference = 'Stop'

$projectDir = $PSScriptRoot
$venvPython = Join-Path $projectDir '.venv\Scripts\python.exe'
$converterScript = Join-Path $projectDir 'xlsx_to_html_converter_all_in_one.py'

if (-not (Test-Path -LiteralPath $venvPython)) {
    throw 'Python仮想環境がありません。先に .\setup_python_env.ps1 を実行してください。'
}

Push-Location $projectDir
try {
    & $venvPython -X utf8 $converterScript
    if ($LASTEXITCODE -ne 0) {
        throw "変換処理に失敗しました（終了コード: $LASTEXITCODE）"
    }
}
finally {
    Pop-Location
}
