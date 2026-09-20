param(
    [ValidateRange(1, 65535)]
    [int]$Port = 8000
)

$ErrorActionPreference = 'Stop'

$projectDir = $PSScriptRoot
$venvPython = Join-Path $projectDir '.venv\Scripts\python.exe'

if (-not (Test-Path -LiteralPath $venvPython)) {
    throw 'Python仮想環境がありません。先に .\setup_python_env.ps1 を実行してください。'
}

$url = "http://localhost:$Port/index.html"
Write-Host "ローカルWebサーバーを起動します: $url"
Write-Host '終了するときは Ctrl+C を押してください。'

# 外部端末からアクセスされないよう、ローカルホストだけで待ち受ける。
Push-Location $projectDir
try {
    & $venvPython -X utf8 -m http.server $Port --bind 127.0.0.1 --directory $projectDir
    if ($LASTEXITCODE -ne 0) {
        throw "ローカルWebサーバーが終了しました（終了コード: $LASTEXITCODE）"
    }
}
finally {
    Pop-Location
}
