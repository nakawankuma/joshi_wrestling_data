$ErrorActionPreference = 'Stop'

$projectDir = $PSScriptRoot
$venvPython = Join-Path $projectDir '.venv\Scripts\python.exe'

if (-not (Test-Path -LiteralPath $venvPython)) {
    throw 'Python仮想環境がありません。先に .\setup_python_env.ps1 を実行してください。'
}

# 呼び出し元のカレントディレクトリに依存せず、プロジェクト直下のテストを実行する。
Push-Location $projectDir
try {
    & $venvPython -X utf8 -m unittest discover -s $projectDir -p 'test_*.py' -v
    if ($LASTEXITCODE -ne 0) {
        throw "Pythonテストに失敗しました（終了コード: $LASTEXITCODE）"
    }

    Write-Host 'Pythonテストがすべて成功しました'
}
finally {
    Pop-Location
}
