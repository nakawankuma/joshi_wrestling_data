$ErrorActionPreference = 'Stop'

$projectDir = $PSScriptRoot
$venvDir = Join-Path $projectDir '.venv'
$venvPython = Join-Path $venvDir 'Scripts\python.exe'
$requirementsFile = Join-Path $projectDir 'requirements.txt'

if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Host 'Python仮想環境を作成します: .venv'
    & py -3 -m venv $venvDir
    if ($LASTEXITCODE -ne 0) {
        throw "仮想環境の作成に失敗しました（終了コード: $LASTEXITCODE）"
    }
}
else {
    Write-Host '既存のPython仮想環境を更新します: .venv'
}

Write-Host 'pipを更新します'
& $venvPython -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) {
    throw "pipの更新に失敗しました（終了コード: $LASTEXITCODE）"
}

Write-Host '必要なPythonパッケージをインストールします'
& $venvPython -m pip install -r $requirementsFile
if ($LASTEXITCODE -ne 0) {
    throw "パッケージのインストールに失敗しました（終了コード: $LASTEXITCODE）"
}

Write-Host 'Python環境の準備が完了しました'
& $venvPython --version
