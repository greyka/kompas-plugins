# Скачивает и ставит переносной Python для плагинов КОМПАС (нужная, проверенная версия).
#   python.org: embeddable Python   +   PyPI: pywin32 (связь с КОМПАС по COM)
# Ставится в отдельную папку, общую для всех плагинов. Права администратора не нужны.
param([Parameter(Mandatory = $true)][string]$Target)

$PythonVersion  = '3.14.7'
$Pywin32Version = '312'
$Tag            = 'cp314'          # должен соответствовать версии Python

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$tmp = Join-Path $env:TEMP ("kompas-python-" + [guid]::NewGuid())
New-Item -ItemType Directory $tmp | Out-Null
try {
    Write-Host "Скачиваю Python $PythonVersion ..."
    $pyZip = Join-Path $tmp 'python.zip'
    Invoke-WebRequest "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-embed-amd64.zip" -OutFile $pyZip -UseBasicParsing

    Write-Host "Скачиваю pywin32 $Pywin32Version ..."
    $info = Invoke-RestMethod "https://pypi.org/pypi/pywin32/$Pywin32Version/json"
    $wheel = $info.urls | Where-Object { $_.filename -eq "pywin32-$Pywin32Version-$Tag-$Tag-win_amd64.whl" } | Select-Object -First 1
    if (-not $wheel) { throw "На PyPI нет pywin32 $Pywin32Version для $Tag" }
    $whl = Join-Path $tmp 'pywin32.zip'
    Invoke-WebRequest $wheel.url -OutFile $whl -UseBasicParsing
    if ((Get-FileHash $whl -Algorithm SHA256).Hash -ne $wheel.digests.sha256.ToUpper()) { throw 'pywin32: не совпала контрольная сумма' }

    Write-Host "Устанавливаю в $Target ..."
    if (Test-Path $Target) { Remove-Item $Target -Recurse -Force }
    Expand-Archive $pyZip -DestinationPath $Target
    $site = Join-Path $Target 'Lib\site-packages'
    $wx = Join-Path $tmp 'wheel'
    Expand-Archive $whl -DestinationPath $wx
    New-Item -ItemType Directory $site -Force | Out-Null
    Get-ChildItem $wx | Where-Object { $_.Name -notlike '*.dist-info' -and $_.Name -notlike '*.data' } |
        Move-Item -Destination $site
    # лишнее: IDE pythonwin, справка, демо
    'pythonwin', 'PyWin32.chm', 'adodbapi', 'isapi', 'win32\Demos' | ForEach-Object {
        $p = Join-Path $site $_; if (Test-Path $p) { Remove-Item $p -Recurse -Force } }

    # пути поиска модулей встроенного Python
    $pth = Get-ChildItem $Target -Filter 'python*._pth' | Select-Object -First 1
    $zipName = (Get-ChildItem $Target -Filter 'python*.zip' | Select-Object -First 1).Name
    [IO.File]::WriteAllText($pth.FullName, "$zipName`r`n.`r`nLib\site-packages`r`nimport site`r`n")

    $ok = & (Join-Path $Target 'python.exe') -c "import win32com.client, pythoncom; print('ok')"
    if ($ok -ne 'ok') { throw 'Python установлен, но pywin32 не загружается' }
    Set-Content (Join-Path $Target 'VERSION.txt') "Python $PythonVersion, pywin32 $Pywin32Version"
    Write-Host "Python готов: $Target"
}
catch {
    # не оставлять недоделанную папку — иначе установщик решит, что Python уже есть
    if (Test-Path $Target) { Remove-Item $Target -Recurse -Force -ErrorAction SilentlyContinue }
    Write-Host "Не удалось: $($_.Exception.Message)"
    exit 1
}
finally {
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
