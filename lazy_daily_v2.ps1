# lazy_daily_windows_select_all_except_docx_fast.ps1
# Windows 11 / PowerShell version
# File encoding: UTF-8 with BOM
# Faster version: extract ZIP to one task folder, open folder, select all files except DOCX.

chcp 65001 | Out-Null
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$Downloads = Join-Path $env:USERPROFILE "Downloads"
$BaseDir = Join-Path $env:USERPROFILE "Documents\Daily"
$HistoryFile = Join-Path $BaseDir ".processed_archives.txt"

# Можно настроить:
$OpenDocxAutomatically = $true     # $false, если Word тормозит и docx лучше открывать вручную
$ExplorerDelayMs = 350             # пауза для Проводника перед выделением файлов
$NoZipCheckDelayMs = 1200          # как часто проверять папку Downloads, если архивов нет
$DownloadCheckDelayMs = 500        # как часто проверять, что ZIP докачался

if (-not (Test-Path $BaseDir)) {
    New-Item -ItemType Directory -Path $BaseDir | Out-Null
}

if (-not (Test-Path $HistoryFile)) {
    New-Item -ItemType File -Path $HistoryFile | Out-Null
}

function Wait-ForFileReadyFast {
    param(
        [string]$Path,
        [int]$Checks = 2,
        [int]$DelayMilliseconds = 500
    )

    $stableCount = 0
    $lastSize = -1

    while ($stableCount -lt $Checks) {
        if (-not (Test-Path $Path)) {
            Start-Sleep -Milliseconds $DelayMilliseconds
            continue
        }

        $file = Get-Item $Path
        $currentSize = $file.Length

        try {
            $stream = [System.IO.File]::Open($Path, 'Open', 'Read', 'None')
            $stream.Close()
            $canOpen = $true
        }
        catch {
            $canOpen = $false
        }

        if ($canOpen -and $currentSize -eq $lastSize) {
            $stableCount++
        }
        else {
            $stableCount = 0
        }

        $lastSize = $currentSize
        Start-Sleep -Milliseconds $DelayMilliseconds
    }
}

function Get-NextTaskNumber {
    param([string]$Directory)

    $maxNumber = 0

    Get-ChildItem -Path $Directory -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.Name -match '^Задача_(\d+)$') {
            $number = [int]$Matches[1]
            if ($number -gt $maxNumber) {
                $maxNumber = $number
            }
        }
    }

    return ($maxNumber + 1)
}

function Select-FilesInExplorerFast {
    param(
        [string]$FolderPath,
        [string[]]$FilePaths
    )

    if (-not $FilePaths -or $FilePaths.Count -eq 0) {
        Start-Process explorer.exe -ArgumentList "`"$FolderPath`""
        return
    }

    Start-Process explorer.exe -ArgumentList "`"$FolderPath`""
    Start-Sleep -Milliseconds $ExplorerDelayMs

    try {
        $shell = New-Object -ComObject Shell.Application
        $windows = $shell.Windows()

        foreach ($window in $windows) {
            try {
                if ($window.Document.Folder.Self.Path -eq $FolderPath) {
                    $first = $true

                    foreach ($filePath in $FilePaths) {
                        $fileName = Split-Path $filePath -Leaf
                        $item = $window.Document.Folder.ParseName($fileName)

                        if ($null -ne $item) {
                            if ($first) {
                                # 29 = select + focus + ensure visible + deselect others
                                $window.Document.SelectItem($item, 29)
                                $first = $false
                            }
                            else {
                                # 1 = add to current selection
                                $window.Document.SelectItem($item, 1)
                            }
                        }
                    }

                    break
                }
            }
            catch {
                continue
            }
        }
    }
    catch {
        Write-Host "Не удалось автоматически выделить файлы в Проводнике."
        Write-Host "Папка задачи открыта, можно выделить файлы вручную."
    }
}

Write-Host "Скрипт запущен."
Write-Host "Папка загрузок: $Downloads"
Write-Host "Папка задач: $BaseDir"
Write-Host "Ожидаю новый ZIP-архив..."

while ($true) {
    $processedArchives = Get-Content $HistoryFile -ErrorAction SilentlyContinue

    $zip = Get-ChildItem -Path $Downloads -Filter "*.zip" -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Where-Object { $processedArchives -notcontains $_.FullName } |
        Select-Object -First 1

    if ($null -eq $zip) {
        Start-Sleep -Milliseconds $NoZipCheckDelayMs
        continue
    }

    Write-Host ""
    Write-Host "Найден архив: $($zip.Name)"
    Write-Host "Жду завершения загрузки..."

    Wait-ForFileReadyFast -Path $zip.FullName -DelayMilliseconds $DownloadCheckDelayMs

    $taskNumber = Get-NextTaskNumber -Directory $BaseDir
    $taskName = "Задача_$taskNumber"
    $taskDir = Join-Path $BaseDir $taskName

    New-Item -ItemType Directory -Path $taskDir | Out-Null
    Write-Host "Создана папка: $taskName"

    try {
        Expand-Archive -Path $zip.FullName -DestinationPath $taskDir -Force
        Write-Host "Архив распакован: $($zip.Name)"
    }
    catch {
        Write-Host "Ошибка при распаковке архива: $($zip.Name)"
        Write-Host $_.Exception.Message
        Add-Content -Path $HistoryFile -Value $zip.FullName
        continue
    }

    $docxFiles = Get-ChildItem -Path $taskDir -Filter "*.docx" -File -Recurse -ErrorAction SilentlyContinue
    $geojsonFiles = Get-ChildItem -Path $taskDir -Filter "*.geojson" -File -Recurse -ErrorAction SilentlyContinue

    Write-Host "Найдено GeoJSON-файлов: $($geojsonFiles.Count)"
    Write-Host "Найдено DOCX-файлов: $($docxFiles.Count)"

    # Выделяются только файлы, которые лежат прямо в основной папке задачи.
    $filesToSelect = Get-ChildItem -Path $taskDir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension.ToLower() -ne ".docx" } |
        Select-Object -ExpandProperty FullName

    if ($filesToSelect.Count -gt 0) {
        Select-FilesInExplorerFast -FolderPath $taskDir -FilePaths $filesToSelect
        Write-Host "Папка задачи открыта. Все файлы, кроме DOCX, должны быть выделены."
    }
    else {
        Start-Process explorer.exe -ArgumentList "`"$taskDir`""
        Write-Host "В основной папке задачи нет файлов для выделения, кроме DOCX."
    }

    $docx = $docxFiles | Select-Object -First 1

    if ($OpenDocxAutomatically -and $null -ne $docx) {
        Write-Host "Открываю DOCX: $($docx.Name)"
        Start-Process -FilePath $docx.FullName
    }
    elseif ($null -ne $docx) {
        Write-Host "DOCX найден, но автооткрытие отключено: $($docx.Name)"
    }
    else {
        Write-Host "DOCX-файл не найден."
    }

    try {
        Remove-Item -Path $zip.FullName -Force
        Write-Host "Архив удалён: $($zip.Name)"
    }
    catch {
        Write-Host "Не удалось удалить архив: $($zip.Name)"
        Write-Host $_.Exception.Message
    }

    Add-Content -Path $HistoryFile -Value $zip.FullName
    Write-Host "Архив добавлен в историю."
    Write-Host "Готово: $(Get-Date -Format 'HH:mm')"
    Write-Host ""
    Write-Host "Ожидаю следующий ZIP-архив..."
}
