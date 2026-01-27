# .envファイルを作成または更新するスクリプト（Windows用）
# 文字エンコーディング: UTF-8 with BOM

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ENV_FILE = ".env"

# .envファイルが存在しない場合は作成
if (-not (Test-Path $ENV_FILE)) {
    Write-Host ".envファイルを作成します..."
    New-Item -ItemType File -Path $ENV_FILE -Force | Out-Null
}

# ファイルの内容を読み込む（UTF-8として）
try {
    $content = [System.IO.File]::ReadAllText($ENV_FILE, [System.Text.Encoding]::UTF8)
} catch {
    $content = ""
}

# 空の場合は空文字列に設定
if ([string]::IsNullOrEmpty($content)) {
    $content = ""
}

# Google Sheets API設定を追加または更新
if ($content -match "GOOGLE_SERVICE_ACCOUNT_KEY_PATH") {
    # 既存の設定を更新
    $content = $content -replace "GOOGLE_SERVICE_ACCOUNT_KEY_PATH=.*", "GOOGLE_SERVICE_ACCOUNT_KEY_PATH=./credentials.json"
    Write-Host "GOOGLE_SERVICE_ACCOUNT_KEY_PATHを更新しました"
} else {
    # 新しい設定を追加
    if ($content.Length -gt 0 -and -not $content.EndsWith("`r`n") -and -not $content.EndsWith("`n")) {
        $content += "`r`n"
    }
    if ($content.Length -gt 0) {
        $content += "`r`n"
    }
    $content += "# Google Sheets API設定`r`n"
    $content += "GOOGLE_SERVICE_ACCOUNT_KEY_PATH=./credentials.json`r`n"
    Write-Host "GOOGLE_SERVICE_ACCOUNT_KEY_PATHを追加しました"
}

if ($content -match "GOOGLE_SPREADSHEET_ID_NIGHT") {
    $content = $content -replace "GOOGLE_SPREADSHEET_ID_NIGHT=.*", "GOOGLE_SPREADSHEET_ID_NIGHT=1YtdXP1SSsHQgyw89b1mkRZdAmQJ0b7y4njMBQzjJDCw"
    Write-Host "GOOGLE_SPREADSHEET_ID_NIGHTを更新しました"
} else {
    if ($content.Length -gt 0 -and -not $content.EndsWith("`r`n") -and -not $content.EndsWith("`n")) {
        $content += "`r`n"
    }
    $content += "GOOGLE_SPREADSHEET_ID_NIGHT=1YtdXP1SSsHQgyw89b1mkRZdAmQJ0b7y4njMBQzjJDCw`r`n"
    Write-Host "GOOGLE_SPREADSHEET_ID_NIGHTを追加しました"
}

if ($content -match "GOOGLE_SPREADSHEET_ID_NORMAL") {
    $content = $content -replace "GOOGLE_SPREADSHEET_ID_NORMAL=.*", "GOOGLE_SPREADSHEET_ID_NORMAL=1S2jxEzLU57NznKALh3YAZme97z6AhiT8dkPS-Ofc9q4"
    Write-Host "GOOGLE_SPREADSHEET_ID_NORMALを更新しました"
} else {
    if ($content.Length -gt 0 -and -not $content.EndsWith("`r`n") -and -not $content.EndsWith("`n")) {
        $content += "`r`n"
    }
    $content += "GOOGLE_SPREADSHEET_ID_NORMAL=1S2jxEzLU57NznKALh3YAZme97z6AhiT8dkPS-Ofc9q4`r`n"
    Write-Host "GOOGLE_SPREADSHEET_ID_NORMALを追加しました"
}

if ($content -match "GOOGLE_SHEET_NAME") {
    $content = $content -replace "GOOGLE_SHEET_NAME=.*", "GOOGLE_SHEET_NAME=Sheet1"
    Write-Host "GOOGLE_SHEET_NAMEを更新しました"
} else {
    if ($content.Length -gt 0 -and -not $content.EndsWith("`r`n") -and -not $content.EndsWith("`n")) {
        $content += "`r`n"
    }
    $content += "GOOGLE_SHEET_NAME=Sheet1`r`n"
    Write-Host "GOOGLE_SHEET_NAMEを追加しました"
}

# ファイルに書き込む（UTF-8 without BOM）
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($ENV_FILE, $content, $utf8NoBom)

Write-Host ""
Write-Host "✓ .envファイルの設定が完了しました"
Write-Host ""
Write-Host "現在の設定:"
try {
    $fileContent = [System.IO.File]::ReadAllLines($ENV_FILE, [System.Text.Encoding]::UTF8)
    $googleSettings = $fileContent | Where-Object { $_ -match "GOOGLE_" }
    if ($googleSettings) {
        $googleSettings | ForEach-Object { Write-Host $_ }
    } else {
        Write-Host "(設定が見つかりませんでした)"
    }
} catch {
    Write-Host "(設定の読み込み中にエラーが発生しました: $($_.Exception.Message))"
}
