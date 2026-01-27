#!/bin/bash

# .envファイルを作成または更新するスクリプト

ENV_FILE=".env"

# .envファイルが存在しない場合は作成
if [ ! -f "$ENV_FILE" ]; then
    echo ".envファイルを作成します..."
    touch "$ENV_FILE"
fi

# Google Sheets API設定を追加または更新
if grep -q "GOOGLE_SERVICE_ACCOUNT_KEY_PATH" "$ENV_FILE"; then
    # 既存の設定を更新
    sed -i '' 's|GOOGLE_SERVICE_ACCOUNT_KEY_PATH=.*|GOOGLE_SERVICE_ACCOUNT_KEY_PATH=./credentials.json|' "$ENV_FILE"
    echo "GOOGLE_SERVICE_ACCOUNT_KEY_PATHを更新しました"
else
    # 新しい設定を追加
    echo "" >> "$ENV_FILE"
    echo "# Google Sheets API設定" >> "$ENV_FILE"
    echo "GOOGLE_SERVICE_ACCOUNT_KEY_PATH=./credentials.json" >> "$ENV_FILE"
    echo "GOOGLE_SERVICE_ACCOUNT_KEY_PATHを追加しました"
fi

if grep -q "GOOGLE_SPREADSHEET_ID_NIGHT" "$ENV_FILE"; then
    sed -i '' 's|GOOGLE_SPREADSHEET_ID_NIGHT=.*|GOOGLE_SPREADSHEET_ID_NIGHT=1YtdXP1SSsHQgyw89b1mkRZdAmQJ0b7y4njMBQzjJDCw|' "$ENV_FILE"
    echo "GOOGLE_SPREADSHEET_ID_NIGHTを更新しました"
else
    echo "GOOGLE_SPREADSHEET_ID_NIGHT=1YtdXP1SSsHQgyw89b1mkRZdAmQJ0b7y4njMBQzjJDCw" >> "$ENV_FILE"
    echo "GOOGLE_SPREADSHEET_ID_NIGHTを追加しました"
fi

if grep -q "GOOGLE_SPREADSHEET_ID_NORMAL" "$ENV_FILE"; then
    sed -i '' 's|GOOGLE_SPREADSHEET_ID_NORMAL=.*|GOOGLE_SPREADSHEET_ID_NORMAL=1S2jxEzLU57NznKALh3YAZme97z6AhiT8dkPS-Ofc9q4|' "$ENV_FILE"
    echo "GOOGLE_SPREADSHEET_ID_NORMALを更新しました"
else
    echo "GOOGLE_SPREADSHEET_ID_NORMAL=1S2jxEzLU57NznKALh3YAZme97z6AhiT8dkPS-Ofc9q4" >> "$ENV_FILE"
    echo "GOOGLE_SPREADSHEET_ID_NORMALを追加しました"
fi

if grep -q "GOOGLE_SHEET_NAME" "$ENV_FILE"; then
    sed -i '' 's|GOOGLE_SHEET_NAME=.*|GOOGLE_SHEET_NAME=Sheet1|' "$ENV_FILE"
    echo "GOOGLE_SHEET_NAMEを更新しました"
else
    echo "GOOGLE_SHEET_NAME=Sheet1" >> "$ENV_FILE"
    echo "GOOGLE_SHEET_NAMEを追加しました"
fi

echo ""
echo "✓ .envファイルの設定が完了しました"
echo ""
echo "現在の設定:"
grep "GOOGLE_" "$ENV_FILE" || echo "（設定が見つかりませんでした）"
