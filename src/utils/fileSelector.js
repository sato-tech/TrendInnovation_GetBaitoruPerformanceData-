import { exec } from 'child_process';
import { promisify } from 'util';
import { platform } from 'os';
import { existsSync, writeFileSync, unlinkSync } from 'fs';
import { join } from 'path';
import { tmpdir } from 'os';

const execAsync = promisify(exec);

/**
 * ファイル選択用のユーティリティクラス
 */
class FileSelector {
  /**
   * Windowsでファイル選択ダイアログを表示する
   * @returns {Promise<string>} 選択されたファイルのパス
   */
  async selectFileWindows() {
    const powershellScript = `Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*"
$dialog.Title = "Excelファイルを選択してください"
$dialog.Multiselect = $false

if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    Write-Output $dialog.FileName
} else {
    Write-Error "ファイルが選択されませんでした"
    exit 1
}`;

    const tempScriptPath = join(tmpdir(), `file-dialog-${Date.now()}.ps1`);
    
    try {
      // 一時ファイルにPowerShellスクリプトを書き込む
      writeFileSync(tempScriptPath, powershellScript, 'utf8');

      // PowerShellスクリプトを実行
      const { stdout, stderr } = await execAsync(
        `powershell -ExecutionPolicy Bypass -File "${tempScriptPath}"`,
        { encoding: 'utf8', maxBuffer: 10 * 1024 * 1024 }
      );

      // 一時ファイルを削除
      try {
        unlinkSync(tempScriptPath);
      } catch (unlinkError) {
        // 削除に失敗しても続行
        console.warn(`一時ファイルの削除に失敗しました: ${unlinkError.message}`);
      }

      if (stderr && (stderr.includes('Write-Error') || stderr.includes('ファイルが選択されませんでした'))) {
        throw new Error('ファイルが選択されませんでした');
      }

      const filePath = stdout.trim();
      
      if (!filePath) {
        throw new Error('ファイルが選択されませんでした');
      }

      if (existsSync(filePath)) {
        return filePath;
      } else {
        throw new Error(`ファイルが見つかりません: ${filePath}`);
      }
    } catch (error) {
      // エラー時も一時ファイルを削除
      try {
        if (existsSync(tempScriptPath)) {
          unlinkSync(tempScriptPath);
        }
      } catch (unlinkError) {
        // 削除に失敗しても続行
      }

      if (error.message.includes('ファイルが選択されませんでした') || 
          error.message.includes('ファイルが見つかりません')) {
        throw error;
      }
      throw new Error(`ファイル選択ダイアログの表示に失敗しました: ${error.message}`);
    }
  }

  /**
   * macOSでファイル選択ダイアログを表示する
   * @returns {Promise<string>} 選択されたファイルのパス
   */
  async selectFileMacOS() {
    // AppleScriptを使用してファイル選択ダイアログを表示
    // ダイアログが確実に閉じるように、シンプルなAppleScriptを使用
    const appleScript = `set theFile to choose file with prompt "Excelファイルを選択してください" of type {"xlsx", "xls", "public.spreadsheet"} default location (path to desktop)
POSIX path of theFile`;

    const tempScriptPath = join(tmpdir(), `file-dialog-${Date.now()}.scpt`);
    
    try {
      // 一時ファイルにAppleScriptを書き込む
      writeFileSync(tempScriptPath, appleScript, 'utf8');

      // osascriptコマンドでAppleScriptを実行
      // 実行後、確実にプロセスが終了するようにする
      const { stdout, stderr } = await execAsync(
        `osascript "${tempScriptPath}" && exit 0`,
        { encoding: 'utf8', maxBuffer: 10 * 1024 * 1024, timeout: 300000 }
      );

      // 一時ファイルを削除
      try {
        unlinkSync(tempScriptPath);
      } catch (unlinkError) {
        // 削除に失敗しても続行
        console.warn(`一時ファイルの削除に失敗しました: ${unlinkError.message}`);
      }

      if (stderr && stderr.trim()) {
        // ユーザーがキャンセルした場合など
        if (stderr.includes('User cancelled') || stderr.includes('cancel') || stderr.includes('128')) {
          throw new Error('ファイルが選択されませんでした');
        }
        console.warn(`警告: ${stderr}`);
      }

      const filePath = stdout.trim();
      
      if (!filePath) {
        throw new Error('ファイルが選択されませんでした');
      }

      if (existsSync(filePath)) {
        return filePath;
      } else {
        throw new Error(`ファイルが見つかりません: ${filePath}`);
      }
    } catch (error) {
      // エラー時も一時ファイルを削除
      try {
        if (existsSync(tempScriptPath)) {
          unlinkSync(tempScriptPath);
        }
      } catch (unlinkError) {
        // 削除に失敗しても続行
      }

      if (error.message.includes('ファイルが選択されませんでした') || 
          error.message.includes('ファイルが見つかりません')) {
        throw error;
      }
      // osascriptのエラー（ユーザーがキャンセルした場合など）
      if (error.code === 1 || error.message.includes('exit code 1') || error.message.includes('User cancelled')) {
        throw new Error('ファイルが選択されませんでした');
      }
      throw new Error(`ファイル選択ダイアログの表示に失敗しました: ${error.message}`);
    }
  }

  /**
   * Linuxでファイル選択ダイアログを表示する（フォールバック）
   * @returns {Promise<string>} 選択されたファイルのパス
   */
  async selectFileLinux() {
    // Linuxでは、zenityやkdialogなどのツールを使用
    // ここでは簡易的な実装として、エラーメッセージを表示
    throw new Error('この機能はWindows/macOS環境でのみ利用可能です。ファイルパスを直接入力してください。');
  }

  /**
   * macOS/Linuxでファイル選択ダイアログを表示する
   * @returns {Promise<string>} 選択されたファイルのパス
   */
  async selectFileUnix() {
    const osPlatform = platform();
    
    if (osPlatform === 'darwin') {
      // macOS
      console.log('\n=== ファイル選択ダイアログを開いています... ===\n');
      return await this.selectFileMacOS();
    } else {
      // Linux
      return await this.selectFileLinux();
    }
  }

  /**
   * ユーザーにファイルを選択してもらう（プラットフォームに応じて適切な方法を使用）
   * @returns {Promise<string>} 選択されたファイルのパス
   */
  async selectFile() {
    const osPlatform = platform();
    
    if (osPlatform === 'win32') {
      console.log('\n=== ファイル選択ダイアログを開いています... ===\n');
      return await this.selectFileWindows();
    } else {
      // macOS/Linuxの場合はフォールバック
      return await this.selectFileUnix();
    }
  }
}

export default FileSelector;
