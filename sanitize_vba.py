import zipfile
import os
import shutil
import tempfile
import re
import sys
from pathlib import Path

# Windows console encoding fix
if sys.platform == 'win32':
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

def sanitize_vba_in_xlsm(file_path, search_text, replace_text):
    """
    .xlsmファイル内のVBAコードから指定文字列を置換する

    Args:
        file_path: 対象の.xlsmファイルパス
        search_text: 検索文字列
        replace_text: 置換文字列

    Returns:
        置換件数
    """
    print(f"\n{'='*60}")
    print(f"ファイル: {os.path.basename(file_path)}")
    print(f"検索文字列: '{search_text}' → 置換文字列: '{replace_text}'")
    print(f"{'='*60}\n")

    # バックアップを作成
    backup_path = file_path + ".backup"
    shutil.copy2(file_path, backup_path)
    print(f"[OK] バックアップ作成: {os.path.basename(backup_path)}")

    # 一時ディレクトリを作成
    temp_dir = tempfile.mkdtemp()

    try:
        # .xlsmをZIPとして展開
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        print(f"[OK] ファイル展開完了")

        # vbaProject.binを検索
        vba_project_path = os.path.join(temp_dir, 'xl', 'vbaProject.bin')

        if not os.path.exists(vba_project_path):
            print("[ERROR] VBAプロジェクトが見つかりません")
            return 0

        print(f"[OK] VBAプロジェクト検出: {vba_project_path}")

        # VBAプロジェクトを読み込み
        with open(vba_project_path, 'rb') as f:
            vba_content = f.read()

        original_size = len(vba_content)
        print(f"[OK] VBAプロジェクトサイズ: {original_size:,} bytes")

        # バイナリデータを文字列として扱う（latin-1エンコーディング）
        vba_text = vba_content.decode('latin-1', errors='ignore')

        # 置換前の出現回数をカウント
        count_before = vba_text.count(search_text)
        print(f"\n検索文字列 '{search_text}' の出現回数: {count_before}件")

        if count_before == 0:
            print("[OK] 置換対象が見つかりませんでした")
            return 0

        # 置換実行
        vba_text_replaced = vba_text.replace(search_text, replace_text)

        # 置換後の確認
        count_after = vba_text_replaced.count(search_text)
        replaced_count = count_before - count_after

        print(f"[OK] 置換完了: {replaced_count}件")
        print(f"  残存確認: {count_after}件")

        # バイナリに戻す
        vba_content_replaced = vba_text_replaced.encode('latin-1', errors='ignore')

        # 置換後のサイズ確認
        new_size = len(vba_content_replaced)
        size_diff = new_size - original_size
        print(f"  新サイズ: {new_size:,} bytes (差分: {size_diff:+,} bytes)")

        # VBAプロジェクトを書き戻す
        with open(vba_project_path, 'wb') as f:
            f.write(vba_content_replaced)

        # 新しい.xlsmファイルを作成
        output_path = file_path + ".temp"
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path_full = os.path.join(root, file)
                    arcname = os.path.relpath(file_path_full, temp_dir)
                    zip_out.write(file_path_full, arcname)

        print(f"[OK] 新ファイル作成完了")

        # 元ファイルを置き換え
        os.remove(file_path)
        os.rename(output_path, file_path)

        print(f"[OK] ファイル更新完了: {os.path.basename(file_path)}")

        return replaced_count

    except Exception as e:
        print(f"\n[ERROR] エラー発生: {e}")
        # エラー時はバックアップから復元
        if os.path.exists(backup_path):
            shutil.copy2(backup_path, file_path)
            print(f"[OK] バックアップから復元しました")
        return 0

    finally:
        # 一時ディレクトリを削除
        shutil.rmtree(temp_dir, ignore_errors=True)


def verify_replacement(file_path, search_text):
    """
    置換後の検証: 指定文字列が残っていないか確認

    Args:
        file_path: 対象の.xlsmファイルパス
        search_text: 検索文字列

    Returns:
        残存件数
    """
    print(f"\n{'='*60}")
    print(f"検証: {os.path.basename(file_path)}")
    print(f"検索文字列: '{search_text}'")
    print(f"{'='*60}\n")

    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            vba_project_path = 'xl/vbaProject.bin'

            if vba_project_path not in zip_ref.namelist():
                print("[ERROR] VBAプロジェクトが見つかりません")
                return -1

            vba_content = zip_ref.read(vba_project_path)
            vba_text = vba_content.decode('latin-1', errors='ignore')

            remaining_count = vba_text.count(search_text)

            if remaining_count == 0:
                print(f"[OK] 検証成功: '{search_text}' は検出されませんでした")
            else:
                print(f"[WARNING] 警告: '{search_text}' が {remaining_count}件 残っています")

            return remaining_count

    except Exception as e:
        print(f"[ERROR] 検証エラー: {e}")
        return -1


def main():
    print("\n" + "="*60)
    print("VBAコード サニタイゼーションツール")
    print("="*60)

    # 対象ファイル
    target_file = "基幹システム連携.xlsm"

    if not os.path.exists(target_file):
        print(f"\n[ERROR] エラー: {target_file} が見つかりません")
        return

    # 置換設定
    search_text = "QBIS"
    replace_text = "ERPSystem"

    # サニタイゼーション実行
    replaced_count = sanitize_vba_in_xlsm(target_file, search_text, replace_text)

    # 検証
    remaining_count = verify_replacement(target_file, search_text)

    # 結果サマリー
    print(f"\n{'='*60}")
    print("サニタイゼーション完了")
    print(f"{'='*60}")
    print(f"ファイル: {target_file}")
    print(f"置換件数: {replaced_count}件")
    print(f"残存確認: {remaining_count}件")
    print(f"バックアップ: {target_file}.backup")
    print(f"{'='*60}\n")

    if remaining_count == 0:
        print("[SUCCESS] すべての置換が成功しました！")
    else:
        print("[WARNING] 一部の文字列が残っています。手動確認が必要です。")


if __name__ == "__main__":
    main()
