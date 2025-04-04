import csv
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

def network_csv_to_testlink_csv(input_file, output_file):
    """
    ネットワークCSVファイルをTestLink用CSVファイルに変換する
    指定された列対応関係に基づいて変換:
    カテゴリ -> testsuite_name
    目的 -> testcase_name および summary
    手順+テストコマンド -> actions
    期待結果 -> expected_results
    """
    # TestLink用のCSVヘッダー
    headers = ['testsuite_name', 'testcase_name', 'summary', 'preconditions', 
               'execution_type', 'importance', 'step_number', 'actions', 'expected_results']
    
    # 入力CSVファイルを読み込む（Shift-JISエンコーディングを指定）
    rows = []
    try:
        with open(input_file, 'r', encoding='shift_jis', errors='replace') as f:
            reader = csv.reader(f)
            rows = list(reader)
    except Exception as e:
        print(f"ファイル読み込み中にエラーが発生しました: {str(e)}")
        raise
    
    if not rows:
        raise ValueError("CSVファイルが空または読み込めませんでした。")
    
    # ヘッダー行を取得
    header = rows[0]
    print(f"読み込まれたヘッダー: {header}")
    
    # ヘッダーの位置を検索
    category_col = header.index('カテゴリ') if 'カテゴリ' in header else 0
    purpose_col = header.index('目的') if '目的' in header else 1
    procedure_col = header.index('手順') if '手順' in header else 2
    expected_col = header.index('期待結果') if '期待結果' in header else 3
    command_col = header.index('テストコマンド') if 'テストコマンド' in header else 4
    
    print(f"カラムマッピング: カテゴリ={category_col}, 目的={purpose_col}, 手順={procedure_col}, 期待結果={expected_col}, テストコマンド={command_col}")
    
    # 出力CSVファイルを書き込む
    with open(output_file, 'w', newline='', encoding='shift_jis', errors='replace') as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        
        # データ行を処理
        for i in range(1, len(rows)):
            row = rows[i]
            if len(row) < 3:  # 最低限必要なデータがなければスキップ
                continue
            
            # 各フィールドを取得（インデックスエラーを防ぐ）
            category = row[category_col] if category_col < len(row) else ""
            purpose = row[purpose_col] if purpose_col < len(row) else ""
            procedure = row[procedure_col] if procedure_col < len(row) else ""
            expected = row[expected_col] if expected_col < len(row) else ""
            command = row[command_col] if command_col < len(row) else ""
            
            # 空の行はスキップ
            if not any([category, purpose, procedure, expected, command]):
                continue
                
            # 手順とテストコマンドを連結
            actions = ""
            if procedure:
                actions += f"{procedure}"
            if command:
                if actions:
                    actions += "\n"
                actions += f"テストコマンド: {command}"
            
            # TestLink用の行を作成（指定された対応関係を使用）
            writer.writerow({
                'testsuite_name': category,  # カテゴリをtestsuite_nameに
                'testcase_name': purpose,    # 目的をtestcase_nameに
                'summary': purpose,          # 目的をsummaryにも
                'preconditions': "",         # 空欄
                'execution_type': '1',       # 手動実行
                'importance': '2',           # 中程度の重要度
                'step_number': '1',          # ステップ番号は1
                'actions': actions,          # 手順+テストコマンド
                'expected_results': expected # 期待結果
            })
    
    return output_file

def main():
    # ルートウィンドウを作成して非表示にする
    root = tk.Tk()
    root.withdraw()
    
    # ファイル選択ダイアログを表示
    input_file = filedialog.askopenfilename(
        title="変換するCSVファイルを選択してください",
        filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
    )
    
    if not input_file:  # ユーザーがキャンセルした場合
        print("ファイル選択がキャンセルされました。")
        return
    
    # 出力CSVファイルのパスを作成（入力CSVと同じフォルダ）
    dir_name = os.path.dirname(input_file)
    file_name = os.path.basename(input_file)
    base_name = os.path.splitext(file_name)[0]
    output_file = os.path.join(dir_name, f"{base_name}_testlink.csv")
    
    try:
        # 変換実行
        result_file = network_csv_to_testlink_csv(input_file, output_file)
        print(f"変換完了: {input_file} -> {result_file}")
        
        # 完了メッセージダイアログを表示
        messagebox.showinfo("変換完了", f"TestLink形式CSVへの変換が完了しました。\n出力ファイル: {result_file}")
        
    except Exception as e:
        # エラーメッセージダイアログを表示
        messagebox.showerror("エラー", f"変換中にエラーが発生しました。\n{str(e)}")
        print(f"エラー: {str(e)}")

if __name__ == "__main__":
    main()