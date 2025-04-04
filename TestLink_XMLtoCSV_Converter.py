import xml.etree.ElementTree as ET
import csv
import re
import html
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

def clean_html(text):
    """HTMLタグを削除し、プレーンテキストを抽出する"""
    if text is None:
        return ""
    # CDataセクションの処理
    text = text.replace("<![CDATA[", "").replace("]]>", "")
    # HTMLタグの削除（単純なアプローチ）
    text = re.sub(r'<[^>]+>', ' ', text)
    # HTMLエンティティをデコード
    text = html.unescape(text)
    # 連続する空白を1つにまとめる
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def extract_list_items(text):
    """HTML内のリスト項目を抽出する"""
    if text is None:
        return ""
    # リスト項目を抽出
    items = re.findall(r'<li>(.*?)</li>', text)
    if items:
        return ", ".join([item.strip() for item in items])
    else:
        return clean_html(text)

def xml_to_csv(xml_file, csv_file):
    """TestLink XMLファイルからCSVファイルを生成する"""
    tree = ET.parse(xml_file)
    root = tree.getroot()
    
    # CSVヘッダー（externalidを追加）
    headers = ['testsuite_name', 'testcase_name', 'externalid', 'summary', 'preconditions', 
               'execution_type', 'importance', 'step_number', 'actions', 'expected_results']
    
    # Shift-JISエンコードでCSVファイルを書き込み
    with open(csv_file, 'w', newline='', encoding='shift_jis', errors='replace') as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        
        # テストスイート名を取得
        testsuite_name = root.attrib.get('name', '')
        
        # テストケースを処理
        for testcase in root.findall('.//testcase'):
            testcase_name = testcase.attrib.get('name', '')
            
            # externalidを取得
            externalid = testcase.find('externalid').text.replace("<![CDATA[", "").replace("]]>", "") if testcase.find('externalid') is not None else ''
            
            # 基本情報を取得
            summary = clean_html(testcase.find('summary').text if testcase.find('summary') is not None else '')
            preconditions = clean_html(testcase.find('preconditions').text if testcase.find('preconditions') is not None else '')
            execution_type = testcase.find('execution_type').text.replace("<![CDATA[", "").replace("]]>", "") if testcase.find('execution_type') is not None else ''
            importance = testcase.find('importance').text.replace("<![CDATA[", "").replace("]]>", "") if testcase.find('importance') is not None else ''
            
            # ステップ情報を処理
            steps = testcase.find('steps')
            if steps is not None and len(steps) > 0:
                for step in steps.findall('step'):
                    step_number = step.find('step_number').text.replace("<![CDATA[", "").replace("]]>", "") if step.find('step_number') is not None else ''
                    actions = extract_list_items(step.find('actions').text if step.find('actions') is not None else '')
                    expected_results = clean_html(step.find('expectedresults').text if step.find('expectedresults') is not None else '')
                    
                    # 行を書き込み
                    writer.writerow({
                        'testsuite_name': testsuite_name,
                        'testcase_name': testcase_name,
                        'externalid': externalid,
                        'summary': summary,
                        'preconditions': preconditions,
                        'execution_type': execution_type,
                        'importance': importance,
                        'step_number': step_number,
                        'actions': actions,
                        'expected_results': expected_results
                    })
            else:
                # ステップがない場合は空の行を作成
                writer.writerow({
                    'testsuite_name': testsuite_name,
                    'testcase_name': testcase_name,
                    'externalid': externalid,
                    'summary': summary,
                    'preconditions': preconditions,
                    'execution_type': execution_type,
                    'importance': importance,
                    'step_number': '',
                    'actions': '',
                    'expected_results': ''
                })
    
    return csv_file

def main():
    # ルートウィンドウを作成して非表示にする
    root = tk.Tk()
    root.withdraw()
    
    # ファイル選択ダイアログを表示
    xml_file = filedialog.askopenfilename(
        title="TestLinkのXMLファイルを選択してください",
        filetypes=[("XML Files", "*.xml"), ("All Files", "*.*")]
    )
    
    if not xml_file:  # ユーザーがキャンセルした場合
        print("ファイル選択がキャンセルされました。")
        return
    
    # 出力CSVファイルのパスを作成（入力XMLと同じフォルダ）
    dir_name = os.path.dirname(xml_file)
    file_name = os.path.basename(xml_file)
    base_name = os.path.splitext(file_name)[0]
    csv_file = os.path.join(dir_name, f"{base_name}.csv")
    
    try:
        # 変換実行
        result_file = xml_to_csv(xml_file, csv_file)
        print(f"変換完了: {xml_file} -> {result_file}")
        
        # 完了メッセージダイアログを表示
        messagebox.showinfo("変換完了", f"XMLからCSVへの変換が完了しました。\n出力ファイル: {result_file}")
        
    except Exception as e:
        # エラーメッセージダイアログを表示
        messagebox.showerror("エラー", f"変換中にエラーが発生しました。\n{str(e)}")
        print(f"エラー: {str(e)}")

if __name__ == "__main__":
    main()