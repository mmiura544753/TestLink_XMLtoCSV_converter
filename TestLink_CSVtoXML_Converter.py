import csv
import os
import re
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from xml.dom import minidom
from xml.etree import ElementTree as ET

def split_numbered_steps(actions_text):
    """番号付きリスト形式のステップを分割する"""
    if not actions_text:
        return []
    
    # 「1. 〜」「2. 〜」のようなパターンを検出
    pattern = r'(\d+\.\s*[^0-9.]+?)(?=\d+\.|$)'
    steps = re.findall(pattern, actions_text)
    
    if not steps:
        # 番号付きリストが見つからない場合は元のテキストを1つのステップとして返す
        return [actions_text.strip()]
    
    return [step.strip() for step in steps]

def add_cdata_section(parent, tag_name, text):
    """CDATA セクションを含む要素を追加する"""
    element = ET.SubElement(parent, tag_name)
    element.text = f"<![CDATA[{text}]]>"
    return element

def csv_to_xml(csv_file, xml_file):
    """TestLink CSV ファイルから XML ファイルを生成する"""
    # CSVファイルを読み込む
    test_cases = {}
    test_suites = {}
    
    with open(csv_file, 'r', encoding='shift_jis', errors='replace') as file:
        reader = csv.DictReader(file)
        headers = reader.fieldnames
        
        # externalidがヘッダーに存在するか確認
        has_externalid = 'externalid' in headers
        
        for row in reader:
            testsuite_name = row.get('testsuite_name', '')
            testcase_name = row.get('testcase_name', '')
            
            # テストスイートの処理
            if testsuite_name not in test_suites:
                test_suites[testsuite_name] = True
            
            # テストケースの処理
            case_key = f"{testsuite_name}_{testcase_name}"
            if case_key not in test_cases:
                test_cases[case_key] = {
                    'name': testcase_name,
                    'testsuite': testsuite_name,
                    'summary': row.get('summary', ''),
                    'preconditions': row.get('preconditions', ''),
                    'execution_type': row.get('execution_type', '1'),
                    'importance': row.get('importance', '2'),
                    'externalid': row.get('externalid', '') if has_externalid else '',
                    'steps': []
                }
            
            # ステップの処理
            if row.get('actions', ''):
                actions = row.get('actions', '')
                expected_results = row.get('expected_results', '')
                
                # 番号付きステップを分割
                action_steps = split_numbered_steps(actions)
                
                # 各ステップをテストケースに追加
                for i, step_action in enumerate(action_steps):
                    # 最後のステップにのみ期待結果を設定
                    step_expected = expected_results if i == len(action_steps) - 1 else ''
                    
                    test_cases[case_key]['steps'].append({
                        'number': i + 1,
                        'actions': step_action,
                        'expected_results': step_expected,
                        'execution_type': '1'
                    })
    
    # XMLを生成
    # ルート要素（テストスイート）を作成
    testsuite_name = next(iter(test_suites))
    root = ET.Element('testsuite', name=testsuite_name)
    
    # ノード順序の要素を追加
    node_order = ET.SubElement(root, 'node_order')
    node_order.text = "<![CDATA[1]]>"
    
    # 詳細の要素を追加
    details = ET.SubElement(root, 'details')
    details.text = "<![CDATA[]]>"
    
    # テストケースを追加
    for key, testcase_data in test_cases.items():
        # テストケース要素
        testcase = ET.SubElement(root, 'testcase', name=testcase_data['name'], internalid="1")
        
        # ノード順序
        node_order = ET.SubElement(testcase, 'node_order')
        node_order.text = "<![CDATA[0]]>"
        
        # 外部ID（存在する場合のみ追加）
        if has_externalid and testcase_data['externalid']:
            externalid = ET.SubElement(testcase, 'externalid')
            externalid.text = f"<![CDATA[{testcase_data['externalid']}]]>"
        
        # バージョン
        version = ET.SubElement(testcase, 'version')
        version.text = "<![CDATA[1]]>"
        
        # 概要
        summary = ET.SubElement(testcase, 'summary')
        summary.text = f"<![CDATA[<p>{testcase_data['summary']}</p>\n]]>"
        
        # 前提条件
        preconditions = ET.SubElement(testcase, 'preconditions')
        preconditions.text = f"<![CDATA[{testcase_data['preconditions']}]]>"
        
        # 実行タイプ
        execution_type = ET.SubElement(testcase, 'execution_type')
        execution_type.text = f"<![CDATA[{testcase_data['execution_type']}]]>"
        
        # 重要度
        importance = ET.SubElement(testcase, 'importance')
        importance.text = f"<![CDATA[{testcase_data['importance']}]]>"
        
        # ステータス関連の標準フィールド
        ET.SubElement(testcase, 'estimated_exec_duration').text = "5.00"
        ET.SubElement(testcase, 'status').text = "1"
        ET.SubElement(testcase, 'is_open').text = "1"
        ET.SubElement(testcase, 'active').text = "1"
        
        # ステップを追加（ステップがある場合）
        if testcase_data['steps']:
            steps = ET.SubElement(testcase, 'steps')
            
            for step_data in testcase_data['steps']:
                step = ET.SubElement(steps, 'step')
                
                step_number = ET.SubElement(step, 'step_number')
                step_number.text = f"<![CDATA[{step_data['number']}]]>"
                
                actions = ET.SubElement(step, 'actions')
                actions.text = f"<![CDATA[<p>{step_data['actions']}</p>\n]]>"
                
                expectedresults = ET.SubElement(step, 'expectedresults')
                expectedresults.text = f"<![CDATA[<p>{step_data['expected_results']}</p>\n]]>" if step_data['expected_results'] else "<![CDATA[]]>"
                
                step_execution_type = ET.SubElement(step, 'execution_type')
                step_execution_type.text = f"<![CDATA[{step_data['execution_type']}]]>"
    
    # XMLファイルを保存
    # きれいな形式にインデントする
    rough_string = ET.tostring(root, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    pretty_xml = reparsed.toprettyxml(indent="  ", encoding="utf-8").decode('utf-8')
    
    with open(xml_file, 'w', encoding='utf-8') as f:
        f.write('<?xml version="1.0" encoding="utf-8"?>\n')
        # XMLの宣言行を除去して書き込む（重複を避けるため）
        f.write(pretty_xml.split('\n', 1)[1])
    
    return xml_file

def main():
    # ルートウィンドウを作成して非表示にする
    root = tk.Tk()
    root.withdraw()
    
    # ファイル選択ダイアログを表示
    csv_file = filedialog.askopenfilename(
        title="TestLinkのCSVファイルを選択してください",
        filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
    )
    
    if not csv_file:  # ユーザーがキャンセルした場合
        print("ファイル選択がキャンセルされました。")
        return
    
    # 出力XMLファイルのパスを作成（入力CSVと同じフォルダ）
    dir_name = os.path.dirname(csv_file)
    file_name = os.path.basename(csv_file)
    base_name = os.path.splitext(file_name)[0]
    xml_file = os.path.join(dir_name, f"{base_name}.xml")
    
    try:
        # 変換実行
        result_file = csv_to_xml(csv_file, xml_file)
        print(f"変換完了: {csv_file} -> {result_file}")
        
        # 完了メッセージダイアログを表示
        messagebox.showinfo("変換完了", f"CSVからXMLへの変換が完了しました。\n出力ファイル: {result_file}")
        
    except Exception as e:
        # エラーメッセージダイアログを表示
        messagebox.showerror("エラー", f"変換中にエラーが発生しました。\n{str(e)}")
        print(f"エラー: {str(e)}")

if __name__ == "__main__":
    main()