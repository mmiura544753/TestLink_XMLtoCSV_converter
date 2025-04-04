import csv
import xml.dom.minidom
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re
import html
import codecs

def create_cdata_section(doc, text):
    """指定されたテキストでCDATAセクションを作成する"""
    if text is None or text == "":
        return doc.createCDATASection("")
    return doc.createCDATASection(text)

def wrap_in_html_tags(text, tag_type="p"):
    """テキストをHTMLタグで囲む"""
    if text is None or text == "":
        return ""
    # すでにHTMLタグがある場合はそのまま返す
    if re.search(r'<\w+>', text):
        return text
    return f"<{tag_type}>{text}</{tag_type}>\n"

def create_list_html(items):
    """リスト項目からHTMLのリストを作成する"""
    if not items:
        return ""
    
    html_list = "<ol>\n"
    for item in items:
        item = item.strip()
        if item:
            html_list += f"\t<li>{item}</li>\n"
    html_list += "</ol>\n"
    return html_list

def csv_to_xml(csv_file, output_file):
    """CSVファイルをTestLink XMLに変換する"""
    
    # CSVファイルを読み込む（Shift-JIS エンコード）
    rows = []
    with open(csv_file, 'r', encoding='shift_jis', errors='replace') as file:
        reader = csv.DictReader(file)
        # 列名を取得
        fieldnames = reader.fieldnames
        
        # 必須フィールドの確認
        required_fields = ['testsuite_name', 'testcase_name']
        for field in required_fields:
            if field not in fieldnames:
                raise ValueError(f"CSVファイルに必須フィールド '{field}' がありません")
        
        # オプションフィールドの確認と追加
        optional_fields = ['testsuite_id', 'testcase_internalid', 'testcase_node_order']
        
        # すべての行を読み込む
        for row in reader:
            rows.append(row)
    
    if not rows:
        raise ValueError("CSVファイルにデータがありません")
    
    # XML文書を作成
    doc = xml.dom.minidom.getDOMImplementation().createDocument(None, "testsuite", None)
    root = doc.documentElement
    
    # テストスイート名と ID を設定
    testsuite_name = rows[0]['testsuite_name']
    root.setAttribute("name", testsuite_name)
    
    # テストスイートIDが指定されている場合は使用、ない場合は252を使用（元のXMLと同じ値）
    testsuite_id = "252"  # デフォルト値
    if 'testsuite_id' in fieldnames and rows[0]['testsuite_id']:
        testsuite_id = rows[0]['testsuite_id']
    root.setAttribute("id", testsuite_id)
    
    # node_orderとdetailsを追加
    node_order = doc.createElement("node_order")
    node_order.appendChild(create_cdata_section(doc, "1"))
    root.appendChild(node_order)
    
    details = doc.createElement("details")
    details.appendChild(create_cdata_section(doc, ""))
    root.appendChild(details)
    
    # テストケースをグループ化
    testcases = {}
    for row in rows:
        testcase_name = row['testcase_name']
        if testcase_name not in testcases:
            testcases[testcase_name] = {
                'summary': row.get('summary', ''),
                'preconditions': row.get('preconditions', ''),
                'execution_type': row.get('execution_type', '1'),
                'importance': row.get('importance', '2'),
                'steps': []
            }
        
        # ステップ情報があれば追加
        step_number = row.get('step_number', '')
        actions = row.get('actions', '')
        expected_results = row.get('expected_results', '')
        
        if step_number:
            testcases[testcase_name]['steps'].append({
                'step_number': step_number,
                'actions': actions,
                'expected_results': expected_results
            })
    
    # テストケースをXMLに追加
    counter = 1
    for testcase_name, testcase_data in testcases.items():
        tc_element = doc.createElement("testcase")
        tc_element.setAttribute("name", testcase_name)
        
        # internalidを取得（CSVに存在する場合はその値を使用、ない場合はオリジナルの形式に似た値を使用）
        internalid = None
        for row in rows:
            if row['testcase_name'] == testcase_name and 'testcase_internalid' in row and row['testcase_internalid']:
                internalid = row['testcase_internalid']
                break
                
        # internalidがない場合は、オリジナルのパターンに近い値を生成
        if not internalid:
            # オリジナルでは不規則だが、カウンターの値を4倍して近似値を生成
            internalid = str(counter * 4 + 8)
            
        tc_element.setAttribute("internalid", internalid)
        
        # node_orderの値を取得（CSVに存在する場合）
        node_order_value = str(counter - 1)  # デフォルト値
        for row in rows:
            if row['testcase_name'] == testcase_name and 'testcase_node_order' in row and row['testcase_node_order']:
                node_order_value = row['testcase_node_order']
                break
                
        # 基本要素を追加
        node_order_elem = doc.createElement("node_order")
        node_order_elem.appendChild(create_cdata_section(doc, node_order_value))
        tc_element.appendChild(node_order_elem)
        
        externalid_elem = doc.createElement("externalid")
        externalid_elem.appendChild(create_cdata_section(doc, str(counter)))
        tc_element.appendChild(externalid_elem)
        
        version_elem = doc.createElement("version")
        version_elem.appendChild(create_cdata_section(doc, "1"))
        tc_element.appendChild(version_elem)
        
        # サマリーを追加
        summary_elem = doc.createElement("summary")
        summary_text = wrap_in_html_tags(testcase_data['summary'])
        summary_elem.appendChild(create_cdata_section(doc, summary_text))
        tc_element.appendChild(summary_elem)
        
        # 前提条件を追加
        preconditions_elem = doc.createElement("preconditions")
        preconditions_text = wrap_in_html_tags(testcase_data['preconditions'])
        preconditions_elem.appendChild(create_cdata_section(doc, preconditions_text))
        tc_element.appendChild(preconditions_elem)
        
        # 実行タイプを追加
        execution_type_elem = doc.createElement("execution_type")
        execution_type_elem.appendChild(create_cdata_section(doc, testcase_data['execution_type']))
        tc_element.appendChild(execution_type_elem)
        
        # 重要度を追加
        importance_elem = doc.createElement("importance")
        importance_elem.appendChild(create_cdata_section(doc, testcase_data['importance']))
        tc_element.appendChild(importance_elem)
        
        # 予定実行時間（オリジナルの形式を維持）
        estimated_exec_duration = doc.createElement("estimated_exec_duration")
        # オリジナルのXMLでは空の値もテキストノードとして追加されていた
        if 'estimated_exec_duration' in testcase_data and testcase_data['estimated_exec_duration']:
            estimated_exec_duration.appendChild(doc.createTextNode(testcase_data['estimated_exec_duration']))
        tc_element.appendChild(estimated_exec_duration)
        
        # ステータス関連の情報
        status_elem = doc.createElement("status")
        status_elem.appendChild(doc.createTextNode("1"))
        tc_element.appendChild(status_elem)
        
        is_open_elem = doc.createElement("is_open")
        is_open_elem.appendChild(doc.createTextNode("1"))
        tc_element.appendChild(is_open_elem)
        
        active_elem = doc.createElement("active")
        active_elem.appendChild(doc.createTextNode("1"))
        tc_element.appendChild(active_elem)
        
        # ステップ情報がある場合
        if testcase_data['steps']:
            steps_elem = doc.createElement("steps")
            
            for step_data in testcase_data['steps']:
                step_elem = doc.createElement("step")
                
                # ステップ番号
                step_number_elem = doc.createElement("step_number")
                step_number_elem.appendChild(create_cdata_section(doc, step_data['step_number']))
                step_elem.appendChild(step_number_elem)
                
                # アクション
                actions_elem = doc.createElement("actions")
                
                # アクションテキストの処理
                actions_text = step_data['actions']
                if actions_text:
                    # カンマで区切られたリストの場合、HTMLリストに変換
                    if "," in actions_text and "<" not in actions_text:
                        actions_items = [item.strip() for item in actions_text.split(",")]
                        actions_text = create_list_html(actions_items)
                    else:
                        # 通常のテキストの場合
                        if not re.search(r'<\w+>', actions_text):
                            actions_text = wrap_in_html_tags(actions_text)
                
                actions_elem.appendChild(create_cdata_section(doc, actions_text))
                step_elem.appendChild(actions_elem)
                
                # 期待結果
                expected_results_elem = doc.createElement("expectedresults")
                expected_results_text = wrap_in_html_tags(step_data['expected_results'])
                expected_results_elem.appendChild(create_cdata_section(doc, expected_results_text))
                step_elem.appendChild(expected_results_elem)
                
                # ステップ実行タイプ
                step_exec_type_elem = doc.createElement("execution_type")
                step_exec_type_elem.appendChild(create_cdata_section(doc, "1"))
                step_elem.appendChild(step_exec_type_elem)
                
                steps_elem.appendChild(step_elem)
            
            tc_element.appendChild(steps_elem)
        
        root.appendChild(tc_element)
        counter += 1
    
    # UTF-8でXMLファイルを書き出す（XML宣言を追加）
    xml_declaration = '<?xml version="1.0" encoding="UTF-8"?>\n'
    formatted_xml = doc.toprettyxml(indent="  ", encoding='utf-8').decode('utf-8')
    
    # XML宣言の重複を防ぐ
    if formatted_xml.startswith('<?xml'):
        formatted_xml_content = formatted_xml
    else:
        formatted_xml_content = xml_declaration + formatted_xml
    
    with codecs.open(output_file, 'w', encoding='utf-8') as f:
        f.write(formatted_xml_content)
    
    return output_file

def update_csv_template():
    """CSVテンプレートファイルを作成して開く"""
    template_file = "testlink_template.csv"
    
    # テンプレートの列名
    fieldnames = [
        'testsuite_name', 'testsuite_id', 
        'testcase_name', 'testcase_internalid', 'testcase_node_order',
        'summary', 'preconditions', 'execution_type', 'importance',
        'estimated_exec_duration', 'status', 'is_open', 'active',
        'step_number', 'actions', 'expected_results'
    ]
    
    # テンプレートCSVを作成
    with open(template_file, 'w', encoding='shift_jis', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        
        # サンプル行を追加
        writer.writerow({
            'testsuite_name': 'テストスイート名',
            'testsuite_id': '252',
            'testcase_name': 'テストケース名',
            'testcase_internalid': '12',
            'testcase_node_order': '0',
            'summary': 'テストケースの概要',
            'preconditions': '前提条件',
            'execution_type': '1',
            'importance': '2',
            'estimated_exec_duration': '5.00',
            'status': '1',
            'is_open': '1',
            'active': '1',
            'step_number': '1',
            'actions': 'アクション1, アクション2',
            'expected_results': '期待される結果'
        })
    
    # ファイルを開く
    try:
        os.startfile(template_file)
    except:
        try:
            import subprocess
            subprocess.call(['open', template_file])
        except:
            pass
    
    return template_file

def main():
    # ルートウィンドウを作成
    root = tk.Tk()
    root.title("TestLink CSV to XML コンバーター")
    root.geometry("400x200")
    
    def select_file():
        # ファイル選択ダイアログを表示
        csv_file = filedialog.askopenfilename(
            title="EXCELで作成したCSVファイルを選択してください",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        
        if not csv_file:  # ユーザーがキャンセルした場合
            return
        
        # 出力XMLファイルのパスを作成（入力CSVと同じフォルダ）
        dir_name = os.path.dirname(csv_file)
        file_name = os.path.basename(csv_file)
        base_name = os.path.splitext(file_name)[0]
        xml_file = os.path.join(dir_name, f"{base_name}.xml")
        
        try:
            # 変換実行
            result_file = csv_to_xml(csv_file, xml_file)
            # 完了メッセージダイアログを表示
            messagebox.showinfo("変換完了", f"CSVからXMLへの変換が完了しました。\n出力ファイル: {result_file}")
            
        except Exception as e:
            # エラーメッセージダイアログを表示
            messagebox.showerror("エラー", f"変換中にエラーが発生しました。\n{str(e)}")
            print(f"エラー: {str(e)}")
    
    def create_template():
        template_file = update_csv_template()
        messagebox.showinfo("テンプレート作成", f"CSVテンプレートファイルを作成しました。\n{template_file}")
    
    # フレームを作成
    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(expand=True, fill="both")
    
    # アプリの説明ラベル
    label = tk.Label(frame, text="TestLink用CSVをXMLに変換するツール", font=("Helvetica", 12))
    label.pack(pady=10)
    
    # 変換ボタン
    convert_button = tk.Button(frame, text="CSVファイルを選択して変換", command=select_file, width=25)
    convert_button.pack(pady=5)
    
    # テンプレート作成ボタン
    template_button = tk.Button(frame, text="CSVテンプレートを作成", command=create_template, width=25)
    template_button.pack(pady=5)
    
    # 終了ボタン
    quit_button = tk.Button(frame, text="終了", command=root.destroy, width=25)
    quit_button.pack(pady=5)
    
    # メインループ
    root.mainloop()
    
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