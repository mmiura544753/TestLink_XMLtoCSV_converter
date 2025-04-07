import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import xml.etree.ElementTree as ET
import csv
import re
import codecs
import traceback

class TestLinkConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("TestLink XML-CSV Converter")
        self.root.geometry("400x200")
        self.root.resizable(False, False)
        
        # GUI要素の作成
        self.create_widgets()
        
        # ステータス表示の初期化
        self.update_status("待機中")
        
    def create_widgets(self):
        # XMLからCSV変換ボタン
        self.btn_xml_to_csv = tk.Button(self.root, text="XML→CSV変換", width=20, height=2,
                                        command=self.xml_to_csv)
        self.btn_xml_to_csv.pack(pady=10)
        
        # CSVからXML変換ボタン
        self.btn_csv_to_xml = tk.Button(self.root, text="CSV→XML変換", width=20, height=2,
                                        command=self.csv_to_xml)
        self.btn_csv_to_xml.pack(pady=10)
        
        # 終了ボタン
        self.btn_exit = tk.Button(self.root, text="終了", width=20, height=2,
                                  command=self.root.destroy)
        self.btn_exit.pack(pady=10)
        
        # ステータス表示ラベル
        self.lbl_status = tk.Label(self.root, text="ステータス: ", anchor="w")
        self.lbl_status.pack(fill=tk.X, padx=10, pady=10)
        
    def update_status(self, message):
        """ステータスメッセージを更新する"""
        self.lbl_status.config(text=f"ステータス: {message}")
        self.root.update()
        
    def xml_to_csv(self):
        """XMLファイルをCSVに変換する"""
        # XMLファイルの選択
        xml_file = filedialog.askopenfilename(
            title="XMLファイルを選択してください",
            filetypes=[("XMLファイル", "*.xml"), ("すべてのファイル", "*.*")]
        )
        
        if not xml_file:
            self.update_status("ファイルが選択されていません")
            return
            
        try:
            self.update_status("XMLファイルを解析中...")
            
            # XML内容の修正（二重CDATAなどの問題を解決）
            with open(xml_file, 'r', encoding='utf-8') as f:
                xml_content = f.read()
                            
            # 二重CDATA問題を修正
            xml_content = self.fix_double_cdata(xml_content)
            
            # 修正したXMLをパース
            root = ET.fromstring(xml_content)
            
            # 出力CSVファイル名の生成
            output_file = os.path.splitext(xml_file)[0] + ".csv"
            
            # データの変換と出力
            self.convert_xml_to_csv(root, output_file)
            
            self.update_status(f"変換完了: {output_file}")
            messagebox.showinfo("変換完了", f"CSVファイルに変換しました:\n{output_file}")
            
        except Exception as e:
            error_details = traceback.format_exc()
            self.update_status(f"エラー: {str(e)}")
            messagebox.showerror("エラー", f"変換中にエラーが発生しました:\n{str(e)}\n\n詳細:\n{error_details}")
    
    def csv_to_xml(self):
        """CSVファイルをXMLに変換する"""
        # CSVファイルの選択
        csv_file = filedialog.askopenfilename(
            title="CSVファイルを選択してください",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
        )
        
        if not csv_file:
            self.update_status("ファイルが選択されていません")
            return
            
        try:
            self.update_status("CSVファイルを解析中...")
            
            # 出力XMLファイル名の生成
            output_file = os.path.splitext(csv_file)[0] + "_converted.xml"
            
            # データの変換と出力
            self.convert_csv_to_xml(csv_file, output_file)
            
            self.update_status(f"変換完了: {output_file}")
            messagebox.showinfo("変換完了", f"XMLファイルに変換しました:\n{output_file}")
            
        except Exception as e:
            self.update_status(f"エラー: {str(e)}")
            messagebox.showerror("エラー", f"変換中にエラーが発生しました:\n{str(e)}")
    
    def convert_xml_to_csv(self, root, output_file):
        """XMLデータをCSVに変換する"""
        # CSVに書き込むデータのリスト
        rows = []
        
        # ヘッダー行
        headers = [
            "ID", "外部ID", "バージョン", "テストケース名", "サマリ（概要）", 
            "重要度", "事前条件", "ステップ番号", "アクション（手順）", "期待結果",
            "実行タイプ", "推定実行時間", "ステータス", "有効/無効", "開いているか",
            "親テストスイート名"
        ]
        rows.append(headers)
        
        # テストスイート名を取得
        testsuite_name = root.get("name", "")
        
        # 各テストケースを処理
        for testcase in root.findall(".//testcase"):
            testcase_id = self.get_element_text(testcase, "internalid")
            external_id = self.get_element_text(testcase, "externalid")
            version = self.get_element_text(testcase, "version")
            testcase_name = testcase.get("name", "")
            summary = self.clean_html(self.get_element_text(testcase, "summary"))
            importance = self.get_element_text(testcase, "importance")
            preconditions = self.clean_html(self.get_element_text(testcase, "preconditions"))
            exec_type = self.get_element_text(testcase, "execution_type")
            exec_duration = self.get_element_text(testcase, "estimated_exec_duration")
            status = self.get_element_text(testcase, "status")
            is_active = self.get_element_text(testcase, "active")
            is_open = self.get_element_text(testcase, "is_open")
            
            # ステップを処理
            steps = testcase.find("steps")
            if steps is not None and len(steps) > 0:
                for step in steps.findall("step"):
                    step_number = self.get_element_text(step, "step_number")
                    actions = self.clean_html(self.get_element_text(step, "actions"))
                    expected = self.clean_html(self.get_element_text(step, "expectedresults"))
                    
                    # 行を追加
                    row = [
                        testcase_id, external_id, version, testcase_name, summary,
                        importance, preconditions, step_number, actions, expected,
                        exec_type, exec_duration, status, is_active, is_open,
                        testsuite_name
                    ]
                    rows.append(row)
            else:
                # ステップがない場合は空行を追加
                row = [
                    testcase_id, external_id, version, testcase_name, summary,
                    importance, preconditions, "", "", "",
                    exec_type, exec_duration, status, is_active, is_open,
                    testsuite_name
                ]
                rows.append(row)
        
        # CSVファイルへの書き込み
        with codecs.open(output_file, 'w', 'shift_jis') as f:
            writer = csv.writer(f, quoting=csv.QUOTE_ALL)
            writer.writerows(rows)
    
    def convert_csv_to_xml(self, csv_file, output_file):
        """CSVデータをXMLに変換する"""
        try:
            # CSVファイルの読み込み
            rows = []
            with codecs.open(csv_file, 'r', 'shift_jis') as f:
                reader = csv.reader(f)
                for row in reader:
                    rows.append(row)
            
            if len(rows) < 2:
                raise ValueError("CSVファイルにデータがありません")
            
            # ヘッダー行
            headers = rows[0]
            
            # 必要なカラムのインデックスを取得
            try:
                id_idx = headers.index("ID")
                external_id_idx = headers.index("外部ID")
                version_idx = headers.index("バージョン")
                testcase_name_idx = headers.index("テストケース名")
                summary_idx = headers.index("サマリ（概要）")
                importance_idx = headers.index("重要度")
                preconditions_idx = headers.index("事前条件")
                step_number_idx = headers.index("ステップ番号")
                actions_idx = headers.index("アクション（手順）")
                expected_idx = headers.index("期待結果")
                exec_type_idx = headers.index("実行タイプ")
                exec_duration_idx = headers.index("推定実行時間") if "推定実行時間" in headers else -1
                status_idx = headers.index("ステータス") if "ステータス" in headers else -1
                is_active_idx = headers.index("有効/無効") if "有効/無効" in headers else -1
                is_open_idx = headers.index("開いているか") if "開いているか" in headers else -1
                testsuite_name_idx = headers.index("親テストスイート名")
            except ValueError as e:
                raise ValueError(f"CSVファイルのフォーマットが正しくありません: {str(e)}")
            
            # XMLツリーを構築
            # ルート要素を作成
            testsuite_name = rows[1][testsuite_name_idx]
            root = ET.Element("testsuite", attrib={"name": testsuite_name})
            
            # ノード順序を追加
            node_order = ET.SubElement(root, "node_order")
            node_order.text = ET.CDATA("1")
            
            # 詳細を追加
            details = ET.SubElement(root, "details")
            details.text = ET.CDATA("")
            
            # テストケースをグループ化
            testcase_groups = {}
            for i in range(1, len(rows)):
                row = rows[i]
                testcase_id = row[id_idx]
                if testcase_id not in testcase_groups:
                    testcase_groups[testcase_id] = []
                testcase_groups[testcase_id].append(row)
            
            # 各テストケースを追加
            node_order_index = 0
            for testcase_id, testcase_rows in testcase_groups.items():
                first_row = testcase_rows[0]
                
                testcase = ET.SubElement(root, "testcase", attrib={
                    "internalid": first_row[id_idx],
                    "name": first_row[testcase_name_idx]
                })
                
                # 順序情報
                tc_node_order = ET.SubElement(testcase, "node_order")
                tc_node_order.text = ET.CDATA(str(node_order_index))
                node_order_index += 1
                
                # 外部ID
                external_id = ET.SubElement(testcase, "externalid")
                external_id.text = ET.CDATA(first_row[external_id_idx])
                
                # バージョン
                version = ET.SubElement(testcase, "version")
                version.text = ET.CDATA(first_row[version_idx])
                
                # サマリ
                summary = ET.SubElement(testcase, "summary")
                summary.text = ET.CDATA(f"<p>{first_row[summary_idx]}</p>\n")
                
                # 事前条件
                preconditions = ET.SubElement(testcase, "preconditions")
                preconditions.text = ET.CDATA(self.text_to_html(first_row[preconditions_idx]))
                
                # 実行タイプ
                exec_type = ET.SubElement(testcase, "execution_type")
                exec_type.text = ET.CDATA(first_row[exec_type_idx])
                
                # 重要度
                importance = ET.SubElement(testcase, "importance")
                importance.text = ET.CDATA(first_row[importance_idx])
                
                # 推定実行時間
                if exec_duration_idx >= 0:
                    exec_duration = ET.SubElement(testcase, "estimated_exec_duration")
                    exec_duration.text = first_row[exec_duration_idx]
                
                # ステータス
                if status_idx >= 0:
                    status = ET.SubElement(testcase, "status")
                    status.text = first_row[status_idx]
                
                # 開いているか
                if is_open_idx >= 0:
                    is_open = ET.SubElement(testcase, "is_open")
                    is_open.text = first_row[is_open_idx]
                
                # 有効/無効
                if is_active_idx >= 0:
                    is_active = ET.SubElement(testcase, "active")
                    is_active.text = first_row[is_active_idx]
                
                # ステップを追加
                steps = ET.SubElement(testcase, "steps")
                for row in testcase_rows:
                    step_number = row[step_number_idx]
                    if step_number:
                        step = ET.SubElement(steps, "step")
                        
                        step_num_elem = ET.SubElement(step, "step_number")
                        step_num_elem.text = ET.CDATA(step_number)
                        
                        actions = ET.SubElement(step, "actions")
                        actions.text = ET.CDATA(self.text_to_html(row[actions_idx]))
                        
                        expected = ET.SubElement(step, "expectedresults")
                        expected.text = ET.CDATA(self.text_to_html(row[expected_idx]))
                        
                        step_exec_type = ET.SubElement(step, "execution_type")
                        step_exec_type.text = ET.CDATA(row[exec_type_idx])
            
            # XMLファイルの書き込み
            tree = ET.ElementTree(root)
            
            # XMLファイルへの書き込み前に、文字列化してCDATAセクションを正しくフォーマット
            xml_str = '<?xml version="1.0" encoding="UTF-8"?>\n' + self.element_to_string(root)
            
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(xml_str)
                
        except Exception as e:
            raise Exception(f"CSVをXMLに変換中にエラーが発生: {str(e)}")
    
    def get_element_text(self, element, tag_name):
        """指定されたタグの要素テキストを取得する"""
        tag = element.find(tag_name)
        if tag is not None and tag.text:
            # CDATAタグが残っている場合は除去（二重CDATAの場合に発生する可能性がある）
            text = tag.text
            # テキスト内のCDATAタグを除去
            text = re.sub(r'<!\[CDATA\[(.*?)\]\]>', r'\1', text, flags=re.DOTALL)
            return text
        return ""
    
    def clean_html(self, text):
        """HTMLタグを適切に処理してプレーンテキストに変換する"""
        if not text:
            return ""
        
        # CDATAタグが残っている場合は除去（二重CDATAの場合に発生する可能性がある）
        text = re.sub(r'<!\[CDATA\[(.*?)\]\]>', r'\1', text, flags=re.DOTALL)
        
        # <p>タグは改行に変換
        text = re.sub(r'<p>(.*?)</p>', r'\1\n', text, flags=re.DOTALL)
        
        # <ol>と<ul>タグは削除
        text = re.sub(r'<ol>|</ol>|<ul>|</ul>', '', text)
        
        # <li>タグは・に変換
        text = re.sub(r'<li>(.*?)</li>', r'・\1\n', text, flags=re.DOTALL)
        
        # その他のHTMLタグを削除
        text = re.sub(r'<[^>]+>', '', text)
        
        # 余分な改行を削除
        text = re.sub(r'\n+', '\n', text)
        
        return text
    
    def text_to_html(self, text):
        """プレーンテキストをHTML形式に変換する"""
        if not text:
            return ""
        
        # 改行を<p>タグに変換
        paragraphs = text.split('\n')
        html_paragraphs = []
        
        for p in paragraphs:
            if p.startswith('・'):
                # ・で始まる行は<li>タグに変換
                li_text = p[1:]  # ・を削除
                html_paragraphs.append(f"<li>{li_text}</li>")
            elif p.strip():
                # 空でない行は<p>タグで囲む
                html_paragraphs.append(f"<p>{p}</p>")
        
        # <li>タグが含まれる場合は<ol>タグで囲む
        html = ""
        has_li = any('<li>' in p for p in html_paragraphs)
        
        if has_li:
            li_items = [p for p in html_paragraphs if '<li>' in p]
            non_li_items = [p for p in html_paragraphs if '<li>' not in p]
            
            if li_items:
                html += "<ol>\n" + "\n".join(li_items) + "\n</ol>\n"
            
            if non_li_items:
                html += "\n".join(non_li_items)
        else:
            html = "\n".join(html_paragraphs)
        
        return html
    
    def fix_double_cdata(self, xml_content):
        """二重CDATAタグの問題を修正する"""
        # <![CDATA[<![CDATA[...]]>]]> パターンを <![CDATA[...]]> に置換
        pattern = r'<!\[CDATA\[<!\[CDATA\[(.*?)\]\]>\]\]>'
        fixed_content = re.sub(pattern, r'<![CDATA[\1]]>', xml_content, flags=re.DOTALL)
        return fixed_content
    
    def element_to_string(self, element):
        """ElementTreeの要素を文字列に変換（CDATAセクションを正しく処理）"""
        tag = element.tag
        attrib_str = ""
        if element.attrib:
            attrib_str = " " + " ".join(f'{k}="{v}"' for k, v in element.attrib.items())
        
        result = f"<{tag}{attrib_str}>"
        
        if element.text and isinstance(element.text, ET._ElementUnicodeResult):
            # CDATAセクションの処理
            cdata_content = element.text
            result += f"\n<![CDATA[{cdata_content}]]>\n"
        elif element.text:
            result += element.text
        
        for child in element:
            result += self.element_to_string(child)
        
        result += f"</{tag}>"
        
        if element.tail:
            result += element.tail
        
        return result

def main():
    root = tk.Tk()
    app = TestLinkConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()
