import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import xml.etree.ElementTree as ET
import csv
import re
import codecs
import traceback
import random

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
            error_details = traceback.format_exc()
            self.update_status(f"エラー: {str(e)}")
            messagebox.showerror("エラー", f"変換中にエラーが発生しました:\n{str(e)}\n\n詳細:\n{error_details}")
    
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
            testcase_id = testcase.get("internalid", "")
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
            # CSVファイルの読み込み (エラーハンドリング強化)
            rows = []
            try:
                with codecs.open(csv_file, 'r', 'shift_jis') as f:
                    reader = csv.reader(f)
                    headers = next(reader) # ヘッダー行を先に取得
                    rows.append(headers)
                    for i, row in enumerate(reader):
                         # 行の列数がヘッダーと一致するか確認 (より厳密なチェック)
                         if len(row) != len(headers):
                              print(f"警告: 行 {i+2} の列数がヘッダー ({len(headers)}列) と異なります ({len(row)}列)。スキップします。")
                              continue
                         rows.append(row)
            except StopIteration: # 空のCSVファイルの場合
                 raise ValueError("CSVファイルにヘッダー行がありません")
            except FileNotFoundError:
                 raise Exception(f"CSVファイルが見つかりません: {csv_file}")
            except Exception as e:
                 raise Exception(f"CSVファイルの読み込み中にエラーが発生しました: {str(e)}\n{traceback.format_exc()}")


            if len(rows) < 2:
                # ヘッダーしかない場合もデータがないとみなす
                raise ValueError("CSVファイルにデータ行がありません")

            # ヘッダー行
            headers = rows[0]

            # 必要なカラムのインデックスを取得 (エラーハンドリング強化)
            required_headers = [
                "ID", "外部ID", "バージョン", "テストケース名", "サマリ（概要）",
                "重要度", "事前条件", "ステップ番号", "アクション（手順）", "期待結果",
                "実行タイプ", "親テストスイート名"
            ]
            optional_headers = ["推定実行時間", "ステータス", "有効/無効", "開いているか"]
            header_indices = {}
            missing_headers = []

            for header in required_headers:
                 try:
                     header_indices[header] = headers.index(header)
                 except ValueError:
                     missing_headers.append(header)

            if missing_headers:
                 raise ValueError(f"CSVファイルに必要なヘッダーが見つかりません: {', '.join(missing_headers)}")

            for header in optional_headers:
                 try:
                     header_indices[header] = headers.index(header)
                 except ValueError:
                     header_indices[header] = -1 # 見つからない場合は -1

            # インデックスを変数に展開
            id_idx = header_indices["ID"]
            external_id_idx = header_indices["外部ID"]
            version_idx = header_indices["バージョン"]
            testcase_name_idx = header_indices["テストケース名"]
            summary_idx = header_indices["サマリ（概要）"]
            importance_idx = header_indices["重要度"]
            preconditions_idx = header_indices["事前条件"]
            step_number_idx = header_indices["ステップ番号"]
            actions_idx = header_indices["アクション（手順）"]
            expected_idx = header_indices["期待結果"]
            exec_type_idx = header_indices["実行タイプ"]
            testsuite_name_idx = header_indices["親テストスイート名"]
            exec_duration_idx = header_indices["推定実行時間"]
            status_idx = header_indices["ステータス"]
            is_active_idx = header_indices["有効/無効"]
            is_open_idx = header_indices["開いているか"]


            # XMLツリーを構築
            # ルート要素を作成 (CSVが空でないことを確認済み)
            testsuite_name = rows[1][testsuite_name_idx]
            testsuite_id = str(random.randint(1, 1000))
            root = ET.Element("testsuite", attrib={"id": testsuite_id, "name": testsuite_name})

            # ノード順序を追加
            node_order = ET.SubElement(root, "node_order")
            # ET.CDATA の呼び出しを削除し、単純な文字列を代入
            node_order.text = "1" # <--- 修正

            # 詳細を追加
            details = ET.SubElement(root, "details")
            # ET.CDATA の呼び出しを削除し、単純な文字列を代入
            details.text = "" # <--- 修正

            # テストケースをグループ化
            testcase_groups = {}
            for i in range(1, len(rows)):
                row = rows[i]
                # ID列が存在するかチェック (IndexError防止)
                if len(row) <= id_idx:
                     print(f"警告: 行 {i+1} のデータが不足しています（ID列なし）。スキップします。")
                     continue
                testcase_id = row[id_idx]
                if testcase_id not in testcase_groups:
                    testcase_groups[testcase_id] = []
                testcase_groups[testcase_id].append(row)

            # 各テストケースを追加
            node_order_index = 0
            for testcase_id, testcase_rows in testcase_groups.items():
                if not testcase_rows: continue # 空のグループはスキップ
                first_row = testcase_rows[0]

                # --- 必要な列が存在するかチェック (IndexError防止) ---
                required_indices_for_tc = [id_idx, testcase_name_idx, external_id_idx, version_idx, summary_idx, preconditions_idx, exec_type_idx, importance_idx]
                if any(idx >= len(first_row) for idx in required_indices_for_tc):
                     print(f"警告: テストケース ID {testcase_id} の最初の行に必要なデータが不足しています。スキップします。")
                     continue
                # --- チェック完了 ---

                testcase = ET.SubElement(root, "testcase", attrib={
                    "internalid": first_row[id_idx],
                    "name": first_row[testcase_name_idx]
                })

                # 順序情報
                tc_node_order = ET.SubElement(testcase, "node_order")
                # ET.CDATA の呼び出しを削除
                tc_node_order.text = str(node_order_index) # <--- 修正
                node_order_index += 1

                # 外部ID
                external_id = ET.SubElement(testcase, "externalid")
                 # ET.CDATA の呼び出しを削除
                external_id.text = first_row[external_id_idx] # <--- 修正

                # バージョン
                version = ET.SubElement(testcase, "version")
                # ET.CDATA の呼び出しを削除
                version.text = first_row[version_idx] # <--- 修正

                # サマリ
                summary = ET.SubElement(testcase, "summary")
                summary_text = self.ensure_paragraph_tags(first_row[summary_idx])
                # ET.CDATA の呼び出しを削除
                summary.text = f"{summary_text}\n" # <--- 修正 (末尾の改行はTestLinkフォーマットに合わせるため)

                # 事前条件
                preconditions = ET.SubElement(testcase, "preconditions")
                # ET.CDATA の呼び出しを削除
                preconditions.text = self.text_to_html(first_row[preconditions_idx]) # <--- 修正

                # 実行タイプ
                exec_type = ET.SubElement(testcase, "execution_type")
                # ET.CDATA の呼び出しを削除
                exec_type.text = first_row[exec_type_idx] # <--- 修正

                # 重要度
                importance = ET.SubElement(testcase, "importance")
                # ET.CDATA の呼び出しを削除
                importance.text = first_row[importance_idx] # <--- 修正

                # 推定実行時間 (オプション)
                if exec_duration_idx != -1 and exec_duration_idx < len(first_row):
                    exec_duration = ET.SubElement(testcase, "estimated_exec_duration")
                    exec_duration.text = first_row[exec_duration_idx]

                # ステータス (オプション)
                if status_idx != -1 and status_idx < len(first_row):
                    status = ET.SubElement(testcase, "status")
                    status.text = first_row[status_idx]

                # 開いているか (オプション)
                if is_open_idx != -1 and is_open_idx < len(first_row):
                    is_open = ET.SubElement(testcase, "is_open")
                    is_open.text = first_row[is_open_idx]

                # 有効/無効 (オプション)
                if is_active_idx != -1 and is_active_idx < len(first_row):
                    is_active = ET.SubElement(testcase, "active")
                    is_active.text = first_row[is_active_idx]

                # ステップを追加
                steps = ET.SubElement(testcase, "steps")
                for row_idx, row in enumerate(testcase_rows):
                     # --- 必要な列が存在するかチェック (IndexError防止) ---
                     required_indices_for_step = [step_number_idx, actions_idx, expected_idx, exec_type_idx]
                     if any(idx >= len(row) for idx in required_indices_for_step):
                           print(f"警告: テストケース ID {testcase_id} のステップ {row_idx+1} に必要なデータが不足しています。スキップします。")
                           continue
                     # --- チェック完了 ---

                     step_number = row[step_number_idx]
                     # ステップ番号が空でない場合のみステップ要素を作成
                     if step_number and step_number.strip():
                        step = ET.SubElement(steps, "step")

                        step_num_elem = ET.SubElement(step, "step_number")
                        # ET.CDATA の呼び出しを削除
                        step_num_elem.text = step_number # <--- 修正

                        actions = ET.SubElement(step, "actions")
                        # ET.CDATA の呼び出しを削除
                        actions.text = self.text_to_html(row[actions_idx]) # <--- 修正

                        expected = ET.SubElement(step, "expectedresults")
                        # ET.CDATA の呼び出しを削除
                        expected.text = self.text_to_html(row[expected_idx]) # <--- 修正

                        step_exec_type = ET.SubElement(step, "execution_type")
                        # ET.CDATA の呼び出しを削除
                        step_exec_type.text = row[exec_type_idx] # <--- 修正

            # XMLファイルの書き込み
            # 修正した element_to_string 関数を使用
            xml_str = '<?xml version="1.0" encoding="UTF-8"?>\n' + self.element_to_string(root)

            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(xml_str)

        except ValueError as ve: # CSVフォーマットエラーやデータ不足をキャッチ
             raise Exception(f"CSVファイルの処理中にエラーが発生しました: {str(ve)}\n{traceback.format_exc()}")
        except Exception as e:
            # 予期せぬエラー
            raise Exception(f"CSVをXMLに変換中に予期せぬエラーが発生しました: {str(e)}\n{traceback.format_exc()}")

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
        # 注意: <p>&nbsp;</p> のような空の段落も改行に変換される
        text = re.sub(r'<p>(.*?)</p>', r'\1\n', text, flags=re.DOTALL | re.IGNORECASE) # IGNORECASE を追加

        # <br> タグも改行に変換
        text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)

        # <ol>と<ul>タグとその中の<li>タグの処理を改善
        # まずリスト全体のHTMLを取得し、リストアイテムをプレーンテキストに変換
        def replace_list(match):
            list_content = match.group(1)
            items = re.findall(r'<li.*?>(.*?)</li>', list_content, flags=re.DOTALL | re.IGNORECASE)
            # 各アイテムのHTMLタグを除去し、先頭に「・」を付けて改行で結合
            plain_items = ['・' + self.clean_html(item).strip() for item in items]
            return '\n'.join(plain_items) + '\n' # リストの後にも改行を追加

        text = re.sub(r'<(?:ul|ol).*?>(.*?)</(?:ul|ol)>', replace_list, text, flags=re.DOTALL | re.IGNORECASE)

        # <li>タグ単体（リストタグの外にある場合など）も処理（通常は考えにくいが念のため）
        text = re.sub(r'<li.*?>(.*?)</li>', r'・\1\n', text, flags=re.DOTALL | re.IGNORECASE)

        # その他のHTMLタグを削除 (<p>, <br>, <ul>, <ol>, <li> 以外)
        text = re.sub(r'<(?!\/?(p|br|ul|ol|li)\b)[^>]+>', '', text, flags=re.IGNORECASE)

        # --- 修正箇所 ---
        # HTMLエンティティをデコード (&nbsp; を空文字列に置換)
        text = text.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&").replace("&nbsp;", "")
        # 必要であれば他のHTMLエンティティもここに追加できます (例: .replace("&quot;", "\""))
        # ----------------

        # 連続する改行やスペースを整理
        text = re.sub(r'[ \t]+', ' ', text) # 連続するスペースやタブを1つのスペースに
        text = re.sub(r'\n\s*\n', '\n', text) # 複数の改行（間の空白含む）を1つの改行に

        return text.strip() # 前後の空白と改行を削除して返す


    def ensure_paragraph_tags(self, text):
        """テキストが<p>タグで囲まれていることを確認する"""
        if not text:
            return "<p></p>"
        
        text = text.strip()
        
        # 既に<p>タグで囲まれている場合はそのまま返す
        if text.startswith("<p>") and text.endswith("</p>"):
            return text
        
        # <p>タグで囲む
        return f"<p>{text}</p>"
    
    def text_to_html(self, text):
        """プレーンテキストをHTML形式に変換する"""
        if not text:
            return ""
        
        # HTMLエンティティをデコード（CSVから読み込んだデータに含まれる可能性がある）
        text = text.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
        
        # すでにHTMLタグが含まれている場合はそのまま返す
        if re.search(r'<[a-z]+>', text):
            return text
            
        # 改行を分割
        paragraphs = text.split('\n')
        html_paragraphs = []
        
        for p in paragraphs:
            if p.startswith('・'):
                # ・で始まる行は<li>タグに変換
                li_text = p[1:].strip()  # ・を削除
                html_paragraphs.append(f"<li>{li_text}</li>")
            elif p.strip():
                # 空でない行は<p>タグで囲む
                html_paragraphs.append(f"<p>{p.strip()}</p>")
        
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
    
    # この関数内で xml.sax.saxutils を使うので、ファイルの先頭でインポートしておきます
    # import xml.sax.saxutils

    def element_to_string(self, element, indent=""):
        """ElementTreeの要素を文字列に変換（特定のタグのみCDATAセクションとして処理）"""
        tag = element.tag
        attrib_str = ""
        if element.attrib:
            attrib_str = " " + " ".join(f'{k}="{v}"' for k, v in element.attrib.items())

        result = f"{indent}<{tag}{attrib_str}>"

        has_children = len(element) > 0
        text_content = element.text # 元のテキストを保持

        # --- ここから修正: タグ名に基づいてCDATA処理を分岐 ---
        # CDATAで囲むべきタグ名のリスト (HTMLや複数行テキストを含む可能性があるもの)
        # 元のXMLを参考に調整
        cdata_tags = [
            'summary', 'preconditions', 'actions', 'expectedresults', 'details',
            'node_order', 'externalid', 'version', 'step_number', 'execution_type', 'importance'
            # execution_type, importance もCDATAの場合があるため含める
            # 'name' 属性はCDATAではないので注意
        ]

        # テキスト内容が存在する場合 (Noneでない)
        if text_content is not None:
            if tag in cdata_tags:
                # CDATAが必要なタグの場合
                escaped_text = text_content.replace(']]>', ']]]]><![CDATA[>')
                result += f"\n{indent}\t<![CDATA[{escaped_text}]]>\n"
                if not has_children:
                     result += f"{indent}</{tag}>"
            else:
                # CDATAが不要なタグの場合 (例: status, is_open, active, estimated_exec_duration)
                # XML特殊文字 (<, >, &) をエスケープ
                import xml.sax.saxutils as saxutils
                escaped_text = saxutils.escape(text_content)
                # テキストが空文字列でない場合のみ出力 (空の<tag></tag>にする)
                if escaped_text:
                     result += escaped_text
                if not has_children:
                     result += f"</{tag}>" # 閉じタグを追加 (<tag>text</tag> or <tag></tag>)

        # テキスト内容がない場合 (None)
        else:
             # estimated_exec_duration のような空要素は <tag></tag> とする
             # (TestLinkは <tag/> 形式も <tag></tag> 形式も解釈できるはず)
             if has_children:
                  result += "\n"

        # 子要素を再帰的に処理
        if has_children:
            for child in element:
                result += self.element_to_string(child, indent + "\t") + "\n"
            result += f"{indent}</{tag}>"
        # 子要素がなく、テキストもない場合 (text_content is None)
        elif text_content is None:
             result += f"</{tag}>" # <tag></tag> 形式

        return result

def main():
    root = tk.Tk()
    app = TestLinkConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()
