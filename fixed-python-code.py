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
import xml.sax.saxutils as saxutils # 追加 (element_to_string で必要)

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
            # ルート要素が testsuite か testcases かを判定
            if '<testsuite ' in xml_content[:200]: # ファイルの先頭付近で判定
                root_element = ET.fromstring(xml_content)
                testsuite_name = root_element.get("name", "")
                testcases_root = root_element # testsuite の下に testcase がある場合
            elif '<testcases>' in xml_content[:200]:
                root_element = ET.fromstring(xml_content)
                testsuite_name = "" # testcases 直下の場合、特定のスイート名はなし
                testcases_root = root_element # testcases の下に testcase がある場合
            else:
                 # 想定外の形式の場合、最初の要素から取得を試みる
                 root_element = ET.fromstring(xml_content)
                 # testsuite または testcases を探す
                 possible_root = root_element.find('.//testsuite') or root_element.find('.//testcases') or root_element
                 testsuite_name = possible_root.get("name", "") if possible_root.tag == 'testsuite' else ""
                 testcases_root = possible_root

            # 出力CSVファイル名の生成
            output_file = os.path.splitext(xml_file)[0] + ".csv"

            # データの変換と出力
            self.convert_xml_to_csv_internal(testcases_root, testsuite_name, output_file)

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
            self.convert_csv_to_xml_internal(csv_file, output_file) # 内部関数を呼び出す

            self.update_status(f"変換完了: {output_file}")
            messagebox.showinfo("変換完了", f"XMLファイルに変換しました:\n{output_file}")

        except Exception as e:
            error_details = traceback.format_exc()
            self.update_status(f"エラー: {str(e)}")
            messagebox.showerror("エラー", f"変換中にエラーが発生しました:\n{str(e)}\n\n詳細:\n{error_details}")

    def convert_xml_to_csv_internal(self, testcases_root, testsuite_name, output_file):
        """XMLデータをCSVに変換する（内部処理）"""
        # CSVに書き込むデータのリスト
        rows = []

        # ヘッダー行
        headers = [
            "ID", "外部ID", "バージョン", "テストケース名", "サマリ（概要）",
            "重要度", "事前条件", "ステップ番号", "アクション（手順）", "期待結果",
            "実行タイプ", "推定実行時間", "ステータス", "有効/無効", "開いているか",
            "親テストスイート名" # testsuite名を取得できればここに入れる
        ]
        rows.append(headers)

        # 各テストケースを処理 (XPathで testcase を検索)
        for testcase in testcases_root.findall(".//testcase"):
            testcase_id = testcase.get("internalid", "")
            external_id = self.get_element_text(testcase, "externalid")
            version = self.get_element_text(testcase, "version")
            testcase_name = testcase.get("name", "")
            summary = self.clean_html(self.get_element_text(testcase, "summary"))
            importance = self.get_element_text(testcase, "importance")
            preconditions = self.clean_html(self.get_element_text(testcase, "preconditions"))
            exec_type_elem = testcase.find("execution_type")
            exec_type = exec_type_elem.text if exec_type_elem is not None and exec_type_elem.text is not None else "" # CDATA対応
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
                    step_exec_type_elem = step.find("execution_type")
                    step_exec_type = step_exec_type_elem.text if step_exec_type_elem is not None and step_exec_type_elem.text is not None else "" # CDATA対応

                    # 行を追加
                    row = [
                        testcase_id, external_id, version, testcase_name, summary,
                        importance, preconditions, step_number, actions, expected,
                        step_exec_type, exec_duration, status, is_active, is_open,
                        testsuite_name # testsuite名を追加
                    ]
                    rows.append(row)
            else:
                # ステップがない場合は空行を追加
                row = [
                    testcase_id, external_id, version, testcase_name, summary,
                    importance, preconditions, "", "", "",
                    exec_type, exec_duration, status, is_active, is_open,
                    testsuite_name # testsuite名を追加
                ]
                rows.append(row)

        # CSVファイルへの書き込み
        with codecs.open(output_file, 'w', 'shift_jis', errors='ignore') as f: # エラー無視を追加
            writer = csv.writer(f, quoting=csv.QUOTE_ALL)
            writer.writerows(rows)

    def convert_csv_to_xml_internal(self, csv_file, output_file):
        """CSVデータをXMLに変換する（内部処理、ルート要素を<testcases>に修正）"""
        try:
            # CSVファイルの読み込み (エラーハンドリング強化)
            rows = []
            try:
                with codecs.open(csv_file, 'r', 'shift_jis', errors='replace') as f: # エラー時の置換を設定
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
            # 親テストスイート名はXML生成には直接使わないが、CSVには存在するのでインデックス取得は試みる
            required_headers = [
                "ID", "外部ID", "バージョン", "テストケース名", "サマリ（概要）",
                "重要度", "事前条件", "ステップ番号", "アクション（手順）", "期待結果",
                "実行タイプ"
            ]
            optional_headers = ["推定実行時間", "ステータス", "有効/無効", "開いているか", "親テストスイート名"]
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
            # オプショナルなインデックス
            exec_duration_idx = header_indices["推定実行時間"]
            status_idx = header_indices["ステータス"]
            is_active_idx = header_indices["有効/無効"]
            is_open_idx = header_indices["開いているか"]
            # testsuite_name_idx はXML生成に使わないが、存在は確認しておく
            # testsuite_name_idx = header_indices["親テストスイート名"]


            # --- ここから修正 ---
            # XMLツリーを構築
            # ルート要素を <testcases> に変更 (TestLinkのテストケースインポート形式に合わせる)
            root = ET.Element("testcases")
            # 以前の <testsuite> 要素に関連するコード (id, name, node_order, details の生成) は削除
            # --- 修正ここまで ---

            # テストケースをグループ化
            testcase_groups = {}
            for i in range(1, len(rows)):
                row = rows[i]
                # ID列が存在するかチェック (IndexError防止)
                if len(row) <= id_idx:
                     print(f"警告: 行 {i+1} のデータが不足しています（ID列なし）。スキップします。")
                     continue
                testcase_id = row[id_idx]
                 # テストケースIDが空や空白でないことを確認
                if not testcase_id or testcase_id.isspace():
                     print(f"警告: 行 {i+1} のテストケースIDが空です。スキップします。")
                     continue
                if testcase_id not in testcase_groups:
                    testcase_groups[testcase_id] = []
                testcase_groups[testcase_id].append(row)

            # 各テストケースを追加
            node_order_index = 0 # testcase内のnode_order用インデックス
            for testcase_id, testcase_rows in testcase_groups.items():
                if not testcase_rows: continue # 空のグループはスキップ
                first_row = testcase_rows[0]

                # --- 必要な列が存在するかチェック (IndexError防止) ---
                required_indices_for_tc = [id_idx, testcase_name_idx, external_id_idx, version_idx, summary_idx, preconditions_idx, exec_type_idx, importance_idx]
                if any(idx >= len(first_row) or idx < 0 for idx in required_indices_for_tc): # idx < 0 もチェック
                     print(f"警告: テストケース ID {testcase_id} の最初の行に必要なデータが不足またはインデックスが無効です。スキップします。")
                     continue
                # --- チェック完了 ---

                testcase = ET.SubElement(root, "testcase", attrib={
                    "internalid": first_row[id_idx],
                    "name": first_row[testcase_name_idx]
                })

                # 順序情報 (testcase内のnode_order)
                tc_node_order = ET.SubElement(testcase, "node_order")
                # CDATA の呼び出しを削除
                tc_node_order.text = str(node_order_index)
                node_order_index += 1

                # 外部ID
                external_id_elem = ET.SubElement(testcase, "externalid")
                # CDATA の呼び出しを削除
                external_id_elem.text = first_row[external_id_idx]

                # バージョン
                version_elem = ET.SubElement(testcase, "version")
                # CDATA の呼び出しを削除
                version_elem.text = first_row[version_idx]

                # サマリ
                summary = ET.SubElement(testcase, "summary")
                # self.ensure_paragraph_tags を削除し、text_to_html で <p> を付けるように試みる
                summary_text = self.text_to_html(first_row[summary_idx])
                # CDATA の呼び出しを削除
                summary.text = f"{summary_text}\n" if summary_text else "\n" # 空でも改行は入れる

                # 事前条件
                preconditions = ET.SubElement(testcase, "preconditions")
                # CDATA の呼び出しを削除
                preconditions.text = self.text_to_html(first_row[preconditions_idx])

                # 実行タイプ
                exec_type = ET.SubElement(testcase, "execution_type")
                # CDATA の呼び出しを削除
                exec_type.text = first_row[exec_type_idx]

                # 重要度
                importance = ET.SubElement(testcase, "importance")
                # CDATA の呼び出しを削除
                importance.text = first_row[importance_idx]

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
                     if any(idx >= len(row) or idx < 0 for idx in required_indices_for_step): # idx < 0 もチェック
                           print(f"警告: テストケース ID {testcase_id} のステップ {row_idx+1} に必要なデータが不足またはインデックスが無効です。スキップします。")
                           continue
                     # --- チェック完了 ---

                     step_number = row[step_number_idx]
                     # ステップ番号が空でない場合のみステップ要素を作成
                     if step_number and step_number.strip():
                        step = ET.SubElement(steps, "step")

                        step_num_elem = ET.SubElement(step, "step_number")
                        # CDATA の呼び出しを削除
                        step_num_elem.text = step_number

                        actions = ET.SubElement(step, "actions")
                        # CDATA の呼び出しを削除
                        actions.text = self.text_to_html(row[actions_idx])

                        expected = ET.SubElement(step, "expectedresults")
                        # CDATA の呼び出しを削除
                        expected.text = self.text_to_html(row[expected_idx])

                        step_exec_type = ET.SubElement(step, "execution_type")
                        # CDATA の呼び出しを削除
                        step_exec_type.text = row[exec_type_idx]

            # XMLファイルの書き込み
            # element_to_string 関数を使って整形して出力
            xml_str = '<?xml version="1.0" encoding="UTF-8"?>\n' + self.element_to_string(root)

            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(xml_str)

        except ValueError as ve: # CSVフォーマットエラーやデータ不足をキャッチ
             raise Exception(f"CSVファイルの処理中にエラーが発生しました: {str(ve)}\n{traceback.format_exc()}")
        except Exception as e:
            # 予期せぬエラー
            raise Exception(f"CSVをXMLに変換中に予期せぬエラーが発生しました: {str(e)}\n{traceback.format_exc()}")

    def get_element_text(self, element, tag_name):
        """指定されたタグの要素テキストを取得する (CDATA対応強化)"""
        tag = element.find(tag_name)
        if tag is not None:
            # .text を直接参照し、Noneでないか確認
            text = tag.text
            if text:
                 # テキスト内のCDATAタグを除去 (念のため)
                 text = re.sub(r'<!\[CDATA\[(.*?)\]\]>', r'\1', text, flags=re.DOTALL)
                 return text.strip() # 前後の空白除去を追加
        return ""

    def clean_html(self, text):
        """HTMLタグを適切に処理してプレーンテキストに変換する"""
        if not text:
            return ""

        # CDATAタグが残っている場合は除去
        text = re.sub(r'<!\[CDATA\[(.*?)\]\]>', r'\1', text, flags=re.DOTALL)

        # <p>タグは改行に変換 (前後の空白トリムを追加)
        text = re.sub(r'<p>(.*?)</p>', lambda m: m.group(1).strip() + '\n', text, flags=re.DOTALL | re.IGNORECASE)

        # <br> タグも改行に変換
        text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)

        # リストタグ処理の改善 (リストアイテム間の改行を考慮)
        def replace_list(match):
            list_content = match.group(1)
            items = re.findall(r'<li.*?>(.*?)</li>', list_content, flags=re.DOTALL | re.IGNORECASE)
            plain_items = []
            for item in items:
                # 各アイテム内の<p>や<br>も改行として扱い、最後に結合
                cleaned_item = self.clean_html(item).strip()
                # アイテム内の改行を保持しつつ先頭に「・」を追加
                lines = [f"・{line.strip()}" for line in cleaned_item.split('\n') if line.strip()]
                plain_items.extend(lines)
            # リストアイテム間は改行で結合し、リストの後にも改行を追加
            return '\n'.join(plain_items) + '\n' if plain_items else '\n'

        text = re.sub(r'<(ul|ol).*?>(.*?)</\1>', replace_list, text, flags=re.DOTALL | re.IGNORECASE)

        # 残ったリストタグの外の<li>タグ（通常はないはず）
        text = re.sub(r'<li.*?>(.*?)</li>', lambda m: '・' + self.clean_html(m.group(1)).strip() + '\n', text, flags=re.DOTALL | re.IGNORECASE)

        # その他のHTMLタグを削除
        text = re.sub(r'<(?!\/?(p|br|ul|ol|li)\b)[^>]+>', '', text, flags=re.IGNORECASE)

        # HTMLエンティティをデコード (&nbsp; を半角スペースに、他は標準的なもの)
        text = text.replace("&nbsp;", " ")
        text = text.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&").replace("&quot;", "\"").replace("&#39;", "'")

        # 連続する改行やスペースを整理
        text = re.sub(r'[ \t]+', ' ', text) # 連続するスペースやタブを1つのスペースに
        text = re.sub(r'\n\s*\n+', '\n', text) # 複数の改行（間の空白含む）を1つの改行に

        return text.strip() # 前後の空白と改行を削除して返す

    def text_to_html(self, text):
        """プレーンテキストをTestLinkが期待するHTML形式（主に<p>, <ol>, <li>）に変換する"""
        if not text:
            return "<p></p>" # 空の場合は空の<p>タグを返す

        # HTMLエンティティをエスケープ ( < > & )
        text = saxutils.escape(text)

        lines = text.split('\n')
        html_parts = []
        in_list = False

        for line in lines:
            stripped_line = line.strip()
            if stripped_line.startswith('・'):
                # リストアイテム
                item_text = stripped_line[1:].strip()
                if not in_list:
                    html_parts.append("<ol>")
                    in_list = True
                html_parts.append(f"<li><p>{item_text}</p></li>") # <li>の中に<p>を入れる形式に変更
            else:
                # 通常のテキスト行
                if in_list:
                    html_parts.append("</ol>")
                    in_list = False
                if stripped_line: # 空行は無視
                    html_parts.append(f"<p>{stripped_line}</p>")

        if in_list:
            html_parts.append("</ol>") # 最後にリストが閉じていなければ閉じる

        # 何も要素がない場合は空の<p>を返す
        if not html_parts:
             return "<p></p>"

        return "\n".join(html_parts)

    def fix_double_cdata(self, xml_content):
        """二重CDATAタグの問題を修正する"""
        # <![CDATA[<![CDATA[...]]>]]> パターンを <![CDATA[...]]> に置換
        pattern = r'<!\[CDATA\[\s*<!\[CDATA\[(.*?)]]>\s*]]>' # 前後の空白も考慮
        fixed_content = re.sub(pattern, r'<![CDATA[\1]]>', xml_content, flags=re.DOTALL)
        # <![CDATA[<p>...]]> のようなパターンも修正対象か検討 (今回は見送り)
        return fixed_content

    def element_to_string(self, element, indent=""):
        """ElementTreeの要素を文字列に変換（特定のタグのみCDATAセクションとして処理）"""
        tag = element.tag
        attrib_str = ""
        if element.attrib:
            attrib_str = " " + " ".join(f'{k}="{saxutils.escape(str(v))}"' for k, v in element.attrib.items()) # 属性値もエスケープ

        result = f"{indent}<{tag}{attrib_str}"

        text_content = element.text
        has_children = len(element) > 0

        # 子要素があるか、またはテキスト内容が空でない文字列の場合のみ閉じタグの前に > をつける
        if has_children or (text_content is not None and text_content.strip()):
            result += ">"
        else:
             # 子要素がなくテキストも実質ない場合は自己終了タグ <tag/> とする (TestLinkが許容するか注意)
             # TestLink は <tag></tag> の方が安全かもしれないので、やっぱり閉じタグ形式にする
             # result += " />"
             result += "></" + tag + ">" # <tag></tag> 形式
             return result # 自己終了タグ(または空要素)の場合はここで終了


        # CDATAで囲むべきタグ名 (TestLinkの慣習に合わせる)
        # HTMLを含む可能性のある要素や、改行を含むテキスト要素を指定
        cdata_tags = [
            'summary', 'preconditions', 'actions', 'expectedresults', 'details'
        ]

        if text_content is not None:
            text_content = text_content.strip() # 前後の空白を除去して判定
            if text_content: # 空白のみでない場合
                if tag in cdata_tags:
                    # CDATAが必要なタグの場合、エスケープ処理を追加
                    escaped_text = text_content.replace(']]>', ']]]]><![CDATA[>') # CDATA終了区切り文字のエスケープ
                    result += f"\n{indent}\t<![CDATA[{escaped_text}]]>\n{indent}" # インデント調整
                else:
                    # CDATAが不要なタグの場合、XML特殊文字をエスケープ
                    escaped_text = saxutils.escape(text_content)
                    # 必要なら改行でインデントする (今回はシンプルにインラインで)
                    result += escaped_text # インデントなしでテキストを直接追加

        # 子要素を再帰的に処理
        if has_children:
            result += "\n" # 子要素の前に改行
            for child in element:
                result += self.element_to_string(child, indent + "\t") + "\n"
            result += f"{indent}</{tag}>" # 子要素を閉じるインデント
        # 子要素がなく、テキスト内容があった場合
        elif text_content is not None and text_content.strip():
            result += f"</{tag}>" # 閉じタグ
        # 子要素がなく、テキストもない場合は上で処理済み (<tag></tag>)

        return result

def main():
    root = tk.Tk()
    app = TestLinkConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()