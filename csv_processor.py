import xml.etree.ElementTree as ET
import csv
import re
import codecs
import xml.sax.saxutils as saxutils
import traceback

def text_to_html(text):
    """プレーンテキストをTestLinkが期待するHTML形式（主に<p>, <ol>, <li>）に変換する"""
    if not text:
        return "<p></p>"
    # HTMLエスケープ
    text = saxutils.escape(text)
    lines = text.split('\n')
    html_parts = []
    in_list = False
    for line in lines:
        stripped_line = line.strip()
        if stripped_line.startswith('・'):
            item_text = stripped_line[1:].strip()
            if not in_list:
                html_parts.append("<ol>")
                in_list = True
            html_parts.append(f"<li><p>{item_text}</p></li>")
        else:
            if in_list:
                html_parts.append("</ol>")
                in_list = False
            if stripped_line:
                html_parts.append(f"<p>{stripped_line}</p>")
    if in_list:
        html_parts.append("</ol>")
    if not html_parts:
         return "<p></p>"
    return "\n".join(html_parts)

def element_to_string(element, indent=""):
    """ElementTreeの要素を整形された文字列に変換（特定のタグのみCDATA）"""
    tag = element.tag
    attrib_str = ""
    if element.attrib:
        attrib_str = " " + " ".join(f'{k}="{saxutils.escape(str(v))}"' for k, v in element.attrib.items())

    result = f"{indent}<{tag}{attrib_str}"
    text_content = element.text
    has_children = len(element) > 0

    if has_children or (text_content is not None and text_content.strip()):
        result += ">"
    else:
         result += "></" + tag + ">" # 空要素 <tag></tag>
         return result

    # CDATAで囲むべきタグ
    cdata_tags = ['summary', 'preconditions', 'actions', 'expectedresults', 'details']

    if text_content is not None:
        # テキストをエスケープするかCDATAで囲む
        stripped_text = text_content # strip() しないで元のテキストを保持
        if stripped_text:
            if tag in cdata_tags:
                # CDATA終了区切り文字のエスケープ
                escaped_text = stripped_text.replace(']]>', ']]]]><![CDATA[>')
                # result += f"\n{indent}\t<![CDATA[{escaped_text}]]>\n{indent}" # インデントとCDATA
                # TestLinkはCDATA内のインデントや改行をそのまま解釈することがあるため、
                # CDATA開始直後と終了直前の改行は避ける方が安全かもしれない
                result += f"<![CDATA[{escaped_text}]]>"
            else:
                # 通常のテキストはXMLエスケープ
                result += saxutils.escape(stripped_text)

    if has_children:
        result += "\n"
        for child in element:
            result += element_to_string(child, indent + "\t") + "\n"
        result += f"{indent}</{tag}>"
    elif text_content is not None and text_content.strip(): # テキストのみの場合
        result += f"</{tag}>"
    # 子要素もテキストもない場合は上で処理済み (<tag></tag>)

    return result

def convert_csv_to_xml(csv_file, output_xml_file):
    """CSVファイルを読み込み、TestLinkインポート用のXMLファイルに変換する"""
    try:
        # CSV読み込み
        rows = []
        try:
            with codecs.open(csv_file, 'r', 'shift_jis', errors='replace') as f:
                reader = csv.reader(f)
                headers = next(reader)
                rows.append(headers)
                line_num = 1 # ヘッダーが1行目
                for row in reader:
                    line_num += 1
                    if len(row) != len(headers):
                         print(f"警告: 行 {line_num} の列数がヘッダー ({len(headers)}列) と異なります ({len(row)}列)。スキップします。")
                         continue
                    rows.append(row)
        except StopIteration:
             raise ValueError("CSVファイルにヘッダー行がありません")
        except FileNotFoundError:
             raise Exception(f"CSVファイルが見つかりません: {csv_file}")
        except Exception as e:
             raise Exception(f"CSVファイルの読み込み中にエラーが発生しました: {str(e)}\n{traceback.format_exc()}")

        if len(rows) < 2:
            raise ValueError("CSVファイルにデータ行がありません")

        headers = rows[0]
        # 必須・オプショナルヘッダーの定義とインデックス取得
        required_headers = [
            "テストケース名", "バージョン", "サマリ（概要）", "重要度",
            "ステップ番号", "アクション（手順）", "期待結果", "実行タイプ"
        ]
        # IDと外部IDは必須ではない（新規作成のため）
        optional_headers = ["ID", "外部ID",  "事前条件", "推定実行時間", "ステータス", "有効/無効", "開いているか", "親テストスイート名"]
        header_indices = {}
        missing_required = []
        for header in required_headers:
            try:
                header_indices[header] = headers.index(header)
            except ValueError:
                missing_required.append(header)
        if missing_required:
            raise ValueError(f"CSVファイルに必要なヘッダーが見つかりません: {', '.join(missing_required)}")

        for header in optional_headers:
            try:
                header_indices[header] = headers.index(header)
            except ValueError:
                header_indices[header] = -1 # 見つからない場合は -1

        # インデックスを変数に展開 (可読性のため)
        # オプショナルなものは存在チェック (-1) が必要
        id_idx = header_indices.get("ID", -1)
        external_id_idx = header_indices.get("外部ID", -1)
        version_idx = header_indices["バージョン"]
        testcase_name_idx = header_indices["テストケース名"]
        summary_idx = header_indices["サマリ（概要）"]
        importance_idx = header_indices["重要度"]
        preconditions_idx = header_indices["事前条件"]
        step_number_idx = header_indices["ステップ番号"]
        actions_idx = header_indices["アクション（手順）"]
        expected_idx = header_indices["期待結果"]
        exec_type_idx = header_indices["実行タイプ"]
        exec_duration_idx = header_indices.get("推定実行時間", -1)
        status_idx = header_indices.get("ステータス", -1)
        is_active_idx = header_indices.get("有効/無効", -1)
        is_open_idx = header_indices.get("開いているか", -1)

        # XMLルート要素 <testcases> を作成
        root = ET.Element("testcases")

        # テストケースをグループ化 (IDまたは名前で)
        testcase_groups = {}
        line_num = 1 # ヘッダーが1行目
        for row in rows[1:]: # データ行のみ処理
            line_num += 1
            # IDと名前を取得 (存在しない場合は空文字)
            tc_id = row[id_idx].strip() if id_idx != -1 and id_idx < len(row) else ""
            tc_name = row[testcase_name_idx].strip() if testcase_name_idx != -1 and testcase_name_idx < len(row) else ""

            group_key = ""
            if tc_id:
                group_key = f"ID_{tc_id}"
            elif tc_name: # IDがなく名前がある場合、名前をキーにする
                group_key = f"NAME_{tc_name}"
            else:
                print(f"警告: 行 {line_num} にはテストケースIDも名前もありません。スキップします。")
                continue

            if group_key not in testcase_groups:
                testcase_groups[group_key] = []
            testcase_groups[group_key].append(row)

        # 各グループからテストケースXML要素を生成
        for group_key, testcase_rows in testcase_groups.items():
            if not testcase_rows: continue
            first_row = testcase_rows[0]

            # 必須データの存在チェック (より厳密に)
            missing_data = False
            for header in required_headers:
                idx = header_indices[header]
                # ステップ関連以外は first_row でチェック
                if header not in ["ステップ番号", "アクション（手順）", "期待結果"] and (idx >= len(first_row) or not first_row[idx].strip()):
                     # ステップ実行タイプはステップ行でチェックする or デフォルト値を使う
                     if header == "実行タイプ" and step_number_idx != -1 and first_row[step_number_idx].strip(): # ステップがあれば無視
                          continue
                     print(f"警告: テストケース {group_key} の必須データ「{header}」が不足または空です。スキップします。")
                     missing_data = True
                     break
            if missing_data: continue

            # <testcase> 要素の属性を設定 (internalid は ID があれば追加)
            tc_attributes = {"name": first_row[testcase_name_idx].strip()}
            current_internal_id = first_row[id_idx].strip() if id_idx != -1 and id_idx < len(first_row) else ""
            if current_internal_id:
                tc_attributes["internalid"] = current_internal_id

            testcase = ET.SubElement(root, "testcase", attrib=tc_attributes)

            # <externalid> (値があれば追加)
            external_id_text = first_row[external_id_idx].strip() if external_id_idx != -1 and external_id_idx < len(first_row) else ""
            if external_id_text:
                external_id_elem = ET.SubElement(testcase, "externalid")
                external_id_elem.text = external_id_text

            # <version>
            version_text = first_row[version_idx].strip() if version_idx < len(first_row) else "1"
            version_elem = ET.SubElement(testcase, "version")
            version_elem.text = version_text

            # <summary>
            summary_text = first_row[summary_idx].strip() if summary_idx < len(first_row) else ""
            summary = ET.SubElement(testcase, "summary")
            summary.text = text_to_html(summary_text)

            # <preconditions>
            preconditions_text = first_row[preconditions_idx].strip() if preconditions_idx < len(first_row) else ""
            preconditions = ET.SubElement(testcase, "preconditions")
            preconditions.text = text_to_html(preconditions_text)

            # <execution_type> (Testcaseレベル) - ステップがない場合 or ステップがあっても設定する場合
            # ステップがある場合、Testcaseレベルの実行タイプは必須ではないことが多い
            # ここでは first_row の値を使うが、ステップがあればステップ側が優先される
            tc_exec_type_text = first_row[exec_type_idx].strip() if exec_type_idx < len(first_row) else "1" # デフォルト Manual
            # ステップが存在するかどうかで要素を追加するか決めることもできる
            has_steps = any(row[step_number_idx].strip() for row in testcase_rows if step_number_idx != -1 and step_number_idx < len(row))
            if not has_steps: # ステップがなければTestcaseレベルの実行タイプを設定
                 exec_type_elem = ET.SubElement(testcase, "execution_type")
                 exec_type_elem.text = tc_exec_type_text

            # <importance>
            importance_text = first_row[importance_idx].strip() if importance_idx < len(first_row) else "2" # デフォルト Medium
            importance = ET.SubElement(testcase, "importance")
            importance.text = importance_text

            # オプショナル要素 (値があれば追加)
            if exec_duration_idx != -1 and exec_duration_idx < len(first_row) and first_row[exec_duration_idx].strip():
                exec_duration = ET.SubElement(testcase, "estimated_exec_duration")
                exec_duration.text = first_row[exec_duration_idx].strip()
            if status_idx != -1 and status_idx < len(first_row) and first_row[status_idx].strip():
                status = ET.SubElement(testcase, "status")
                status.text = first_row[status_idx].strip()
            if is_open_idx != -1 and is_open_idx < len(first_row) and first_row[is_open_idx].strip():
                is_open = ET.SubElement(testcase, "is_open")
                is_open.text = first_row[is_open_idx].strip()
            if is_active_idx != -1 and is_active_idx < len(first_row) and first_row[is_active_idx].strip():
                active = ET.SubElement(testcase, "active")
                active.text = first_row[is_active_idx].strip()

            # <steps> 要素
            steps_container = ET.SubElement(testcase, "steps")
            step_line_num = 1 # グループ内の行番号 (デバッグ用)
            for row in testcase_rows:
                step_line_num += 1
                # ステップ番号がある行のみステップとして処理
                step_number = row[step_number_idx].strip() if step_number_idx != -1 and step_number_idx < len(row) else ""
                if step_number:
                    # ステップに必要なデータのチェック
                    actions_text = row[actions_idx].strip() if actions_idx < len(row) else ""
                    expected_text = row[expected_idx].strip() if expected_idx < len(row) else ""
                    step_exec_type_text = row[exec_type_idx].strip() if exec_type_idx < len(row) else "1" # ステップ実行タイプ
                    if not actions_text or not expected_text:
                         print(f"警告: テストケース {group_key} のステップ番号 {step_number} (CSV行: {step_line_num}) でアクションまたは期待結果が空です。")
                         # 空でもステップを生成するか、スキップするかは仕様による。ここでは生成する。

                    step = ET.SubElement(steps_container, "step")
                    step_num_elem = ET.SubElement(step, "step_number")
                    step_num_elem.text = step_number
                    actions = ET.SubElement(step, "actions")
                    actions.text = text_to_html(actions_text)
                    expected = ET.SubElement(step, "expectedresults")
                    expected.text = text_to_html(expected_text)
                    step_exec_type = ET.SubElement(step, "execution_type")
                    step_exec_type.text = step_exec_type_text

        # XMLをファイルに書き込み
        xml_string = '<?xml version="1.0" encoding="UTF-8"?>\n' + element_to_string(root)
        # 出力前に不要な空行などを削除する（オプション）
        xml_string = "\n".join(line for line in xml_string.splitlines() if line.strip())

        with open(output_xml_file, 'w', encoding='utf-8') as f:
            f.write(xml_string)

    except ValueError as ve: # CSVフォーマットエラーなど
        raise Exception(f"CSVファイルの処理中にエラーが発生しました: {str(ve)}\n{traceback.format_exc()}")
    except Exception as e:
        raise Exception(f"CSVからXMLへの変換中に予期せぬエラーが発生しました: {str(e)}\n{traceback.format_exc()}")