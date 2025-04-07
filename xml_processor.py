import xml.etree.ElementTree as ET
import csv
import re
import codecs
import traceback

def fix_double_cdata(xml_content):
    """二重CDATAタグの問題を修正する"""
    pattern = r'<!\[CDATA\[\s*<!\[CDATA\[(.*?)]]>\s*]]>'
    fixed_content = re.sub(pattern, r'<![CDATA[\1]]>', xml_content, flags=re.DOTALL)
    return fixed_content

def parse_xml_root(xml_content):
    """XML文字列をパースし、適切なルート要素とテストスイート名を取得する"""
    try:
        # ルート要素が testsuite か testcases かを判定
        if '<testsuite ' in xml_content[:200]:
            root_element = ET.fromstring(xml_content)
            testsuite_name = root_element.get("name", "")
            testcases_root = root_element # testsuite の下に testcase があると想定
        elif '<testcases>' in xml_content[:200]:
            root_element = ET.fromstring(xml_content)
            testsuite_name = "" # testcases 直下の場合、特定のスイート名はなし
            testcases_root = root_element # testcases の下に testcase があると想定
        else:
             # 想定外の形式の場合、最初の要素から取得を試みる
             root_element = ET.fromstring(xml_content)
             # testsuite または testcases を探す
             possible_root = root_element.find('.//testsuite') or root_element.find('.//testcases') or root_element
             testsuite_name = possible_root.get("name", "") if possible_root.tag == 'testsuite' else ""
             testcases_root = possible_root

        # testcase要素を含む最も適切な要素を返す
        # findall('.//testcase') で検索するので、ルートが testsuite でも testcases でも良い
        return testcases_root, testsuite_name

    except ET.ParseError as pe:
        raise ValueError(f"XMLの解析に失敗しました: {pe}")
    except Exception as e:
        raise Exception(f"XMLルート要素の解析中に予期せぬエラーが発生しました: {e}")


def get_element_text(element, tag_name):
    """指定されたタグの要素テキストを取得する (CDATA対応強化)"""
    tag = element.find(tag_name)
    if tag is not None:
        text = tag.text
        if text:
             text = re.sub(r'<!\[CDATA\[(.*?)\]\]>', r'\1', text, flags=re.DOTALL)
             return text.strip()
    return ""

def clean_html(text):
    """HTMLタグを適切に処理してプレーンテキストに変換する"""
    if not text:
        return ""
    # CDATA除去
    text = re.sub(r'<!\[CDATA\[(.*?)\]\]>', r'\1', text, flags=re.DOTALL)
    # <p> -> 改行
    text = re.sub(r'<p>(.*?)</p>', lambda m: m.group(1).strip() + '\n', text, flags=re.DOTALL | re.IGNORECASE)
    # <br> -> 改行
    text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
    # リスト処理
    def replace_list(match):
        list_content = match.group(1)
        items = re.findall(r'<li.*?>(.*?)</li>', list_content, flags=re.DOTALL | re.IGNORECASE)
        plain_items = []
        for item in items:
            cleaned_item = clean_html(item).strip() # 再帰呼び出しでネストに対応
            lines = [f"・{line.strip()}" for line in cleaned_item.split('\n') if line.strip()]
            plain_items.extend(lines)
        return '\n'.join(plain_items) + '\n' if plain_items else '\n'
    text = re.sub(r'<(ul|ol).*?>(.*?)</\1>', replace_list, text, flags=re.DOTALL | re.IGNORECASE)
    # 残った<li>処理
    text = re.sub(r'<li.*?>(.*?)</li>', lambda m: '・' + clean_html(m.group(1)).strip() + '\n', text, flags=re.DOTALL | re.IGNORECASE)
    # その他タグ除去
    text = re.sub(r'<(?!\/?(p|br|ul|ol|li)\b)[^>]+>', '', text, flags=re.IGNORECASE)
    # HTMLエンティティデコード
    text = text.replace("&nbsp;", " ").replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&").replace("&quot;", "\"").replace("&#39;", "'")
    # 空白・改行整理
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n\s*\n+', '\n', text)
    return text.strip()

def convert_xml_to_csv(testcases_root, testsuite_name, output_csv_file):
    """XML要素ツリーからデータを抽出し、CSVファイルに書き込む"""
    try:
        rows = []
        headers = [
            "ID", "外部ID", "バージョン", "テストケース名", "サマリ（概要）",
            "重要度", "事前条件", "ステップ番号", "アクション（手順）", "期待結果",
            "実行タイプ", "推定実行時間", "ステータス", "有効/無効", "開いているか",
            "親テストスイート名"
        ]
        rows.append(headers)

        # testcases_root (testsuite または testcases 要素) から testcase を検索
        for testcase in testcases_root.findall(".//testcase"):
            testcase_id = testcase.get("internalid", "")
            external_id = get_element_text(testcase, "externalid")
            version = get_element_text(testcase, "version")
            testcase_name = testcase.get("name", "")
            summary = clean_html(get_element_text(testcase, "summary"))
            importance = get_element_text(testcase, "importance")
            preconditions = clean_html(get_element_text(testcase, "preconditions"))
            # execution_type は Testcase直下とStep内にある。ここではTestcase直下のものを取得
            tc_exec_type_elem = testcase.find("execution_type")
            tc_exec_type = tc_exec_type_elem.text.strip() if tc_exec_type_elem is not None and tc_exec_type_elem.text else ""

            exec_duration = get_element_text(testcase, "estimated_exec_duration")
            status = get_element_text(testcase, "status")
            is_active = get_element_text(testcase, "active")
            is_open = get_element_text(testcase, "is_open")

            steps = testcase.find("steps")
            if steps is not None and len(steps) > 0:
                for step in steps.findall("step"):
                    step_number = get_element_text(step, "step_number")
                    actions = clean_html(get_element_text(step, "actions"))
                    expected = clean_html(get_element_text(step, "expectedresults"))
                    step_exec_type_elem = step.find("execution_type")
                    step_exec_type = step_exec_type_elem.text.strip() if step_exec_type_elem is not None and step_exec_type_elem.text else ""

                    row = [
                        testcase_id, external_id, version, testcase_name, summary,
                        importance, preconditions, step_number, actions, expected,
                        step_exec_type, exec_duration, status, is_active, is_open,
                        testsuite_name
                    ]
                    rows.append(row)
            else:
                # ステップがない場合
                row = [
                    testcase_id, external_id, version, testcase_name, summary,
                    importance, preconditions, "", "", "", # ステップ関連は空
                    tc_exec_type, exec_duration, status, is_active, is_open, # 実行タイプはTestcaseのものを採用
                    testsuite_name
                ]
                rows.append(row)

        # CSVファイル書き込み
        with codecs.open(output_csv_file, 'w', 'shift_jis', errors='ignore') as f:
            writer = csv.writer(f, quoting=csv.QUOTE_ALL)
            writer.writerows(rows)

    except Exception as e:
        # ここで発生したエラーは呼び出し元 (main_app) に伝播させる
        raise Exception(f"XMLからCSVへの変換処理中にエラーが発生しました: {str(e)}\n{traceback.format_exc()}")