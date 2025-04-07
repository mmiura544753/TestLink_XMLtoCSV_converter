import traceback
from csv_reader import read_csv_file, get_header_indices
from xml_builder import group_testcases, build_testcase_element, create_root_element
from xml_utils import element_to_string

def convert_csv_to_xml(csv_file, output_xml_file):
    """CSVファイルを読み込み、TestLinkインポート用のXMLファイルに変換する"""
    try:
        # CSV読み込み
        rows = read_csv_file(csv_file)
        headers = rows[0]
        
        # ヘッダーインデックスの取得
        header_indices = get_header_indices(headers)

        # XMLルート要素 <testcases> を作成
        root = create_root_element()

        # テストケースをグループ化 (IDまたは名前で)
        testcase_groups = group_testcases(rows, header_indices)

        # 各グループからテストケースXML要素を生成
        for group_key, testcase_rows in testcase_groups.items():
            try:
                build_testcase_element(root, testcase_rows, header_indices)
            except Exception as e:
                print(f"警告: テストケース {group_key} の処理中にエラーが発生しました: {str(e)}")
                continue

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