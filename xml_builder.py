import xml.etree.ElementTree as ET
from text_utils import text_to_html

def group_testcases(rows, header_indices):
    """テストケースをIDまたは名前でグループ化する"""
    # インデックスを変数に展開
    id_idx = header_indices.get("ID", -1)
    testcase_name_idx = header_indices["テストケース名"]

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
        
    return testcase_groups

def build_testcase_element(root, testcase_rows, header_indices):
    """テストケース行からXML要素を構築する"""
    if not testcase_rows:
        return
        
    # インデックスを変数に展開
    id_idx = header_indices.get("ID", -1)
    external_id_idx = header_indices.get("外部ID", -1)
    version_idx = header_indices["バージョン"]
    testcase_name_idx = header_indices["テストケース名"]
    summary_idx = header_indices["サマリ（概要）"]
    importance_idx = header_indices["重要度"]
    preconditions_idx = header_indices.get("事前条件", -1)
    step_number_idx = header_indices["ステップ番号"]
    actions_idx = header_indices["アクション（手順）"]
    expected_idx = header_indices["期待結果"]
    exec_type_idx = header_indices["実行タイプ"]
    exec_duration_idx = header_indices.get("推定実行時間", -1)
    status_idx = header_indices.get("ステータス", -1)
    is_active_idx = header_indices.get("有効/無効", -1)
    is_open_idx = header_indices.get("開いているか", -1)
    
    first_row = testcase_rows[0]

    # 必須データの存在チェック
    missing_data = False
    for header in ["テストケース名", "バージョン", "サマリ（概要）", "重要度", "実行タイプ"]:
        idx = header_indices[header]
        # ステップ関連以外は first_row でチェック
        if idx >= len(first_row) or not first_row[idx].strip():
             # ステップ実行タイプはステップ行でチェックする or デフォルト値を使う
             if header == "実行タイプ" and step_number_idx != -1 and first_row[step_number_idx].strip(): # ステップがあれば無視
                  continue
             print(f"警告: 必須データ「{header}」が不足または空です。スキップします。")
             missing_data = True
             break
    if missing_data:
        return

    # <testcase> 要素の属性を設定
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
    preconditions_text = first_row[preconditions_idx].strip() if preconditions_idx != -1 and preconditions_idx < len(first_row) else ""
    preconditions = ET.SubElement(testcase, "preconditions")
    preconditions.text = text_to_html(preconditions_text)

    # <execution_type> (Testcaseレベル)
    tc_exec_type_text = first_row[exec_type_idx].strip() if exec_type_idx < len(first_row) else "1" # デフォルト Manual
    # ステップが存在するかの確認
    has_steps = any(row[step_number_idx].strip() for row in testcase_rows if step_number_idx != -1 and step_number_idx < len(row))
    if not has_steps: # ステップがなければTestcaseレベルの実行タイプを設定
         exec_type_elem = ET.SubElement(testcase, "execution_type")
         exec_type_elem.text = tc_exec_type_text

    # <importance>
    importance_text = first_row[importance_idx].strip() if importance_idx < len(first_row) else "2" # デフォルト Medium
    importance = ET.SubElement(testcase, "importance")
    importance.text = importance_text

    # オプショナル要素 (値があれば追加)
    add_optional_elements(testcase, first_row, header_indices)

    # <steps> 要素
    build_steps_elements(testcase, testcase_rows, header_indices)
    
    return testcase

def add_optional_elements(testcase, row, header_indices):
    """オプショナル要素を追加する"""
    exec_duration_idx = header_indices.get("推定実行時間", -1)
    status_idx = header_indices.get("ステータス", -1)
    is_open_idx = header_indices.get("開いているか", -1)
    is_active_idx = header_indices.get("有効/無効", -1)
    
    if exec_duration_idx != -1 and exec_duration_idx < len(row) and row[exec_duration_idx].strip():
        exec_duration = ET.SubElement(testcase, "estimated_exec_duration")
        exec_duration.text = row[exec_duration_idx].strip()
    if status_idx != -1 and status_idx < len(row) and row[status_idx].strip():
        status = ET.SubElement(testcase, "status")
        status.text = row[status_idx].strip()
    if is_open_idx != -1 and is_open_idx < len(row) and row[is_open_idx].strip():
        is_open = ET.SubElement(testcase, "is_open")
        is_open.text = row[is_open_idx].strip()
    if is_active_idx != -1 and is_active_idx < len(row) and row[is_active_idx].strip():
        active = ET.SubElement(testcase, "active")
        active.text = row[is_active_idx].strip()

def build_steps_elements(testcase, testcase_rows, header_indices):
    """ステップ要素を構築する"""
    step_number_idx = header_indices["ステップ番号"]
    actions_idx = header_indices["アクション（手順）"]
    expected_idx = header_indices["期待結果"]
    exec_type_idx = header_indices["実行タイプ"]
    
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
                 print(f"警告: ステップ番号 {step_number} (CSV行: {step_line_num}) でアクションまたは期待結果が空です。")

            step = ET.SubElement(steps_container, "step")
            step_num_elem = ET.SubElement(step, "step_number")
            step_num_elem.text = step_number
            actions = ET.SubElement(step, "actions")
            actions.text = text_to_html(actions_text)
            expected = ET.SubElement(step, "expectedresults")
            expected.text = text_to_html(expected_text)
            step_exec_type = ET.SubElement(step, "execution_type")
            step_exec_type.text = step_exec_type_text

def create_root_element():
    """ルートのXML要素を作成する"""
    return ET.Element("testcases")