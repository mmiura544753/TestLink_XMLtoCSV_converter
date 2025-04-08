import xml.sax.saxutils as saxutils

def element_to_string(element, indent=""):
    """ElementTreeの要素を整形された文字列に変換（TestLink形式に合わせてCDATA対応）"""
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

    # 常にCDATAで囲むタグ
    html_tags = ['summary', 'preconditions', 'actions', 'expectedresults', 'details']
    # 値があればCDATAで囲むタグ
    cdata_tags = [
        'node_order', 'externalid', 'version', 'step_number', 
        'execution_type', 'importance', 'status', 
        'is_open', 'active', 'name', 'value'
    ]

    if text_content is not None:
        # テキストをエスケープするかCDATAで囲む
        stripped_text = text_content # strip() しないで元のテキストを保持
        if stripped_text:
            if tag in html_tags:
                # HTML要素はCDATAで囲む
                escaped_text = stripped_text.replace(']]>', ']]]]><![CDATA[>')
                result += f"<![CDATA[{escaped_text}]]>"
            elif tag in cdata_tags:
                # 通常値もCDATAで囲む（TestLink形式に合わせる）
                escaped_text = stripped_text.replace(']]>', ']]]]><![CDATA[>')
                result += f"<![CDATA[{escaped_text}]]>"
            else:
                # それ以外のテキストはXMLエスケープ
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