import xml.sax.saxutils as saxutils

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