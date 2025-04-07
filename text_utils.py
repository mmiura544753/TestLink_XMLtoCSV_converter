import xml.sax.saxutils as saxutils

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