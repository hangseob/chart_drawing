import zipfile
import xml.etree.ElementTree as ET

def docx_to_markdown(file_path, output_path):
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    try:
        with zipfile.ZipFile(file_path, 'r') as docx:
            content = docx.read('word/document.xml')
            root = ET.fromstring(content)
            markdown_content = []
            for para in root.findall('.//w:p', ns):
                para_text = []
                for run in para.findall('.//w:r', ns):
                    t_elem = run.find('.//w:t', ns)
                    if t_elem is not None and t_elem.text:
                        para_text.append(t_elem.text)
                if para_text:
                    markdown_content.append("".join(para_text))
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write("\n\n".join(markdown_content))
            return True
    except Exception as e:
        print(f"Error: {e}")
        return False

if __name__ == "__main__":
    docx_to_markdown("짜투리_국채선물 이론가의 가격 결정 요소.docx", "국채선물_이론가_결정요소.md")
