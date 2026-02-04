import zipfile
import xml.etree.ElementTree as ET
import os

def docx_to_markdown(file_path, output_path):
    # Namespace for Word XML
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    try:
        with zipfile.ZipFile(file_path, 'r') as docx:
            # Extract document.xml
            content = docx.read('word/document.xml')
            root = ET.fromstring(content)
            
            markdown_content = []
            
            # Find all paragraphs
            for para in root.findall('.//w:p', ns):
                para_text = []
                # Find all text runs in the paragraph
                for run in para.findall('.//w:r', ns):
                    t_elem = run.find('.//w:t', ns)
                    if t_elem is not None and t_elem.text:
                        para_text.append(t_elem.text)
                
                if para_text:
                    markdown_content.append("".join(para_text))
            
            # Write to markdown file
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write("\n\n".join(markdown_content))
            
            print(f"Successfully converted to {output_path}")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    input_file = "국채선물 from 2023 기초금융 교육.v.1.2.docx"
    output_file = "국채선물_기초금융_교육.md"
    docx_to_markdown(input_file, output_file)
