import os
from docx import Document
from docx.shared import Pt  # 用于字体大小设置
from docx.oxml.ns import qn  # 用于中文字体支持
import re


def split_document_by_timestamp(input_file):
    doc = Document(input_file)

    base_name = os.path.splitext(os.path.basename(input_file))[0]

    output_dir = os.path.join('./files/out', base_name)
    os.makedirs(output_dir, exist_ok=True)

    segments = []
    current_segment = ""

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        # 匹配非数字开头加时间戳
        if re.match(r'^\D+\d{2}:\d{2}:\d{2}', text):
            if current_segment:
                segments.append(current_segment)
            current_segment = text + "\n"
        else:
            current_segment += text + "\n"

    if current_segment:
        segments.append(current_segment)

    current_text = ""
    file_counter = 1

    for segment in segments:
        if len(current_text) + len(segment) > 8000:
            if current_text:
                output_file = os.path.join(output_dir, f"{base_name}_{file_counter}.docx")
                save_document(current_text, output_file)
                file_counter += 1
                current_text = segment
        else:
            current_text += segment

    if current_text:
        output_file = os.path.join(output_dir, f"{base_name}_{file_counter}.docx")
        save_document(current_text, output_file)


def save_document(text, output_file):
    doc = Document()
    for line in text.split("\n"):
        paragraph = doc.add_paragraph(line)

        # 设置字体为等线
        for run in paragraph.runs:
            run.font.name = "Microsoft YaHei"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
            run.font.size = Pt(11)
    doc.save(output_file)


def process_files():
    input_dir = './files/LiveStreamerText'

    # Process each .docx file in the input directory
    for filename in os.listdir(input_dir):
        if filename.endswith('.docx'):
            input_file = os.path.join(input_dir, filename)
            split_document_by_timestamp(input_file)


if __name__ == "__main__":
    process_files()
