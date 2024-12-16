import os
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from openai import OpenAI

# 设置 OpenAI API 密钥
api_key = os.getenv("API_KEY")

# 设置文件夹路径
input_folder = "./files/out"
output_folder = "./files/out/AI"

# 确保输出文件夹存在
os.makedirs(output_folder, exist_ok=True)


def load_prompt_from_txt(file_path):
    """从 txt 文件中加载 Prompt"""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"指定的 Prompt 文件不存在: {file_path}")
    with open(file_path, "r", encoding="utf-8") as f:
        return f.read()


def process_text_with_ai(prompt, text):
    """使用 OpenAI API 处理文本"""
    client = OpenAI(
        api_key=api_key,
        base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
    )

    response = client.chat.completions.create(
        model="qwen-plus",
        messages=[
            {"role": "system", "content": prompt},
            {"role": "user", "content": text}
        ]
    )
    return response.choices[0].message.content


def process_doc_file(file_path, output_path, prompt):
    """处理单个 DOC 文件并保存结果"""
    doc = Document(file_path)

    # 合并所有段落文本
    full_text = "\n".join(paragraph.text.strip() for paragraph in doc.paragraphs if paragraph.text.strip())

    # 调用 AI 处理整个文档文本
    ai_text = process_text_with_ai(prompt, full_text)

    # 将生成的文字写入新文档
    new_doc = Document()
    new_paragraph = new_doc.add_paragraph(ai_text)
    # 设置段落中的字体
    for run in new_paragraph.runs:
        run.font.name = '等线'  # 设置字体为等线
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '等线')  # 设置中文字体为等线
        run.font.size = Pt(11)  # 设置字体大小为 11pt

    # 保存结果
    new_doc.save(output_path)


def process_folder(input_folder, output_folder, prompt):
    """处理文件夹中的所有 DOC 文件"""
    for root, _, files in os.walk(input_folder):
        for file in files:
            if file.endswith(".doc") or file.endswith(".docx"):
                input_path = os.path.join(root, file)
                output_path = os.path.join(output_folder, file)
                print(f"Processing {file}...")
                process_doc_file(input_path, output_path, prompt)
    print("处理完成！")


def main():
    # 获取用户输入的文件夹名称
    input_folder_name = input("请输入需要处理的文件夹名称：")
    input_folder = os.path.join("./files/out/Text/", input_folder_name)
    output_folder = os.path.join("./files/out/AI/", input_folder_name)
    prompt_path = "./files/prompt/prompt.txt"

    # 检查输入文件夹是否存在
    if not os.path.exists(input_folder):
        print(f"错误：文件夹 {input_folder} 不存在！")
        return

    # 确保输出文件夹存在
    os.makedirs(output_folder, exist_ok=True)

    # 设置自定义 prompt
    custom_prompt = load_prompt_from_txt(prompt_path)

    # 处理文件夹
    process_folder(input_folder, output_folder, custom_prompt)


if __name__ == '__main__':
    main()
