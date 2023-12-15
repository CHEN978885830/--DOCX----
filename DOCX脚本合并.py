import os
from docx import Document

def merge_docx_files(folder_path, output_path):
    # 获取文件夹下所有的docx文件
    docx_files = [f for f in os.listdir(folder_path) if f.endswith('.docx')]

    # 创建一个新的Word文档作为输出
    merged_doc = Document()

    # 遍历所有docx文件并合并内容
    for docx_file in docx_files:
        docx_path = os.path.join(folder_path, docx_file)
        document = Document(docx_path)

        # 遍历原始文档中的每一个段落，并将其添加到合并文档中
        for element in document.element.body:
            merged_doc.element.body.append(element)

    # 保存合并的文档
    merged_doc.save(output_path)

if __name__ == "__main__":
    # 定义输入文件夹和输出文件路径
    input_folder = r"C:\Users\chen\Desktop\剪辑02\心理导师王愚\视频作品"
    output_file = r"C:\Users\chen\Desktop\剪辑02\心理导师王愚\心理导师王愚.docx"

    # 合并DOCX文件
    merge_docx_files(input_folder, output_file)
