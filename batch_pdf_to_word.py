import os
from pdf2docx import Converter
from tkinter import Tk
from tkinter.filedialog import askdirectory
import sys


def convert_pdf_to_word(pdf_path, word_path):
    try:
        cv = Converter(pdf_path)
        cv.convert(word_path, start=0, end=None)
        cv.close()
        print(f"转换成功: {word_path}")
    except Exception as e:
        print(f"转换失败: {pdf_path} -> {e}")


if __name__ == "__main__":
    Tk().withdraw()
    source_folder = askdirectory(title="选择包含 PDF 文件的文件夹")
    if not source_folder:
        print("未选择源文件夹。程序退出...")
        sys.exit()

    target_folder = os.path.join(os.path.dirname(source_folder), "输出文档")
    os.makedirs(target_folder, exist_ok=True)
    print(f"输出文件夹已创建: {target_folder}")

    for file_name in os.listdir(source_folder):
        if file_name.lower().endswith(".pdf"):
            pdf_path = os.path.join(source_folder, file_name)
            word_file_name = os.path.splitext(file_name)[0] + ".docx"
            word_path = os.path.join(target_folder, word_file_name)

            convert_pdf_to_word(pdf_path, word_path)

    print("所有文件转换完成。")
