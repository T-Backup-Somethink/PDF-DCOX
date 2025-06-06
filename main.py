import os
import sys
import logging
import argparse
from pdf2docx import Converter  # 正确导入 Converter

def pdf_to_word(pdf_file_path, word_file_path):
    """Convert PDF file to Word file"""
    cv = Converter(pdf_file_path)
    cv.convert(word_file_path)
    cv.close()

def main():
    # 设置日志级别为 ERROR
    logging.getLogger().setLevel(logging.ERROR)

    # 创建命令行参数解析器
    parser = argparse.ArgumentParser(description="Convert a PDF file to a Word document.")
    parser.add_argument("pdf_file", help="The PDF file to be converted.")
    parser.add_argument("word_file", help="The output Word file.")

    # 解析命令行参数
    args = parser.parse_args()

    # 检查源文件是否存在
    if not os.path.exists(args.pdf_file):
        print(f"错误: 源 PDF 文件 '{args.pdf_file}' 不存在。")
        sys.exit(1)

    # 确保目标文件夹存在
    word_folder = os.path.dirname(args.word_file)
    if word_folder and not os.path.exists(word_folder):
        print(f"错误: 目标文件夹 '{word_folder}' 不存在。")
        sys.exit(1)

    # 执行 PDF 转换
    print(f"正在处理: {args.pdf_file}")
    pdf_to_word(args.pdf_file, args.word_file)
    print("转换完成")

if __name__ == "__main__":
    main()
