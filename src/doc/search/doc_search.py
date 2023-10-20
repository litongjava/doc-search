import argparse
import os
import traceback

import chardet
from docx import Document


def search_keyword_in_word(keywords, file_path):
    try:
        document = Document(file_path)
    except Exception as err:
        stack_trace = traceback.format_exc()
        print(stack_trace)
        return False
    else:
        keyword_list = [kw.strip() for kw in keywords.split(',')]
        for paragraph in document.paragraphs:
            for keyword in keyword_list:
                if keyword in paragraph.text:
                    return True
    return False


def search_keyword_in_txt(keywords, full_file_path):
    try:
        with open(full_file_path, 'rb') as file:
            byte_content = file.read()

            if not byte_content:  # Check if file is empty
                return False

            encoding = chardet.detect(byte_content)['encoding']
            content = byte_content.decode(encoding)

            keyword_list = [kw.strip() for kw in keywords.split(',')]
            for keyword in keyword_list:
                if keyword in content:
                    return True
    except Exception as e:
        stack_trace = traceback.format_exc()
        print(stack_trace)
        return False
    return False


def main():
    parser = argparse.ArgumentParser(description="Search for a keyword in Word documents.")
    parser.add_argument("-p", "--path", required=True, help="The path to the directory containing the Word documents.")
    parser.add_argument("-k", "--keywords", required=True, help="The keywords to search for, separated by commas.")
    args = parser.parse_args()

    source_file_path = args.path
    keywords = args.keywords

    for dirpath, dirnames, filenames in os.walk(source_file_path):
        for filename in filenames:
            full_file_path = os.path.join(dirpath, filename)
            if filename.startswith('~$'):
                continue
            if filename.endswith('.docx'):
                if search_keyword_in_word(keywords, full_file_path):
                    print(full_file_path)
            if filename.endswith('.md'):
                if search_keyword_in_txt(keywords, full_file_path):
                    print(full_file_path)
            if filename.endswith('.txt'):
                if search_keyword_in_txt(keywords, full_file_path):
                    print(full_file_path)
