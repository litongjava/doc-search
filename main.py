import os
import argparse
from docx import Document


def search_keyword_in_word(keywords, file_path):
    document = Document(file_path)
    keyword_list = [kw.strip() for kw in keywords.split(',')]
    for paragraph in document.paragraphs:
        for keyword in keyword_list:
            if keyword in paragraph.text:
                return True
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


if __name__ == "__main__":
    main()
