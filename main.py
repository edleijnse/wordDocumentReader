# This is a sample Python script.
from docx2python import docx2python
from collections import Counter

import os

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

def extract_words(filename):
    print("---------------------------------------------------------------------------")
    print(filename)
    print("---------------------------------------------------------------------------")
    doc_result = docx2python(filename).text
    print(doc_result)

    # extract words in a table
    res = doc_result.split()
    return res


def word_document_reader(inputDir):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {inputDir}')  # Press Ctrl+F8 to toggle the breakpoint.
    import os

    total_res = []
    for path, currentDirectory, files in os.walk(inputDir):
        for file in files:
            filename = os.path.join(path, file)
            if ("doc" in filename):
                total_res = total_res + extract_words(filename)
    tallied = Counter(total_res)
    print (tallied)

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    inputDir = r"f:\swissedu_attachments"
    word_document_reader(inputDir)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
