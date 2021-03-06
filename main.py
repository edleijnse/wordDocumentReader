# This is a sample Python script.
from docx2python import docx2python
from collections import Counter

import os
import pandas as pd

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

def extract_words(filename, outputFile):
    print("---------------------------------------------------------------------------")
    outputFile.write("\n------------------------------------------------------------------------------------------------\n")
    print(filename)
    outputFile.write(filename+"\n")
    print("---------------------------------------------------------------------------")
    outputFile.write("------------------------------------------------------------------------------------------------\n")
    doc_result = docx2python(filename).text
    print(doc_result)
    outputFile.writelines(doc_result)


    # extract words in a table
    res = doc_result.split()
    return res
def print_tallied(sorted_tallied, outputFile, title):
    #print(sorted_tallied)
    outputFile.write(
        "\n----------------------------------------------------------------------------------------------\n")
    outputFile.write(title)
    outputFile.write(
        "\n----------------------------------------------------------------------------------------------\n")
    ii = 0
    for item in sorted_tallied:
        if (item[1]>0):
           if (item[0].isalpha()):
              print(str(item[0]) + ":" + str(item[1]))
              outputFile.write(str(item[0]) + ":" + str(item[1]) + "\n")
              ii = ii + 1
    print("total: " + str(ii))

    outputFile.write("total: " + str(ii))

def sort_panda(tallied, outputFile):
    outputFile.write(
        "\n----------------------------------------------------------------------------------------------\n")
    outputFile.write("summary words by frequency")
    outputFile.write(
        "\n----------------------------------------------------------------------------------------------\n")

    ii = 0
    wordTab = []
    countTab = []

    for item in tallied:
        if (item[1] > 0):
            if (item[0].isalpha()):
                wordTab.append(item[0])
                countTab.append(item[1])
                print(str(item[0]) + ":" + str(item[1]))
                ii = ii + 1
    data = {'Count': countTab, 'Word': wordTab}
    df = pd.DataFrame(data, columns=['Count', 'Word'])
    df.sort_values(by=['Count','Word'], inplace=True)
    for oneValue in df.values:
        print(str(oneValue[0]) + " " + oneValue[1])
        outputFile.write(str(oneValue[0]) + " " + oneValue[1] + "\n")

def word_document_reader(inputDir, outputFile):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {inputDir}')  # Press Ctrl+F8 to toggle the breakpoint.
    import os

    total_res = []
    for path, currentDirectory, files in os.walk(inputDir):
        for file in files:
            filename = os.path.join(path, file)
            if ("doc" in filename):
                total_res = total_res + extract_words(filename, outputFile)
    tallied = Counter(total_res)
    sorted_tallied = sorted(tallied.items())
    print_tallied(sorted_tallied, outputFile, "summary of words")
    sort_panda(tallied.items(),outputFile)

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    inputDir = r"f:\swissedu_attachments"
    outputFile = open(r"f:\swissedu.txt", "w", encoding="utf-8")


    word_document_reader(inputDir, outputFile)

    outputFile.close()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
