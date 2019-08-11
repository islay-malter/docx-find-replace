#!/bin/python3

import os
import glob

#3rd party library for opening DOCX files
import docx

searchWord = "bullshit"
replaceWord = "cowshit"

# enumerate all DOCX files in current directory
for filename in glob.glob(os.path.join(os.getcwd(), '*.docx')):

    # open file
    testDoc = docx.Document(filename)

    # for every paragraph in DOCX file, find & replace + save
    for p in testDoc.paragraphs:
        p.text = p.text.replace(searchWord, replaceWord)
    testDoc.save(filename)