#-*- coding:utf-8 -*-

import os
from pdf2image import convert_from_path

import re
from io import StringIO
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser

# name of jpg
output_string = StringIO()
strings_page = []
with open('input.pdf', 'rb') as in_file:
    parser = PDFParser(in_file)
    doc = PDFDocument(parser)
    rsrcmgr = PDFResourceManager()
    device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    for page in PDFPage.create_pages(doc):
        interpreter.process_page(page)

        strings_page.append(re.sub('\n\n', '\n', output_string.getvalue().strip()))

        output_string.seek(0)
        output_string.truncate(0)


# pdf to jpg
file_name = "input.pdf"
pages = convert_from_path(file_name)
os.makedirs('./JPGs', exist_ok=True)
for i, page in enumerate(pages):
    strings = strings_page[i].split('\n')
    strings[3] = strings[3].replace('/', '.')
    strings[3] = strings[3].replace(':', '.')
    page.save("./JPGs/" + strings[3] +".jpg", "JPEG")
