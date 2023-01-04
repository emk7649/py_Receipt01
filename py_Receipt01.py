#-*- coding:utf-8 -*-

import os
from pdf2image import convert_from_path

file_name = "input.pdf"

pages = convert_from_path(file_name)

os.makedirs('./JPGs', exist_ok=True)

for i, page in enumerate(pages):
	page.save("./JPGs/" + file_name+str(i + 1)+".jpg", "JPEG")
