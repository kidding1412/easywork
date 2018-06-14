# -*- coding: utf-8 -*-
import random
import re
# import docx
from docx import Document

# 测试docx demo
# 源文件 test.docx
#doc1 = docx.opendocx()
doc = Document(r"/Users/Kidding/Documents/test.docx")
#print doc.name
# 正则匹配所有一位小数
p = re.compile("^[0-9]+\.[0-9]{1}$")
#docx.table([0])
# 遍历所有表格
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            if re.search(p, cell.text):
                # 匹配小数并在后面补齐一位小数
                cell.text += str(random.randint(0, 9))
doc.save(r"/Users/Kidding/Documents/test1.docx")
