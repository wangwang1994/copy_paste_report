from docx import Document
doc=Document('报告轻型1.docx')
para=doc.paragraphs[0]

print(para.text)
print(para.runs)