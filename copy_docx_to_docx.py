from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from copy import deepcopy
# filename=
doc=Document('报告.docx')
# for table in doc.tables:
#     print(table)
# print(doc.tables[2].cell(0,2).text)
# def copy_table_after(table, paragraph):
#     tbl, p = table._tbl, paragraph._p
#     new_tbl = deepcopy(tbl)
#     p.addnext(new_tbl)



doc.tables[2].cell(0,2).text='报告编号：2342'
paragraph=doc.tables[2].cell(0,2).paragraphs[0]
run=paragraph.runs
font=run[0].font
font.size=Pt(10)
paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

doc.tables[4].cell(0,2).text='报告编号：2342'
paragraph=doc.tables[4].cell(0,2).paragraphs[0]
run=paragraph.runs
font=run[0].font
font.size=Pt(10)
paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

doc.tables[6].cell(0,2).text='报告编号：2342'
paragraph=doc.tables[6].cell(0,2).paragraphs[0]
run=paragraph.runs
font=run[0].font
font.size=Pt(10)
paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

doc.tables[8].cell(0,2).text='报告编号：2342'
paragraph=doc.tables[8].cell(0,2).paragraphs[0]
run=paragraph.runs
font=run[0].font
font.size=Pt(10)
paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

doc.paragraphs[0].text='报告编号：232'
paragraph=doc.paragraphs[0]
run=paragraph.runs
font=run[0].font
font.size=Pt(10)
paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
doc.tables[3].cell(10,1).text="检验结果见第3页"
paragraph=doc.tables[3].cell(10,1).paragraphs[0]
run=paragraph.runs
font=run[0].font
font.size=Pt(10)
doc.tables[3].cell(5,3).text='某某检验员'
paragraph=doc.tables[3].cell(5,3).paragraphs[0]
run=paragraph.runs
font=run[0].font
font.size=Pt(10)
paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.tables[7].cell(4,1).text='--'
paragraph=doc.tables[7].cell(4,1).paragraphs[0]
run=paragraph.runs
font=run[0].font
font.size=Pt(10)
paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.tables[7].cell(4,6).text='--'
paragraph=doc.tables[7].cell(4,6).paragraphs[0]
run=paragraph.runs
font=run[0].font
font.size=Pt(10)
paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
# copy_table_after(doc.tables[0],doc.paragraphs[1])



calid1=doc.tables[7].cell(14,5).text
cvn1=doc.tables[7].cell(14,8).text
calid2=doc.tables[7].cell(16,5).text
cvn2=doc.tables[7].cell(16,8).text




def remove_row(table, row):
    tbl = table._tbl
    tr = row._tr
    tbl.remove(tr)
table=doc.tables[7]
row1 = doc.tables[7].rows[14]
row2 = doc.tables[7].rows[15]
row3= doc.tables[7].rows[16]
row4 = doc.tables[7].rows[17]

remove_row(table, row1)
remove_row(table, row2)
remove_row(table, row3)
remove_row(table, row4)
count=0
text=[]
for par in doc.paragraphs:
    count+=1
    print(par.text)
    print(count)
    text.append(par.text)
print(doc.paragraphs[60].text)
print(doc.paragraphs[61].text)
print(text)
for item in text:
    if item.startswith('检验日期')==True:
        jianyanshijian=item
    if item.startswith('检验地点')==True:
        jianyandidian=item
print(jianyanshijian)
print(jianyandidian)
doc.save('修改编号后.docx')




