import openpyxl
from zipfile import ZipFile
zip = ZipFile('原始记录1.xlsx')
zip.extractall()
wb=openpyxl.load_workbook('原始记录1.xlsx')
ws=wb["随车清单"]
img = openpyxl.drawing.image.Image('xl/media/image1.jpeg')
# img.anchor(ws['A1'])
ws.add_image(img)
wb.save('out.xlsx')