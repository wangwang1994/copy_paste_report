import shutil
import os

def panduan_docx(file_list):
    for file in file_list:
        if file.endswith('.docx')==True:
            if file.endswith('参数确认表.docx')==True:
                pass
            else:
                docx_file=file
    return docx_file
file_name=input('请输入公司文件路径：')
os.chdir(file_name)
chexing_list=os.listdir()
chexing_list.remove('报告编号.txt')
# chexing_list.remove('参数确认表_模版.docx')
# chexing_list.remove('轻型汽油车原始记录模板.xlsx')
# chexing_list.remove('样品编号.txt')
chexing_list.remove('.DS_Store')
# chexing_list.remove('vin列表.txt')
# chexing_list.remove('制造商列表.txt')
# chexing_list.remove('车型列表.txt')
print(chexing_list)

for i in range(len(chexing_list)):
    os.chdir(file_name+'/'+chexing_list[i]+'/'+'报告')
    docx_file=panduan_docx(os.listdir())
    shutil.copy(docx_file,'/Users/wangwang/Desktop/9月16与17日报告')
# shutil.copy('报告.docx','/Users/wangwang/Desktop/未命名文件夹')