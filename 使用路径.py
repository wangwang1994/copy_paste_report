import os
import re
import pprint
gongsimingcheng=input('请输入抽检公司名称：')
rule = re.compile('^[a-zA-z]{1}.*$')
filename=input('请输入文件路径：')
os.chdir(filename)
filelist=os.listdir()
def get_docx_xlsx(filepath):
    docx_and_xlsx=[]
    filelist1=os.listdir(filepath)
    for file in filelist1:
        if file.endswith('.docx')==True:
            docx_and_xlsx.append(file)
        if file.endswith('.xlsx')==True:
            docx_and_xlsx.append(file)
    return docx_and_xlsx
# get_docx_xlsx(filename)
# print('查看是否可以得到docx和xlsx文件')
# print(get_docx_xlsx(filename))
print(filelist)
chexing_info={}
print('打印车型名称：')
chexingmingcheng=[]
for file_name in filelist:
    if rule.match(file_name) is not None:
        chexingmingcheng.append(file_name)
print(chexingmingcheng)
for file_name in chexingmingcheng:
    chexing_info[file_name]=[]
print('车型info字典')
print(chexing_info)
print('filelist')
print(filelist)




for second_file_number in range(0,len(chexingmingcheng)):
    #从这里开始进入第二层文件夹，也就是公司文件是filename，而这个secondfile是第二个下面的
    second_file=os.getcwd()+'/'+chexingmingcheng[second_file_number]
    # print('打印二级文件夹名称')
    print(second_file)

    second_file_name=os.listdir(second_file)
    # print(second_file_name)
    print(get_docx_xlsx(second_file))
    chexing_info[chexingmingcheng[second_file_number]]=get_docx_xlsx(second_file)
    # if second_file.endswith('.DS_Store')==False:
    #     second_file_info.append(get_docx_xlsx(second_file))
    #     print(second_file_info)
    #     # chexing_info[chexingmingcheng[second_file_number]]=second_file_info
    #     os.chdir(filename)
pprint.pprint(chexing_info)
print(os.getcwd())

baogaobianhao=open('报告编号.txt')
baogaobianhao_info=[]
for line in baogaobianhao.readlines():
    if line !='':
        print(line)
        line=line.rstrip('\n')
        baogaobianhao_info.append(line)
print(baogaobianhao_info)

chexing_info_bianhao={}
for i in range(len(baogaobianhao_info)):
    chexing_info_bianhao[chexingmingcheng[i]]=baogaobianhao_info[i]
pprint.pprint(chexing_info_bianhao)

# yangpin_kaishi=input('请输入样品开始的编号数如001：')
# print(yangpin_kaishi)
#
# # yangpin_info_bianhao={}
# # for i in range(len(baogaobianhao_info)):
# #     yangpin_info[chexingmingcheng[i]]=baogaobianhao_info[i]
# # pprint.pprint(chexing_info_bianhao)
print('------------')
pprint.pprint(chexing_info)
pprint.pprint(chexing_info_bianhao)
print('------------')
# for i in range(len(chexingmingcheng)):
#     print(chexingmingcheng[i])
#     print(chexing_info[chexingmingcheng[i]])
#     print(chexing_info_bianhao[chexingmingcheng[i]])
#     print("在每个循环中查看docx与xlsx")
#     for item in range(len(chexing_info[chexingmingcheng[i]])):
#         # print(chexing_info[chexingmingcheng[i]][item])
#         if chexing_info[chexingmingcheng[i]][item].endswith('docx'):
#             docx_file=chexing_info[chexingmingcheng[i]][item]
#         if chexing_info[chexingmingcheng[i]][item].endswith('xlsx'):
#             xlsx_file=chexing_info[chexingmingcheng[i]][item]
#         print(docx_file)
#         print(xlsx_file)
for i in range(len(chexingmingcheng)):
    if chexing_info[chexingmingcheng[i]][0].endswith('docx'):
        docx_file=chexing_info[chexingmingcheng[i]][0]
    if chexing_info[chexingmingcheng[i]][0].endswith('xlsx'):
        xlsx_file=chexing_info[chexingmingcheng[i]][0]
    if chexing_info[chexingmingcheng[i]][1].endswith('docx'):
        docx_file = chexing_info[chexingmingcheng[i]][1]
    if chexing_info[chexingmingcheng[i]][1].endswith('xlsx'):
        xlsx_file=chexing_info[chexingmingcheng[i]][1]
    print(docx_file)
    print(xlsx_file)
    baogaobianhao=chexing_info_bianhao[chexingmingcheng[i]]
    print(baogaobianhao)
