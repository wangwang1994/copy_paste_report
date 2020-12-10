import os
def get_info(file_name):
    os.chdir(file_name)
    baogaobianhao=open('报告编号.txt')
    baogaobianhao_list=[]
    for line in baogaobianhao.readlines():
        line=line.rstrip('\n')
        baogaobianhao_list.append(line)
    yangpinbianhao = open('样品编号.txt')
    yangpinbianhao_list = []
    for line in yangpinbianhao.readlines():
        line = line.rstrip('\n')
        yangpinbianhao_list.append(line)

    first_file_list=os.listdir()
    first_file_list.remove('报告编号.txt')
    first_file_list.remove('样品编号.txt')
    try:
        first_file_list.remove('.DS_Store')
    except:
        pass
    docx_dict={}
    xlsx_dict={}
    for i in range(len(first_file_list)):
        second_file_name=file_name+'/'+first_file_list[i]
        os.chdir(second_file_name)
        second_file_list=os.listdir()
        for f in second_file_list:
            if f.endswith('.docx'):
                docx_dict[first_file_list[i]]=f
            if f.endswith(('.xlsx')):
                xlsx_dict[first_file_list[i]]=f
    # print(docx_dict)
    # print(xlsx_dict)
        # print(second_file_list)
        # print(second_file_name)
    baogaobianhao_dict={}
    yangpinbianhao_dict={}
    for i in range(len(first_file_list)):
        baogaobianhao_dict[first_file_list[i]]=baogaobianhao_list[i]
    for j in range(len(first_file_list)):
        yangpinbianhao_dict[first_file_list[j]]=yangpinbianhao_list[j]



    return baogaobianhao_dict,yangpinbianhao_dict,docx_dict,xlsx_dict

baogaobianhao_dict,yangpinbianhao_dict,docx_dict,xlsx_dict=get_info('/Users/wangwang/Library/Mobile Documents/com~apple~CloudDocs/Desktop/2/2020.9.17/东风本田港源店')

print(baogaobianhao_dict)
print(yangpinbianhao_dict)
print(docx_dict)
print(xlsx_dict)
print(os.getcwd())
