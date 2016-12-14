import xlrd
import xlsxwriter
import re


data_path = 'E:\\dataset\\clinic\\胃癌（D2R0术后患者，共1476例）.xlsx'
pathology_path = 'E:\\dataset\\clinic\\pathology.xlsx'
pathology_path1 = 'E:\\dataset\\clinic\\pathology1.xlsx'
pathology_path2 = 'E:\\dataset\\clinic\\pathology2.xlsx'


# 将病例诊断以<淋巴结>分割，分别写进两列
def get_pathology(r_path, w_path):
    gc_data = xlrd.open_workbook(r_path)
    # 通过索引获取工作表
    gc_table = gc_data.sheets()[2]
    # 创建文件
    pathology_file = xlsxwriter.Workbook(w_path)
    # 创建工作表
    pathology_sheet = pathology_file.add_worksheet()
    for i in range(1, gc_table.nrows):
        row_data = gc_table.row_values(i)
        pathology_data = row_data[2]
        info = pathology_data.split('<淋巴结>')
        pathology_sheet.write(i - 1, 0, info[0])
        pathology_sheet.write(i - 1, 1, info[1])
    pathology_file.close()


# get_pathology(data_path, pathology_path)


# 对第一列的胃部信息做处理(删除参见报告)
def del_some(r_path, w_path):
    gc_data = xlrd.open_workbook(r_path)
    gc_table = gc_data.sheets()[1]
    pathology_file = xlsxwriter.Workbook(w_path)
    pathology_sheet = pathology_file.add_worksheet()
    for i in range(gc_table.nrows):
        row_data = gc_table.row_values(i)
        # 胃部信息
        g_info = row_data[0].strip().replace('\n', '').replace('\r\n', '')
        # 淋巴信息
        lb_info = row_data[1].strip().replace('\n', '').replace('\r\n', '')
        # 获取‘再发报告’部分
        zfbg = re.search(r'(.*报告：)', g_info)
        # 除去该部分
        if zfbg:
            zfbg = zfbg.group()
            start = g_info.index(zfbg)
            end = start + len(zfbg)
            g_info = g_info[0:start] + g_info[end:len(g_info)]
        # 获取参见病理诊断报告部分。
        bg = re.search(r'(。(\D*?)参见(.*?)(报告。|。))', g_info)
        # 除去参见病理诊断报告这一部分
        if bg:
            # print(bg.group())
            bg = bg.group()
            print(bg)
            start = g_info.index(bg)
            end = start + len(bg)
            # print(g_info[0:start] + g_info[end:len(g_info)])
            g_info_new = g_info[0:start] + '。' + g_info[end:len(g_info)]
            pathology_sheet.write(i, 0, g_info_new)
        else:
            pathology_sheet.write(i, 0, g_info)
        pathology_sheet.write(i, 1, lb_info)


# del_some(pathology_path, pathology_path1)


def del_some_2(r_path, w_path):
    gc_data = xlrd.open_workbook(r_path)
    gc_table = gc_data.sheets()[0]
    pathology_file = xlsxwriter.Workbook(w_path)
    pathology_sheet = pathology_file.add_worksheet()
    for i in range(gc_table.nrows):
        row_data = gc_table.row_values(i)
        # 胃部信息
        g_info = row_data[0].strip().replace('\n', '').replace('\r\n', '')
        # 淋巴信息
        lb_info = row_data[1].strip().replace('\n', '').replace('\r\n', '')

        # 获取……未见癌累及部分。
        alj = re.search(r'((。送检|。标本|送检|标本|。“|。<|“)(.*?)(未见|未查见)癌累及。)', g_info)
        # 除去参见病理诊断报告这一部分
        if alj:
            # print(bg.group())
            alj = alj.group()
            print(alj)
            start = g_info.index(alj)
            end = start + len(alj)
            # print(g_info[0:start] + g_info[end:len(g_info)])
            g_info_new = g_info[0:start] + g_info[end:len(g_info)]
            pathology_sheet.write(i, 0, g_info_new)
        else:
            pathology_sheet.write(i, 0, g_info)
        pathology_sheet.write(i, 1, lb_info)


del_some_2(pathology_path1, pathology_path2)
