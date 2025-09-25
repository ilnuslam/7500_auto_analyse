import numpy as np
import re
import xlrd
import tkinter as tk
from tkinter import N, filedialog
import sys
import os
import msvcrt

import horizontal_to_vertical
import sheet_copy_and_delete
import xls_safe_read
import xls_write

def choose_xls_doc():
    #选择文件
    print("please choose document")
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', 1)
    excel_path = filedialog.askopenfilename(
        title="选择qPCR数据表格",
        filetypes=[("excel", "*.xls"), ("所有文件", "*.*")]
    )
    if not excel_path:
        print("未选择文件，程序退出")
        sys.exit(0)
    print(excel_path)
    file_type = os.path.splitext(excel_path)
    return excel_path, file_type 

def get_nonempty_list(nc_num):
    # 删除列表中的空字符串元素后计算均值
    if '' in nc_num:
        print("The list is empty, cannot calculate the average.")
        b_list = []
        for element in nc_num:
            if element:
                b_list.append(element)
        print(b_list)
        level = np.mean(b_list)
        return level
    else:
        # nc均值
        level = np.mean(nc_num)
        return level

def convert_numbers_to_letters(number):
    if not 0 <= number <= 95:
        return None
    
    # 计算行(1-12)和列(A-H)
    row = number // 8 + 1
    column = number % 8
    
    # 将列数字转换为大写字母(A=0, B=1,...H=7)
    if 0 <= column <= 7:
        return f"{chr(65 + column)}{row}"
    else:
        return None

def main():
    #选择Excel文件
    input_source, file_type = choose_xls_doc()
    path = os.path.dirname(input_source)
    file_name = os.path.basename(input_source)

    workbook = xlrd.open_workbook(input_source)        # 打开Excel文件
    sheet = workbook.sheet_by_name('Multicomponent Data')        # 获取工作表
    data = []        # 初始化一个列表来存储D列的数据
        
    # 循环读取C2793到C2888单元格的内容
    for row in range(2792, 2888):
        cell_value = sheet.cell_value(row, 2)
        data.append(cell_value)

    v_data = horizontal_to_vertical.check_dict_empty_string_lists(data)
    identify, read_nc_location, target, (r1, c1, r2, c2) = xls_safe_read.read_excel_range(input_source, 0, 5, 1, 12, 12)

    v_num = len(identify)
    sum_identify = sum(identify, [])
    sum_v_data = sum(v_data, [])

    level = 'N/A'

    if read_nc_location == []:
        print("警告：未找到阴性对照，请检查是否高亮标记阴性对照单元格\a")
    else:
        pr_nc_location = []
        for i in read_nc_location:
            pr_nc_location.append(convert_numbers_to_letters(i))
        print("阴性对照：","、".join(str(item) for item in pr_nc_location))

        #从data检索nc对应数值
        nc_num = [sum_v_data[i] for i in read_nc_location]

        #计算均值，返回均值（如果有空字符串则删除后再计算平均值）
        level = get_nonempty_list(nc_num)
        print("阴性均值：",level)

    ratio = []
    for item in sum_v_data:
        if item == '':
            item = 0

        if level == 'N/A' and item == 0:
            cache_ratio = ''
        elif level == 'N/A' and item != 0:
            cache_ratio  ='N/A'
        else:
            cache_ratio = np.asarray(item, dtype=float) / np.asarray(level, dtype=float)
        ratio.append(str(cache_ratio))

    #print(sum_identify)
    #print(sum_v_data)
    #print(ratio)
    
    xls_write.modify_existing_excel(input_source, target, v_num, sum_identify, sum_v_data, ratio, level, read_nc_location)

    sheet_copy_and_delete.copy_sheet_to_position(input_source, target, v_num + 2)

    print("文件已创建")
    print("按任意键退出...")
    msvcrt.getch()

if __name__ == "__main__":
    main()