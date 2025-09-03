import numpy as np
import re
import xlrd
import tkinter as tk
from tkinter import N, filedialog
import sys
import os

import horizontal_to_vertical
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

def get_integer_input(prompt):
    while True:
        user_input = input(prompt)
        try:
            integer_value = int(user_input)
            #print(f"您输入的整数是：{integer_value}")
            return integer_value
        except ValueError:
            print("输入无效，请输入一个整数。")

def get_nonempty_list(nc_num):
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
        # 使用numpy.mean()函数计算平均值
        level = np.mean(nc_num)
        return level

def convert_letters_to_numbers(letter_list):
    letters_to_numbers = {
    'A': 1,
    'B': 2,
    'C': 3,
    'D': 4,
    'E': 5,
    'F': 6,
    'G': 7,
    'H': 8
}
    return [letters_to_numbers[letter] for letter in letter_list]

def extract_last_one_or_two_digits(s):
    # 尝试提取末尾的两位数字
    if len(s) >= 2 and s[-2:].isdigit():
        return int(s[-2:])
    # 如果末尾没有两位数字，则提取末尾的一位数字
    elif len(s) >= 1 and s[-1].isdigit():
        return int(s[-1])
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
    identify, (r1, c1, r2, c2) = xls_safe_read.read_excel_range(input_source, 0, 1, 1, 8, 12)
    #print(identify)
    v_num = len(identify)
    #print(v_num)
    sum_identify = sum(identify, [])
    sum_v_data = sum(v_data, [])

    nc_count = get_integer_input("请输入阴性对照个数：")
    nc_count = int(nc_count)

    if nc_count == 0:
        level = input("输入阴性均值：").strip()
        nc_level = level
        nc_location = []
        result = [100]
    else:
        nc_exc = []
        for i in range(nc_count):
            element = input(f"请输入第{i+1}个阴性对照格号(如A1)：")
            element = element.strip().upper()
            nc_exc.append(element)
        #根据输入的单元格格式nc_exc转换成列表排序
        letters = [item[0] for item in nc_exc if item[0].isalpha()]
        numbers = [extract_last_one_or_two_digits(s) for s in nc_exc]
        letter_tr = convert_letters_to_numbers(letters)
        result = [8 * (int(y) - 1) + (int(x) - 1) for x, y in zip(letter_tr, numbers)]
        nc_location = result
        #print("阴性对照位置：",nc_location)
    
    if result != [100]:
        nc_level = ""
        #从data检索nc对应数值
        nc_num = [sum_v_data[i] for i in result]

        v_nc_loc = []
        for item in nc_location:
            v_nc_loc.append((item // 12) + 1 + (item % 12) * 8)
        #print(v_nc_loc)

        #计算均值，返回均值（如果有空字符串则删除后再计算平均值）
        level = get_nonempty_list(nc_num)
    print("阴性均值：",level)



    ratio = []
    for item in sum_v_data:
        if item == '':
            item = 0
        cache_ratio = np.asarray(item, dtype=float) / np.asarray(level, dtype=float)
        ratio.append(str(cache_ratio))
    #print(ratio)

    #print(sum_identify)
    #print(sum_v_data)
    xls_write.modify_existing_excel(input_source, v_num, sum_identify, sum_v_data, ratio, level, nc_location)

if __name__ == "__main__":
    main()