import numpy as np
import xlrd
import xlwt
from xlutils.copy import copy

def modify_existing_excel(file_path, target, v_num, data_list1, data_list2, data_list3, data_list4, data_list5):

    # 打开现有文件(保留格式)
    rb = xlrd.open_workbook(file_path, formatting_info=True)
    # 创建可写副本
    wb = copy(rb)
    try:
        # 获取工作表
        for i in range(wb._Workbook__worksheets.__len__()):
            sheet = wb._Workbook__worksheets[i]
            if sheet.name == 'identifier':
                sheet_num = sheet
    
        new_sheet = wb.add_sheet(target)
    except Exception as e:
        print(f"错误: {e}")
        return

    # 定义高亮样式
    highlight_yellow = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
    highlight_red = xlwt.easyxf('pattern: pattern solid, fore_colour red;')
    highlight_green = xlwt.easyxf('pattern: pattern solid, fore_colour green;')

    # 创建边框样式(包含斜线)
    style = xlwt.XFStyle()

    # 设置边框(需要通过位运算组合边框属性)
    borders = xlwt.Borders()
    #borders.left = xlwt.Borders.THIN
    #borders.right = xlwt.Borders.THIN
    #borders.top = xlwt.Borders.THIN
    #borders.bottom = xlwt.Borders.THIN
    borders.diag = xlwt.Borders.MEDIUM  # 斜线样式
    borders.diag_color = 0x40  # 斜线颜色
    
    # 设置斜线方向(左下到右上)
    borders.left_line_style = xlwt.Borders.MEDIUM
    borders.right_line_style = xlwt.Borders.MEDIUM
    borders.diag_line_style = xlwt.Borders.MEDIUM
    borders.need_diag1 = 0  # 左下到右上
    borders.need_diag2 = 1  # 左上到右下
    
    style.borders = borders
    
    # 竖排写入编号，NC填充背景色
    nc_lo = 0
    for i, value in enumerate(data_list1):
        if nc_lo in data_list5:
            new_sheet.write(13 + i, 0, value, highlight_yellow)
        else:
            new_sheet.write(13 + i, 0, value)
        nc_lo += 1

    # 写入编号表格
    nc_lo = 0
    for col in range(v_num + 2, (v_num * 2 + 2)):
        for row in range(1, 9):
            new_sheet.write(row + 4, col + 1, data_list1[nc_lo]) if data_list1[nc_lo] != '' else new_sheet.write(row + 4, col + 1, '', style)
            nc_lo += 1
    
    # 竖排写入数值
    for i, value in enumerate(data_list2):
        new_sheet.write(13 + i, 1, value)

    # 写入数值表格
    data_lo = 0
    for col in range(1, (v_num +1)):
        for row in range(1, 9):
            new_sheet.write(row, col, data_list2[data_lo]) if data_list2[data_lo] != '' else new_sheet.write(row, col, '', style)
            data_lo += 1

    # 写入表头
    for col in range(1, v_num +1):
        new_sheet.write(0, col, col)
    cha = 65
    for row in range(1, 9):
        new_sheet.write(row, 0, chr(cha))
        new_sheet.write(row + 4, v_num + 2, chr(cha))
        cha += 1

    # 竖排写入比值
    if data_list5 != []:
        data_list3 = np.asarray(data_list3, dtype=float)
        for i, value in enumerate(data_list3):
            if value < 2 and value >=1.5:
                new_sheet.write(13 + i, 2, value, highlight_green)
            elif value >=2:
                new_sheet.write(13 + i, 2, value, highlight_red)
            elif value:
                new_sheet.write(13 + i, 2, value)
    elif data_list5 == []:
        for i, value in enumerate(data_list3):
            new_sheet.write(13 + i, 2, 'N/A')
    else:
        for i, value in enumerate(data_list3):
            new_sheet.write(13 + i, 2, "error")

    new_sheet.write(12, 1, data_list4)

    new_sheet.write(11, 0, "样本")
    new_sheet.write(12, 0, "对照均值")
    new_sheet.write(11, 1, "荧光值")
    new_sheet.write(11, 2, "比值")
    new_sheet.write(12, 2, 1)
    new_sheet.write(11, 3, "可见荧光")
    new_sheet.write(11, 4, "备注")
    # 保存新文件
    new_file = file_path.replace('.xls', '_auto_analyse.xls')
    #wb.save(new_file)
    try:
        wb.save(new_file)
    except Exception as e:
        print(f"保存失败: {e}")
    return new_file

def test():
    if __name__ == "__main__":
        # 准备测试数据
        test_data = [['GZ-HP\n0816', 'GZ-SMU\n0816', 'GZ-XJ\n0816', 'GZ-TJ\n0816', 'FS-WT-13Egg⑥', 'FS-WT-13Egg③', 'RY475', 'RY476'], ['RY477', 'RY478', 'RY479', 'RY480', 'RY481', 'RY482', 'RY483', 'RY484'], ['RY485', 'RY486', 'RY487', 'RY488', 'RY489', 'RY490', 'RY491', 'RY492'], ['RY493', 'RY494', 'RY495', 'RY496', 'RY497', 'FS-WT-56④', 'FS-WT-56②', 'SMURY1'], ['SMURY2', 'SMURY3', 'SMURY4', 'SMURY5', 'SMURY6', 'SMURY7', 'SMURY8', 'LC-6\n0808'], ['LC-6\n0809', 'GA-1\n0811', 'LC-6\n0812', 'FS-GA\n0818', 'FS-LC\n0818', 'FS-TC\n0818', 'FS-XL\n0818', 'FS-LC\n0819']]
        test_data = sum(test_data, [])
        start_row = 0
        start_col = 0
        # 修改现有文件(假设已存在test.xls)
        try:
            output_file = modify_existing_excel("D:\Desktop\副本20250829CHIKV.xls", test_data, start_row, start_col)
            print(f"数据已写入修改后的文件: {output_file}")
        except FileNotFoundError:
            print("错误: 请先创建test.xls文件用于测试")
