import xlrd
import msvcrt

def read_excel_range(file_path, sheet_index=0, start_row=0, start_col=0, end_row=None, end_col=None):
    """
    安全读取Excel数据（自动处理合并单元格和边界）
    :param file_path: Excel文件路径
    :param sheet_index: 工作表索引（默认第一个）
    :param start_row: 起始行（Excel行号-1）
    :param start_col: 起始列（Excel列号-1）
    :param end_row: 结束行（自动适配数据边界）
    :param end_col: 结束列（自动适配数据边界）
    :return: (数据二维列表, 实际读取范围)
    """
    try:
        # 打开工作簿（支持合并单元格需设置formatting_info=True）
        workbook = xlrd.open_workbook(file_path, formatting_info=True)
        sheet = workbook.sheet_by_name('identifier')
        
        # 自动确定数据边界
        max_row = sheet.nrows - 1 if end_row is None else min(end_row, sheet.nrows - 1)
        max_col = sheet.ncols - 1 if end_col is None else min(end_col, sheet.ncols - 1)
        
        # 处理合并单元格
        merged_map = {}
        for (r1, r2, c1, c2) in sheet.merged_cells:
            for row in range(r1, r2):
                for col in range(c1, c2):
                    merged_map[(row, col)] = (r1, c1)  # 映射到合并区域左上角
        
        # 读取数据
        data = []
        nc_location = []
        nc = 0
        for row in range(start_col, max_col + 1):
            row_data = []
            for col in range(start_row, max_row + 1):
                if (row, col) in merged_map:
                    # 如果是合并单元格，读取左上角的值
                    r, c = merged_map[(col, row)]
                    row_data.append(sheet.cell_value(r, c))
                else:
                    row_data.append(sheet.cell_value(col, row))
                    # 检查背景色
                    bg_color = workbook.xf_list[sheet.cell(col, row).xf_index].background.background_colour_index
                    if bg_color != 0x41:
                        nc_location.append(nc)
                nc += 1
            data.append(row_data)
        
        return data, nc_location, (start_row, start_col, max_row, max_col)
    
    except Exception as e:
        print(f"读取错误: {str(e)}")
        print("请检查是否创建了名为 'identifier' 的工作表，且导入样品编号。\a")
        print("按任意键退出...")
        msvcrt.getch()
        return [], [], None

def test():
    # 示例：读取B2到M9区域（行1-8，列1-12）
    data, (r1, c1, r2, c2) = read_excel_range("D:\Desktop\副本20250829CHIKV.xls", 0, 1, 1, 8, 12)
    if data:
        print(f"实际读取: 行{r1+1}-{r2+1} 列{chr(65+c1)}-{chr(65+c2)}")
        print(data)

