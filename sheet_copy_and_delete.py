import xlrd
from xlutils.copy import copy

def copy_sheet_to_position(source_file, target, start_row=0, start_col=9):
    """
    将第一个工作表内容复制到指定工作表的指定位置
    :param source_file: Excel文件路径(.xls)
    :param target: 目标工作表索引
    :param start_row: 目标起始行(默认0对应J1的行)
    :param start_col: 目标起始列(默认9对应J列)
    """
    try:
        # 打开工作簿
        source_file = source_file.replace('.xls', '_auto_analyse.xls')
        book = xlrd.open_workbook(source_file, formatting_info=True)
        source_sheet = book.sheet_by_name('identifier')
        target_sheet = book.sheet_by_name(target)
        
        # 创建可写副本
        wb = copy(book)
        target_sheet = wb.get_sheet(target)

        
        # 复制数据到指定位置
        for row in range(5):
            for col in range(source_sheet.ncols):
                target_sheet.write(
                    row + start_row, 
                    col + start_col, 
                    source_sheet.cell_value(row, col)
                )
        
        # 删除工作表
        sheet_index = book.sheet_names().index('identifier')
        wb._Workbook__worksheets.pop(sheet_index)


        # 保存修改
        wb.save(source_file)
        
    except Exception as e:
        print(f"复制数据时出错: {e}")

if __name__ == "__main__":
    copy_sheet_to_position("example.xls")
