import win32com.client
import os


def xls_to_xlsx(xls_path):
    """
    将xls文件转换成xlsx文件
    Parameters:
        xls_path - xls文件路径
    Returns:
        转换后的xlsx文件路径
    """
    # 初始化Excel应用
    excel = win32com.client.Dispatch('Excel.application')
    # 打开xls文件
    wb = excel.Workbooks.Open(xls_path)
    # 生成xlsx文件路径
    xlsx_path = os.path.splitext(xls_path)[0] + "_new.xlsx"
    # 保存为xlsx格式
    wb.SaveAs(xlsx_path, FileFormat=51)
    # 关闭workbook
    wb.Close()
    # 退出Excel
    excel.Application.Quit()
    return xlsx_path


def batch_convert(xls_folder, xlsx_folder):
    """批量转换文件夹下所有xls为xlsx"""
    # 获取所有xls文件
    xls_files = [f for f in os.listdir(xls_folder) if f.endswith('.xls')]

    # 遍历转换
    for f in xls_files:
        xls_path = os.path.join(xls_folder, f)
        xlsx_path = os.path.join(xlsx_folder, f.replace('.xls', '.xlsx'))
        xls_to_xlsx(xls_path)

    return 0


if __name__ == '__main__':
    xls_folder = r'D:\离线数据\all_xls'
    xlsx_folder = r'D:\离线数据\conv_xlsx'
    batch_convert(xls_folder, xlsx_folder)
