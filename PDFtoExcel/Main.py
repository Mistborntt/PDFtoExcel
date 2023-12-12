from TransferNotice import extract_data_notice
from TransferResult import extract_data_result

# 将当前工作目录设置为脚本所在的目录
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

pdf_path1 = r'./PDF样例/转让公告1.pdf'
excel_path1 = r'/Users/tt/Desktop/转让公告1.xlsx'

pdf_path2 = r'./PDF样例/转让结果1.pdf'
excel_path2 = r'/Users/tt/Desktop/转让结果1.xlsx'

extract_data_notice(pdf_path1, excel_path1)
extract_data_result(pdf_path2, excel_path2)
