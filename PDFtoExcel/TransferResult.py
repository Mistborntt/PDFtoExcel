import pandas as pd
import pdfplumber
import re
from datetime import datetime

def extract_data_result(pdf_path, excel_path):
    with pdfplumber.open(pdf_path) as pdf:
        # 首页
        page = pdf.pages[0]

        # 提取文本，分割成行
        text = page.extract_text()
        lines = text.splitlines()

        # 项目名称
        prj_name = ''.join(lines[:3])

        match = re.search(r'(.+转让项目)', prj_name)
        if match:
            prj_name = match.group(1)
        else:
            prj_name = ''

        # 出让方主体名称
        match1 = re.search(r'(.+公司|.+联社)', prj_name)
        if match1:
            mainbody = match1.group(1)
        else:
            mainbody = ''

        # 下属机构
        match2 = re.search(r'公司(.*?)(?=关于|$)', prj_name)
        if match2:
            sub = match2.group(1)
        else:
            sub = ''

        # 读取表格
        table = page.extract_table()

        # 项目编号
        prj_no = table[1][1]

        # 受让方全称
        transferee = table[3][1]

        # 转让协议签署日期
        date = table[4][1]

        input_datetime = datetime.strptime(date, '%Y年%m月%d日')
        date = input_datetime.strftime('%Y-%m-%d')

        # 数据列表
        data_list = [['项目名称', prj_name], ['项目编号', prj_no], ['出让方主体名称', mainbody], ['下属机构', sub],
                     ['受让方全称', transferee], ['转让协议签署日期', date], ['成交价格（元）', '']]

        # 转换成DataFrame
        df = pd.DataFrame(data_list).T
        header = df.iloc[0]
        df = pd.DataFrame(df.values[1:], columns=header)

        # 导出Excel
        df.to_excel(excel_path, index=False)

if __name__ == '__main__':
    pdf_path = r'./PDF样例/转让结果1.pdf'
    excel_path = r'/Users/tt/Desktop/转让结果1.xlsx'

    extract_data_result(pdf_path, excel_path)
