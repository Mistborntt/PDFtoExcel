import pandas as pd
import pdfplumber
import os
import re
from datetime import datetime

def extract_data_notice(pdf_path, excel_path):
    with pdfplumber.open(pdf_path) as pdf:
        # 所有页面
        pages = pdf.pages

        # 第一页
        page1 = pages[0]

        # 第二页
        page2 = pages[1]

        # 提取文本，分割成行
        text = page1.extract_text() + page2.extract_text()
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

        # 读取不良贷款基本信息表格
        table1 = page1.extract_tables()

        # 交易基准日
        benchmark_date = table1[0][0][1]
        benchmark_date = re.sub(r'(\d{4})(\d{2})(\d{2})', r'\1-\2-\3', benchmark_date)

        # 资产笔数（笔）
        asset_amount = table1[0][1][1]

        # 借款人户数（户）
        borrower_num = table1[0][2][1]

        # 加权平均逾期天数
        overdue_days = table1[0][3][1]

        # 单一借款人最高未偿本息余额（元）
        max_principal_interest = table1[0][4][1]

        # 借款人加权平均年龄
        avg_age = table1[0][5][1]

        # 未偿本金总额（元）
        principal = table1[0][0][3]

        # 未偿利息总额（元）
        interest = table1[0][1][3]

        # 未偿本息总额（元）
        principal_interest = table1[0][2][3]

        # 其他费用（元）
        other_fees = table1[0][3][3]

        # 借款人平均未偿本息余额（元）
        avg_principal_interest = table1[0][4][3]

        # 借款人加权平均授信额度（元）
        avg_credit_line = table1[0][5][3]

        # 五级分类情况，次级、可疑、损失
        five_level_cls = table1[0][6][1]

        match3 = re.search(r'次级(.*?)(?=笔|$)', five_level_cls)
        if match3:
            secondary = match3.group(1)
        else:
            secondary = ''

        match4 = re.search(r'可疑(.*?)(?=笔|$)', five_level_cls)
        if match4:
            doubt = match4.group(1)
        else:
            doubt = ''
        match5 = re.search(r'损失(.*?)(?=笔|$)', five_level_cls)

        if match5:
            loss = match5.group(1)
        else:
            loss = ''

        # 核销情况，已核销
        verification = table1[0][7][1]

        match6 = re.search(r'已核销(.*?)(?=笔|$)', verification)
        if match6:
            verified = match6.group(1)
        else:
            verified = ''

        # 担保情况，信用、保证
        guarantee_situation = table1[0][8][1]

        match7 = re.search(r'信用(.*?)(?=笔|$)', guarantee_situation)
        if match7:
            credit = match7.group(1)
        else:
            credit = ''

        match8 = re.search(r'保证(.*?)(?=笔|$)', guarantee_situation)
        if match8:
            guarantee = match8.group(1)
        else:
            guarantee = ''

        # 诉讼情况概述，未诉、诉讼中、已诉讼、仲裁中、已仲裁、已判未执、执行中、撤回执行、执行中止、终结执行、终本执行、已调解、其他
        litigation = table1[0][9][1]

        match9 = re.search(r'未诉(.*?)(?=笔|$)', litigation)
        if match9:
            not_sued = match9.group(1)
        else:
            not_sued = ''

        match10 = re.search(r'诉讼中(.*?)(?=笔|$)', litigation)
        if match10:
            proceeding = match10.group(1)
        else:
            proceeding = ''

        match11 = re.search(r'已判未执(.*?)(?=笔|$)', litigation)
        if match11:
            determined_not_executed = match11.group(1)
        else:
            determined_not_executed = ''

        match12 = re.search(r'执行中(.*?)(?=笔|$)', litigation)
        if match12:
            execute = match12.group(1)
        else:
            execute = ''

        match13 = re.search(r'终结执行(.*?)(?=笔|$)', litigation)
        if match13:
            termination_of_execution = match13.group(1)
        else:
            termination_of_execution = ''

        match14 = re.search(r'终本执行(.*?)(?=笔|$)', litigation)
        if match14:
            final_execution = match14.group(1)
        else:
            final_execution = ''

        match15 = re.search(r'已调解(.*?)(?=笔|$)', litigation)
        if match15:
            mediate = match15.group(1)
        else:
            mediate = ''

        match16 = re.search(r'其他(.*?)(?=笔|$)', litigation)
        if match16:
            other = match16.group(1)
        else:
            other = ''

        # 备注
        notes = ''
        if len(table1[0]) == 11:
            notes = table1[0][10][1]

        # 竞价报名截止时间
        deadline = [item for item in lines if '竞价报名截止时间' in item][0]
        match = re.search(r'竞价报名截止时间：(.*)', deadline)
        if match:
            deadline = match.group(1)
        else:
            deadline = ''

        input_datetime = datetime.strptime(deadline, '%Y年%m月%d日 %H:%M')
        deadline = input_datetime.strftime('%Y-%m-%d %H:%M')

        # 读取转让方式表格
        table2 = page2.extract_tables()

        # 自由竞价开始时间
        start_time = table2[0][0][1]

        # 自由竞价结束时间
        end_time = table2[0][1][1]

        # 延时周期
        delay_period = table2[0][2][1]

        # 起始价
        starting_price = table2[0][3][1]

        # 加价幅度
        price_increase_range = table2[0][4][1]
        
        # 上面是跨页表格，以下不跨页表格  
        if len(table2[0]) == 6:
            start_time = table2[0][1][1]

            end_time = table2[0][2][1]

            delay_period = table2[0][3][1]

            starting_price = table2[0][4][1]

            price_increase_range = table2[0][5][1]

        # 数据列表
        data_list = [['项目名称', prj_name], ['出让方主体名称', mainbody], ['下属机构', sub], ['交易基准日', benchmark_date], ['资产笔数（笔）', asset_amount],
                     ['借款人户数', borrower_num], ['加权平均逾期天数', overdue_days], ['单一借款人最高未偿本息余额', max_principal_interest], ['借款人加权平均年龄', avg_age],
                     ['未偿本金总额（元）', principal], ['未偿利息总额（元）', interest], ['未偿本息总额（元）', principal_interest], ['其他费用（元）', other_fees],
                     ['借款人平均未偿本息余额（元）', avg_principal_interest], ['借款人加权平均授信额度（元）', avg_credit_line], ['次级（笔）', secondary],
                     ['可疑（笔）', doubt], ['损失（笔）', loss], ['已核销（笔）', verified], ['信用（笔）', credit], ['保证（笔）', guarantee],
                     ['未诉（笔）', not_sued], ['诉讼中（笔）', proceeding], ['已诉讼（笔）', ''], ['仲裁中（笔）', ''], ['已仲裁（笔）', ''],
                     ['已判未执（笔）', determined_not_executed], ['执行中（笔）', execute], ['撤回执行（笔）', ''], ['执行中止（笔）', ''],
                     ['终结执行（笔）', termination_of_execution], ['终本执行（笔）', final_execution], ['已调解（笔）', mediate], ['其他（笔）', other],
                     ['备注', notes], ['竞价报名截止时间', deadline], ['竞价日', start_time], ['转让方式', '线上公开竞价'], ['竞价方式', '多轮竞价'],
                     ['自由竞价开始时间', start_time], ['自由竞价结束时间', end_time], ['延时周期（分钟）', delay_period], ['起始价（元）', starting_price],
                     ['加价幅度（元）', price_increase_range]]

        # 转换成DataFrame
        df = pd.DataFrame(data_list).T
        header = df.iloc[0]
        df = pd.DataFrame(df.values[1:], columns=header)

        # 导出Excel
        df.to_excel(excel_path, index=False)

if __name__ == '__main__':
    # 将当前工作目录设置为脚本所在的目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    pdf_path = r'./PDF样例/转让公告1.pdf'
    excel_path = r'/Users/tt/Desktop/转让公告1.xlsx'

    extract_data_notice(pdf_path, excel_path)
