# 内容介绍

这里抓取个贷或信用卡转让公告PDF中不良贷款基本信息、时间安排和转让方式的数据，以及转让结果PDF中的项目名称等数据，并以Excel的形式导出。

代码使用Python语言，用到了pandas、pdfplumber、re和datetime库，其中pandas和pdfplumber库需要pip install下载。

其中：

'TransferNotice.py'是转让公告PDF数据的获取，封装了extract_data_notice函数，获取的字段包括：项目名称、出让方主体名称、下属机构、交易基准日、资产笔数（笔）、加权平均逾期天数、单一借款人最高未偿本息余额（元）、借款人加权平均年龄、未偿本金总额（元）、未偿利息总额（元）、未偿本息总额（元）、其他费用（元）、借款人平均未偿本息余额（元）、借款人加权平均授信额度（元）、次级（笔）、可疑（笔）、损失（笔）、已核销（笔）、信用（笔）、保证（笔）、未诉（笔）、诉讼中（笔）、已诉讼（笔）、仲裁中（笔）、已仲裁（笔）、已判未执（笔）'、执行中（笔）、撤回执行（笔）、执行中止（笔）、终结执行（笔）、终本执行（笔）、已调解（笔）、其他（笔）、备注、竞价报名截止时间、竞价日、转让方式、竞价方式、自由竞价开始时间、自由竞价结束时间、延时周期（分钟）、起始价（元）和加价幅度（元）。其中已诉讼（笔）、仲裁中（笔）、已仲裁（笔）、撤回执行（笔）和执行中止（笔）均为空，主要是早期转让公告格式不统一所致，但是字段予以保留；转让方式均为线上公开竞价、竞价方式均为多轮竞价。

'TransferResult.py'是转让结果PDF数据的获取，封装了extract_data_result函数，获取的字段包括：项目名称、项目编号、出让方主体名称、下属机构、受让方全称、转让协议签署日期和成交价格（元），其中成交价格为空。

'Main.py'用来运行以上两个函数。

PDF样例各挑选了近日的2份转让公告和转让结果，可用来测试运行。

导出数据的格式具体可见'2023年个人不良贷款转让公告及转让结果.xlsx'。

后续看看能不能批量处理或者加入爬虫。
