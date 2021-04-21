Fama-French 五因素模型说明文档


# 目录结构

财务数据

    按季度合并的资产负债表和利润表

市场回报数据

    按月份合并的月个股回报、综合月市场回报和无风险利率

    rf

        TRD_Cnmont.xlsx     A股市场月回报率文档
        TRD_Nrrate.xlsx     月无风险收益率文档

分组

    因子指标计算部分完成后，按月份生成的分组数据

结果

    result.xls      因子指标计算部分所得的每个月份的 SMB, HML,RMW,CMA 数据

    data.xlsx       线性回归部分所得的用于计算的五因子与因变量


# 代码结构

数据导入处理和因子指标计算部分
线性回归部分


# 运行

run main.py