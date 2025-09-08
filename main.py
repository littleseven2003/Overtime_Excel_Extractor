import pandas as pd

# 读取source.xlsx文件
df = pd.read_excel('source.xlsx')

# 定义研究室选项
research_labs = ['智能通信', '数据算法', '新型能源', '新型材料']

# 提示用户选择研究室
print("请选择研究室：")
for i, lab in enumerate(research_labs, 1):
    print(f"{i}. {lab}")

choice = int(input("请输入选项编号: "))
selected_lab = research_labs[choice - 1]

# 根据选择的研究室过滤数据
filtered_df = df[df['研究室'] == selected_lab]

# 提取姓名、周末加班记录、周末加班奖励总额
# 假设列名分别为 '姓名', '周末加班记录', '周末加班奖励总额'
result_df = filtered_df[['姓名', '周末加班记录', '周末加班奖励总额']]

# 去除“周末加班记录”为空（即内容为[]）的行
result_df = result_df[result_df['周末加班记录'].astype(str) != '[]']

# 整理“奖惩明细”栏内容
import re
def format_overtime_records(record):
    if pd.isna(record) or record == '[]':
        return ''
    # 使用正则表达式提取完整的日期
    dates = re.findall(r'\d{4}-(\d{2})-(\d{2})', record)
    # 格式化日期，去除前导零
    formatted_dates = [f'{int(month)}月{int(day)}日加班' for month, day in dates]
    # 按日期排序
    formatted_dates.sort(key=lambda x: int(x.split('月')[1].split('日')[0]))
    return '\n'.join(formatted_dates)

result_df['奖惩明细'] = result_df['周末加班记录'].apply(format_overtime_records)

# 调整列顺序，将姓名列移到最后
result_df = result_df[['奖惩明细', '周末加班奖励总额', '姓名']]

# 添加序号列
result_df.insert(0, '序号', [1] + [''] * (len(result_df) - 1))

# 添加奖惩项点列
result_df.insert(1, '奖惩项点', ['节假日交通奖励'] + [''] * (len(result_df) - 1))

# 重命名列
result_df.rename(columns={'周末加班奖励总额': '建议奖励金额', '姓名': '涉及人员'}, inplace=True)

# 添加建议处罚金额列
result_df['建议处罚金额'] = ''

# 添加开发室列
result_df['开发室'] = [f'{selected_lab}研究室'] + [''] * (len(result_df) - 1)

# 调整列顺序
result_df = result_df[['序号', '奖惩项点', '奖惩明细', '建议奖励金额', '建议处罚金额', '涉及人员', '开发室']]

# 将结果写入result.xlsx文件
result_df.to_excel('result.xlsx', index=False)

print(f"数据提取完成并写入result.xlsx文件。已选择研究室: {selected_lab}")