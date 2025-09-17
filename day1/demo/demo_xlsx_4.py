# python之xlsxwriter模块——实现对Excel文件中的内容进行关键词标红
"""
在日常使用Excel进行数据处理时，有时需要在Excel文件中对单元格内的特定字符进行标红，
以便更好的筛选需要的数据，针对这种情况python提供了一种好用的库xlsxwriter用于解决该问题。

******** xlsxwriter 只能创建新文件或覆盖现有文件，无法读取或修改想现有文件，********
因此通常我们会结合openpyxl库进行使用

"""
# 使用xlsxwriter库对单元格内的部分内容进行标红
import xlsxwriter

#  创建以额新的excel 文件
workbook = xlsxwriter.Workbook('标红实例.xlsx')
#  生成一个表对象
sheet = workbook.add_worksheet()

# 创建一个红色字体格式
red_format = workbook.add_format({'color': 'red'})
text = "这是一段包含重要内容的文本"
# 需要标红的关键词
keyword = '重要'

position = text.find(keyword)
print(f'pos: {position}')

if position != 1:
    # 如果找到关键词，将文本分成三部分：前部分、关键词、后部分
    before_text = text[:position]  # 前部分内容
    after_text = text[position + len(keyword):]  # 后部分内容
    # 使用 write_rich_string 方法写入富文本
    # 参数依次为：行、列、普通文本、格式、需要标红的文本、普通文本
    sheet.write_rich_string(1, 1, before_text, red_format, keyword, after_text)  # 0行0列就表示A1单元格
else:
    sheet.write(0, 0, text)

workbook.close()
