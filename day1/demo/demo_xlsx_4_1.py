
# 处理多个关键词的标红
import xlsxwriter
import re
# 定义一个功能模块，最后返回一个富文本列表，方便使用write_rich_string方法写入富文本
def gen_rich_text(content, keywords, red_format):
    """
    生成富文本列表，将指定关键词标红

    参数:
    content - 原始文本内容
    keywords - 需要标红的关键词列表
    red_format - 红色格式对象

    返回:
    富文本参数列表，可直接用于write_rich_string方法
    """
    if not  content:
        return [content]

    # 创建正则表达式模式，匹配所有关键词
    pattern = re.compile('|'.join(keywords))
    # 找出所有匹配的关键词及其位置
    matches = list(pattern.finditer(content))
    if not matches:
        return [matches]

    # 构建富文本参数列表
    rich_text_parts = []
    last_end = 0
    for match in matches:
        start, end = match.span()

        # 添加关键词前的普通文本
        if start > last_end:
            rich_text_parts.append(content[last_end:start])

        # 添加带红色格式的关键词
        rich_text_parts.append(red_format)
        rich_text_parts.append(content[start:end])

        last_end = end

    # 添加最后一个关键词后的普通文本
    if last_end < len(content):
        rich_text_parts.append(content[last_end:])

    return rich_text_parts


# 创建Excel文件
workbook = xlsxwriter.Workbook('多关键词标红示例.xlsx')
worksheet = workbook.add_worksheet()

# 创建红色字体格式
red_format = workbook.add_format({'color': 'red'})

# 示例文本,将文本中的“关键词”进行标红
text = "这是一个Excel示例文件，我们需要将特定关键词标红，比如Excel和关键词"
# 需要标红的关键词
keywords = ["Excel", "关键词", "标红"]

# 生成富文本参数列表
rich_text_parts = gen_rich_text(text, keywords, red_format)

# 写入富文本
worksheet.write_rich_string(0, 0, *rich_text_parts)  # 在0行0列（A1单元格）内插入指定格式的文本，*rich_text_parts表示以可变长度参数传入

# 保存并关闭工作簿
workbook.close()