import csv

# # 1、假设我们想要写入以下数据 ，数据必须构建成列表的形式才能写入
# rows = [
#     ["姓名", "年龄", "城市"],
#     ["张三", 28, "北京"],
#     ["李四", 34, "上海"],
#     ["王五", 29, "广州"]
# ]
#
# with open('test.csv', 'w', encoding='utf-8', newline='') as csvfile:
#     # 创建一个写入器
#     writer = csv.writer(csvfile)
#     # 遍历 rows 列表，csv会将每个列表中的数据一行行的写入文件
#     for row in rows:
#         writer.writerow(row)

# # 读取CSV文件
# with open('test.csv', 'r', encoding='utf-8') as file:
#     #  创建一个读取器
#     reader = csv.reader(file)
#     # 遍历 CSV 文件中的每一行数据，每行数据单独存放在一个列表中返回
#     for row in reader:
#         print(row)

"""
使用 DictReader 和 DictWriter

对于更复杂的场景，当 CSV 文件的列名很重要时，
可以使用 csv.DictReader 和 csv.DictWriter 类。
这些类允许你将 CSV 文件中的行作为字典来处理，其中列名作为键
"""

# # 写入数据
# # 文件的表头，也就是CSV文件中的第一行内容，列表中每个数据相当于下面字典数据中的键
# fieldnames = ['姓名', '年龄', '城市']
# # 将列表中的数据 以字典的形式 写入CSV文件
# rows = [
#     {'姓名': '张三', '年龄': 28, '城市': '北京'},
#     {'姓名': '李四', '年龄': 34, '城市': '上海'},
# ]
#
# with open('test1.csv','w',newline='',encoding='utf-8') as file:
#     # 创建一个写入器对象,fieldnames= 用于接收上面写的表头，不写的话，插入字典数据会报错
#     writer = csv.DictWriter(file, fieldnames=fieldnames)
#     # 先写入表头
#     writer.writeheader()
#     # 遍历rows, 将字典一行行的写入 CSV 文件
#     for row in rows:
#         writer.writerow(row)

# 读取数据
with open('test1.csv', 'r', encoding='utf-8') as file:
    # 创建一个读取器
    reader = csv.DictReader(file)
    # 遍历 CSV 文件中的每一行数据，以字典的形式返回
    """
        {'姓名': '张三', '年龄': '28', '城市': '北京'}
        {'姓名': '李四', '年龄': '34', '城市': '上海'}
    """
    for row in reader:
        print(row)
