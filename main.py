import openpyxl
import yaml

xlsx_name = "data.xlsx"

# 打开Excel文件
wb = openpyxl.load_workbook(xlsx_name)
sheet = wb.active

# 读取Excel表格数据并存储为Python对象
data = {'data': []}
for row in sheet.iter_rows(min_row=2, values_only=True):
    item = {'name': row[0], 'age': row[1], 'email': row[2]}
    data['data'].append(item)

for i in data['data']:
    print(i)

# 将Python对象转换成YAML字符串
yaml_data = yaml.dump(data)

# 将YAML字符串写入YAML文件
with open('data.yaml', 'w') as file:
    file.write(yaml_data)

# 读取yaml文件内容
with open('data.yaml', 'r') as file:
    dataR = yaml.load(file, Loader=yaml.FullLoader)
print(dataR)

'''
{'name': 's1', 'age': 22, 'email': 's1@xcel.com'}
{'name': 's2', 'age': 23, 'email': 's3@xcel.com'}
{'name': '宋', 'age': 24, 'email': 's4@xcel.com'}
{'data': [{'age': 22, 'email': 's1@xcel.com', 'name': 's1'}, {'age': 23, 'email': 's3@xcel.com', 'name': 's2'}, {'age': 24, 'email': 's4@xcel.com', 'name': '宋'}]}
'''