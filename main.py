import openpyxl
import yaml

xlsx_name = "data.xlsx"

# 打开Excel文件
wb = openpyxl.load_workbook(xlsx_name)
sheet = wb.active

# 读取Excel表格数据并存储为Python对象
data = {'data': []}
for row in sheet.iter_rows(min_row=2, values_only=True):
    item = {'first': row[0], 'tag': row[1], 'type': row[2], 'number': row[3], 'sum': row[4], 'last': row[5]}
    data['data'].append(item)

for i in data['data']:
    print(i)

# 将Python对象转换成YAML字符串
yaml_data = yaml.dump(data)

# 将YAML字符串写入YAML文件
with open('data.yaml', 'w', encoding='utf-8') as file:
    file.write(yaml_data)

# 读取yaml文件内容
with open('data.yaml', 'r') as file:
    dataR = yaml.load(file, Loader=yaml.FullLoader)
print(dataR)

