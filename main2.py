import xlrd
import collections
import sys


def main(table_name, outfile_name):
    table_name = xlrd.open_workbook(table_name)
    sheet_name = table_name.sheet_by_index(0)
    title = sheet_name.row_values(0)

    for i in range(1, sheet_name.nrows):
        data = sheet_name.row_values(i)

        info_dict = collections.OrderedDict()
        for k in range(1, len(title)):

            # 如果空字符串或空置，跳过
            if data[k] == '' or data[k] is None:
                continue

            # 如果是字符串，转换为列表或字符串
            if isinstance(data[k], str):
                # 如果有逗号，识别为列表
                if ',' in data[k]:
                    dl = data[k].split(',')
                    dl = [x for x in dl if x != '' and x is not None]   # 去除列表中空字符串和空值
                    dl_type = int

                    # 检测列表中是否全是数字
                    for j in dl:
                        try:
                            int(j)
                        except:
                            dl_type = str
                            break

                    # 如果列表中全是数字，把列表转换为int类型
                    if dl_type == int:
                        dl = list(map(int, dl))

                    info_dict[title[k]] = dl
                else:
                    info_dict[title[k]] = data[k]

            # 如果为float类型，转换int
            if isinstance(data[k], float):
                info_dict[title[k]] = int(data[k])

        data_dict = collections.OrderedDict()
        data_dict[data[0]] = info_dict

        # 打开文件把字典以yaml格式写入
        with open(outfile_name, mode='a') as f:
            f.write(dict_to_yaml(data_dict))


def dict_to_yaml(dict_data):
    s = ""
    hh_code = "\n"  # 换行符
    sj_code = " " * 2  # 缩进符，两个空格
    for key, value in dict_data.items():
        s += key + ":" + hh_code
        print(s)
        for key2, value2 in value.items():
            s += sj_code + key2 + ": "
            if isinstance(value2, list):
                s += hh_code
                for value3 in value2:
                    s += sj_code * 2 + "- " + str(value3) + hh_code
            else:
                s += str(value2) + hh_code
    return s


if __name__ == '__main__':
    # 函数运行
    main(table_name='data3.xls', outfile_name='hosts.yaml')

    # 命令行参数运行  python3 test.py 主机名表格.xls yaml文件.txt
    # main(table_name=sys.argv[1], outfile_name=sys.argv[2])
