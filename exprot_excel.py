# coding=utf-8
# -*- coding: utf-8 -*-

import re
import sys
import xlwt

reload(sys)
sys.setdefaultencoding("utf-8")


# 扫描iOS的strings国际化文件，返回词汇字典
def find_strings(path):
    source = open(path).read()

    # 正则匹配
    content = source.decode("utf8")
    # keyword = "(\"^(?!\")[\w\W].+$\".localized\(\))"
    keyword = "\"[^\"]+\".localized"
    pattern = re.compile(keyword)
    results = pattern.findall(content)
    print results

    words = []
    for result in results:
        words.append(result)

    return words


# 扫描Android的XML国际化文件，返回词汇字典
def find_xml(path):
    import xml.etree.ElementTree as ET
    tree = ET.parse(path)
    node_dict = {}
    # root = tree.getroot()
    # print root.tag
    # print root.attrib

    string_nodes = tree.getiterator('string')
    for string_node in string_nodes:
        attr = string_node.attrib
        key = attr.get('name')
        value = string_node.text
        node_dict[key] = value

    array_nodes = tree.getiterator('string-array')
    for array_node in array_nodes:
        attr = array_node.attrib
        item_nodes = array_node.getiterator('item')
        index = 0
        for item_node in item_nodes:
            key = '<Array>%s<%d>' % (attr.get('name'), index)
            index += 1
            node_dict[key] = item_node.text

    return node_dict


# 将Android国际化内容写入到`excel`文件
def save_xml_to_excel(path, word_list):
    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet('Localized by Zed', cell_overwrite_ok=True)  # 创建sheet

    sheet.write(0, 0, label='Key')
    sheet.write(0, 1, label='Chinese')
    # 将数据写入到第row行第col列
    row = 1
    key_list = sorted(word_list.keys())
    for key in key_list:
        sheet.write(row, 0, label=key)
        sheet.write(row, 1, label=word_list[key])
        row += 1
    book.save(path)
    print("保存文件成功")


# 提取安卓的国际化文本
local_words = find_xml("/Users/zed/Desktop/localized/Export/OEPay/Android/strings-cn.xml")
save_xml_to_excel("/Users/zed/Desktop/localized/Export/OEPay/Android/strings-cn.xls", local_words)
