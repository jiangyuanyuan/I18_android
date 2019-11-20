# coding=utf-8
# -*- coding: utf-8 -*-

import xlrd
import re


def transform(xls_path, xml_dir):
    data = xlrd.open_workbook(xls_path)
    table = data.sheets()[0]    # 获取第一张表
    if table.ncols <= 2:
        print "表格仅有两列内容, 无需导出xml文件"

    for tar_col in range(2, table.ncols):
        file_name = table.cell(0, tar_col).value
        file_path = xml_dir + '/' + file_name + '.xml'
        export_xml(table, tar_col, file_path)


def export_xml(table, tar_col, xml_path):
    from xml.etree.ElementTree import Element, SubElement, ElementTree
    root = Element('resources')

    row_count = table.nrows
    caches = {}

    for i in range(1, row_count-1):
        key = table.cell(i, 0).value
        value = table.cell(i, tar_col).value
        if key.startswith('<Array>'):
            pattern = "<Array>(.+?)<.+>"
            array_key = re.findall(pattern, key)[0]

            if array_key in caches.keys():
                array_node = caches[array_key]
                item_node = SubElement(array_node, 'item')
                item_node.text = value
            else:
                array_node = SubElement(root, 'string-array')
                array_node.set('name', array_key)
                item_node = SubElement(array_node, 'item')
                item_node.text = value
                caches[array_key] = array_node
        else:
            string_node = SubElement(root, 'string')
            string_node.text = value
            string_node.set('name', key)
    indent(root)
    tree = ElementTree(root)
    tree.write(xml_path, encoding='utf-8')


def indent(elem, level=0):
    i = "\n" + level*"\t"
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "\t"
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level+1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i


# 从Excel文件中导出安卓国际化文件
xls_path = "/Users/zed/Desktop/localized/Export/OEPay/Android/Android.xlsx"
xml_path = "/Users/zed/Desktop/localized/Export/OEPay/Android"
transform(xls_path, xml_path)
