"""
    生成训练用数据集
"""


import docx.document
import docx.oxml.table
import docx.oxml.text.paragraph
import docx.table
import docx.text.paragraph
import os
import regex as re
import pandas as pd
from docx import Document
from typing import Iterator


def iter_paragraphs(parent: Document, recursive: bool = True) -> Iterator[object]:
    """接收Document类型的数据，按文档顺序在*parent*中生成每个段落和表子项，每个返回值都是Paragraph的实例

    Args:
        parent·(Document): Document实例化的对象
        recursive·(bool): 决定是否递归处理表格中的段落。当 recursive 为 True 时，函数才会递归处理表格中的段落

    Return:
        Iterator[object]: Paragraph类型的实例
    """
    if isinstance(parent, docx.document.Document):
        parent_elm = parent.element.body
    elif isinstance(parent, docx.table._Cell):
        parent_elm = parent._tc
    else:
        raise TypeError(repr(type(parent)))

    for child in parent_elm.iterchildren():
        if isinstance(child, docx.oxml.text.paragraph.CT_P):
            yield docx.text.paragraph.Paragraph(child, parent)
        elif isinstance(child, docx.oxml.table.CT_Tbl):
            if recursive:
                table = docx.table.Table(child, parent)
                yield table


def feature_extraction(path: str) -> pd.DataFrame:
    """对文件夹内所有的.docx文件进行特征提取

    Args:
        path·(str): 文件夹路径

    Return:
        pd.DataFrame: 特征数据集
    """
    '''
        10类正则化形式如下：
        0:   一  一、
        1：   1  1.  1、
        2：  1.1   1-1  1.1、
        3：  （一）  (一)
        4：  1.1.1  1-1-1  1.1.1、
        5：  （1）  (1)   1）   1)
        6：  第一章
        7：  第一节
        8：  第一卷
        9：  第一部分
    '''
    regs = [r"^(?:[零一二三四五六七八九壹贰叁肆伍陆柒捌玖]+[\、]?)\s*",
            r"^(?:[0-9]+(?:[\.\、])?)\s*[a-zA-Z\u4e00-\u9fa5]+",
            r"^(?:[0-9]+(?:[\.\-][0-9]+)(?:[\.\、])?)\s*[a-zA-Z\u4e00-\u9fa5]+",
            r"^(?:[\(\（][零一二三四五六七八九壹贰叁肆伍陆柒捌玖]+[\)\）])\s*",
            r"^(?:[0-9]+(?:[\.\-][0-9]+)(?:[\.\-][0-9]+)(?:[\.\、])?)\s*[a-zA-Z\u4e00-\u9fa5]+",
            r"^(?:[\(\（]?[0-9]+[\）\)](?:[\.\、])?)\s*[a-zA-Z\u4e00-\u9fa5]+",
            r"^(第[零一二三四五六七八九壹贰叁肆伍陆柒捌玖]+章)\s*",
            r"^(第[零一二三四五六七八九壹贰叁肆伍陆柒捌玖]+节)\s*",
            r"^(第[零一二三四五六七八九壹贰叁肆伍陆柒捌玖]+卷)\s*",
            r"^(第[零一二三四五六七八九壹贰叁肆伍陆柒捌玖]+部分)\s*"
            ]

    docx_files = [f for f in os.listdir(path) if f.endswith('.docx')]
    datas = []
    for file_name in docx_files:
        h1type, h2type, h3type = -1, -1, -1
        data = []
        file_path = os.path.join(path, file_name)
        document = Document(file_path)
        for element in iter_paragraphs(document):
            # 判断是否包含图片或表格
            if isinstance(element, docx.text.paragraph.Paragraph):
                text = element.text
                contains_image = 'graphicData' in element._element.xml
                contains_table = False
            elif isinstance(element, docx.table.Table):
                contains_image = False
                contains_table = True

            # 定义段落开头类型
            def start_type(x, patterns):
                for patterns_idx, pattern in enumerate(patterns):
                    if re.compile(pattern).match(x):
                        return patterns_idx
                return -1

            # 判断样式，分别对应一二三级标题和正文
            if element.style.name == "Heading 1":
                style = 0
                if h1type == -1:
                    h1type = start_type(element.text, regs)
            elif element.style.name == "Heading 2":
                style = 1
                if h2type == -1:
                    h2type = start_type(element.text, regs)
            elif element.style.name == "Heading 3":
                style = 2
                if h3type == -1:
                    h3type = start_type(element.text, regs)
            else:
                style = 3

            data.append({'Text': text,
                         'Style': style,
                         'Contains Image': int(contains_image),
                         'Contains Table': int(contains_table)})

        data = [{**dict_item, 'H1 Type': h1type, 'H2 Type': h2type, 'H3 Type': h3type} for dict_item in data]
        datas.extend(data)

    df = pd.DataFrame(datas)
    df.replace('', pd.NA, inplace=True)
    df = df.dropna(subset=['Text'])  # 删除空字符
    df['Text'] = df['Text'].str.strip()
    df.reset_index(drop=True, inplace=True)
    df['Text'] = df['Text'].astype(str)
    df['words'] = df['Text'].apply(len)  # 统计字数
    df['endswithdot'] = df['Text'].str.endswith('。').astype(int)  # 是否以“。”结尾
    df['contains symbol'] = df['Text'].str.contains('，' or '？' or '！' or '；').astype(int)  # 是否包含，？！；
    df['contains percentage'] = df['Text'].str.contains('%').astype(int)  # 是否包含“%”

    def start_type2(x, patterns):
        if re.compile(patterns).match(x):
            return 1
        else:
            return 0

    # 段落开头样式
    for index, reg in enumerate(regs):
        df[f'Start Type{index}'] = df['Text'].apply(lambda x: start_type2(x, reg))
    return df


# 示例
if __name__ == '__main__':
    folder_path = "./documents/folder/path"
    dataset = feature_extraction(folder_path)
    dataset.to_csv("./dataset/save/path", encoding='utf_8_sig')
    print('成功生成数据集')
