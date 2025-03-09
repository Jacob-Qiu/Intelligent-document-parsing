"""
    FileProcessing:
    desensitization用于将输入的文字按照人名/公司/地点进行脱敏；
    seg2json用于将输入的文件夹内所有的.docx文档进行拆分，如果输入模型文件，则基于随机森林分类器、按照一级、二级、三级、正文内容进行切分，若不输入模型文件，则基于段落样式拆分，并将结果存储到指定格式的json文件中。
"""

import joblib
import json
import logging
import os
import torch
import warnings
import pandas as pd
import regex as re
import docx
from docx import Document
from ltp import LTP
from presidio_analyzer import AnalyzerEngine, RecognizerResult, EntityRecognizer, RecognizerRegistry
from presidio_anonymizer import AnonymizerEngine, OperatorConfig
from typing import Iterator
warnings.filterwarnings("ignore", category=FutureWarning)


# 自定义识别器
class LTPRecognizer(EntityRecognizer):
    def __init__(self, ltp):
        super().__init__(supported_entities=['Nh', 'Ni', 'Ns'])  # 命名实体类别：Nh（人名）；Ni（机构名）；Ns（地名）
        self.ltp = ltp
        # 日志记录器
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.ERROR)
        file_handler = logging.FileHandler('error.log')
        file_handler.setLevel(logging.ERROR)
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)

    def analyze(self, text, entities, nlp_artifacts=None):
        try:
            doc = self.ltp.pipeline(text, tasks=["cws", "ner"])
            words_num = [len(element) for element in doc.cws]
            '''
                cws：Chinese Word Segmentation（中文分词）
                pos：Part-of-Speech Tagging（词性标注）
                ner：Named Entity Recognition（命名实体识别）
                srl：语义角色标注
                dep：依存句法分析
                sdp：语义依存分析树
            '''
            results = []
            for ent in doc.ner:
                entity_type, entity_text, start_pos, end_pos = ent
                start = sum(words_num[:start_pos])
                end = sum(words_num[:end_pos]) + words_num[end_pos]
                if entity_type in entities:
                    result = RecognizerResult(
                        entity_type=entity_type,
                        start=start,
                        end=end,
                        score=0.85
                    )
                    results.append(result)
            return results

        except Exception as e:
            self.logger.error('本段文字无法正常进行命名实体类型识别：%s', text, exc_info=True)
            return None


# 智能切分
class AutoStyle:
    def __init__(self, m_path: str, cmd: tuple = None) -> None:
        self.m_path = m_path
        self.cmd = cmd

    def file_operation(self, in_path: str) -> list:
        """基于分类器的分类结果，对.docx文件各段落进行样式设定

        Args:
            in_path·(str): 单个.docx文件的路径

        Return:
            list：切分后的内容
        """
        try:
            contents = _auto_seg(in_path, self.m_path, self.cmd)
            return contents
        except Exception as e:
            raise Exception(f'无法转换样式: {in_path}, {e}')


# 文档处理
class FileProcessing:
    def __init__(self, model_path: str = '', input_path: str = '', output_path: str = '', cmd: tuple = None):
        self.model_path = model_path
        self.input_path = input_path
        self.output_path = output_path
        self.cmd = cmd
        self.text = str
        self.entities = []
        self.configs = []
        self.objects = []
        if self.input_path == '' and self.output_path == '' and self.model_path:
            # 匿名操作
            self.ltp = LTP(self.model_path)
            # 自定义识别器并放入注册表中
            registry = RecognizerRegistry()
            registry.load_predefined_recognizers()
            registry.add_recognizer(LTPRecognizer(self.ltp))
            self.analyzer = AnalyzerEngine(registry=registry)
            # 执行匿名器
            self.anonymizer = AnonymizerEngine()

        elif self.input_path and self.output_path and self.model_path:
            # 如果给定模型路径，则基于训练模型自动切分
            self.seg_method = 'auto'
            self.at = AutoStyle(self.model_path, self.cmd)

        elif self.input_path and self.output_path and self.model_path == '':
            # 如果未给定模型路径，则基于样式切分
            self.seg_method = 'style'

    # 文本脱敏
    def desensitization(self, text: str,
                        anomy_p: str = '', anomy_o: str = '', anomy_l: str = '',
                        specify: bool = False) -> str:
        '''将输入的文字按照人名/公司/地点进行脱敏

        Args:
            text·(str):需要脱敏的文本
            model_path·(str):预训练模型的路径
            anomy_p·(str):人名匿名，填写需要替换的文字，若为空，则不匿名该类型
            anomy_o·(str):公司匿名，填写需要替换的文字，若为空，则不匿名该类型
            anomy_l·(str):地点匿名，填写需要替换的文字，若为空，则不匿名该类型
            specify·(bool):若为True，则匿名命名实体（人名/公司/地点）中的特定对象

        Return:
            str:脱敏后的文本
        '''
        self.text = text
        if torch.cuda.is_available():
            self.ltp.to('cuda')
            print('使用GPU执行')
        else:
            print('使用CPU执行')

        # 确定命名实体类别及相应的替换内容
        #     Nh：人名（Person Name）
        #     Ni：机构名（Organization Name）
        #     Ns：地名（Location Name）
        if specify:
            # 人名、公司名、地点的特定对象
            self.objects = [re.compile(''), re.compile(r'华东(建筑)?(设计)?(研究)?(总)?院(有限公司)?'), re.compile('')]
        else:
            self.objects = [re.compile(''), re.compile(''), re.compile('')]

        if anomy_p:
            self.entities.append(['Nh'])
            self.configs.append({"Nh": OperatorConfig(operator_name="replace", params={"new_value": f"{anomy_p}"})})
        if anomy_o:
            self.entities.append(['Ni'])
            self.configs.append({"Ni": OperatorConfig(operator_name="replace", params={"new_value": f"{anomy_o}"})})
        if anomy_l:
            self.entities.append(['Ns'])
            self.configs.append({"Ns": OperatorConfig(operator_name="replace", params={"new_value": f"{anomy_l}"})})

        anonymized_text = text
        # 文字脱敏
        for index in range(len(self.entities)):
            entity = self.entities[index]
            objet = self.objects[index]
            config = self.configs[index]
            results = self.analyzer.analyze(text=text, entities=entity, language='en')
            results.sort(reverse=True)
            for result in results:
                entity_text = text[result.start:result.end]
                if objet is not None:
                    if objet.search(entity_text):
                        anonymized_text = self.anonymizer.anonymize(text=anonymized_text, analyzer_results=[result],
                                                                    operators=config).text
                else:
                    anonymized_text = self.anonymizer.anonymize(text=anonymized_text, analyzer_results=[result],
                                                                operators=config).text

        return anonymized_text

    # 文档切分（存入多份json）
    def seg2json(self) -> list:
        '''将输入的文件夹内所有的.docx文档按照一级、二级、三级、正文内容进行切分，并分别存储到单独的json文件中

        Args:
            None

        Return:
            list: 所有文档切分后的内容汇总，这个返回值是为后续数据入库。
        '''
        if os.path.isdir(self.input_path):
            file_paths, file_names = _get_file_path(self.input_path)
            succeed_file = 0
            file_datas = []
            for index, file_path in enumerate(file_paths):
                try:
                    if self.seg_method == 'auto':
                        # 由模型自动判别标题正文，并进行切分
                        file_content = self.at.file_operation(file_path)
                    elif self.seg_method == 'style':
                        # 基于文档的样式判别标题正文，并进行切分
                        file_content = _style_seg(file_path)
                    file_data = {
                        "title": file_names[index],
                        "content": "",
                        "children": file_content
                    }
                    file_datas.append(file_data)

                    # 保存每个文档为单独的JSON文件，这部分不是必须的，为方便检查各个文件的输出效果
                    json_filename = f'{file_names[index]}.json'
                    json_path = os.path.join(self.output_path, json_filename)
                    with open(json_path, 'w', encoding='utf-8') as f:
                        json.dump(file_data, f, ensure_ascii=False, indent=4)
                    print(f"已成功存储 {json_filename}")

                    succeed_file += 1
                except Exception as e:
                    print(f"无法读取原始文件: {file_path}，{e}")

            print(f"已成功读取的文件数：{succeed_file}")
            return file_datas

        else:
            raise FileNotFoundError(f"未找到指定路径的文件夹：{self.input_path}")


def _get_file_path(folder_path: str) -> tuple:
    # 获取文件夹中所有的文件相对路径及文件名
    doc_files = []
    filenames = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.docx'):
                filenames.append(file[:-5])
                doc_files.append(root + '/' + file)
            else:
                print(f"输入路径非.docx文件：{root + '/' + file}")
    return doc_files, filenames


def _auto_seg(in_path: str, m_path: str, cmd: tuple) -> list:
    # 基于分类器的分类结果进行切分
    clf = joblib.load(open(m_path, "rb"))
    _, lines = _extract_docx_to_paragraph(in_path)
    df = _feature_extraction(in_path, cmd)
    df = df.drop(['Text', 'Style'], axis=1)
    contents = []
    current_hierarchy = []
    for index, line in enumerate(lines):
        if clf.predict(df)[index] == 0:
            if not re.compile(r'\s\d+$').search(line.strip()):  # 防止把目录里文字的识别成标题，目录一般标题后会跟一个数字
                current_hierarchy = [{"title": line, "content": "", "children": []}]
                contents.append(current_hierarchy[0])
        elif clf.predict(df)[index] == 1:
            if not re.compile(r'\s\d+$').search(line.strip()):
                if current_hierarchy:
                    second_level = {"title": line, "content": "", "children": []}
                    current_hierarchy[0]["children"].append(second_level)
                    current_hierarchy = [current_hierarchy[0], second_level]
        elif clf.predict(df)[index] == 2:
            if not re.compile(r'\s\d+$').search(line.strip()):
                if len(current_hierarchy) > 1:
                    third_level = {"title": line, "content": "", "children": []}
                    current_hierarchy[1]["children"].append(third_level)
                    current_hierarchy.append(third_level)
        else:
            if current_hierarchy:
                current_hierarchy[-1]["content"] += line.strip()
    return contents


def _style_seg(in_path: str) -> list:
    # 基于文档的样式进行拆分
    lines, _ = _extract_docx_to_paragraph(in_path)
    contents = []
    current_hierarchy = []
    for line in lines:
        for key, value in line.items():
            if value == 0:
                current_hierarchy = [{"title": key, "content": "", "children": []}]
                contents.append(current_hierarchy[0])
            elif value == 1:
                if current_hierarchy:
                    second_level = {"title": key, "content": "", "children": []}
                    current_hierarchy[0]["children"].append(second_level)
                    current_hierarchy = [current_hierarchy[0], second_level]
            elif value == 2:
                if len(current_hierarchy) > 1:
                    third_level = {"title": key, "content": "", "children": []}
                    current_hierarchy[1]["children"].append(third_level)
                    current_hierarchy.append(third_level)
            else:
                if current_hierarchy:
                    current_hierarchy[-1]["content"] += key.strip()
    return contents


def _extract_docx_to_paragraph(file_path: str) -> (list, list):
    # 将word文档转换为markdown
    doc = Document(file_path)
    styled_paragraphs = []
    paragraphs = []
    for index, element in enumerate(_iter_paragraphs(doc)):
        # 判断是否包含图片或表格
        if isinstance(element, docx.text.paragraph.Paragraph):
            if element.style.name.startswith('Heading'):
                level = int(element.style.name[-1]) - 1
                styled_paragraphs.append({element.text: level})
            else:
                styled_paragraphs.append({element.text: 3})
            paragraphs.append(element.text)
        elif isinstance(element, docx.table.Table):
            styled_paragraphs.append({"": 3})
            paragraphs.append("")
    return styled_paragraphs, paragraphs


def _iter_paragraphs(parent: Document, recursive: bool = True) -> Iterator[object]:
    # 接收Document类型的数据，按文档顺序在*parent*中生成每个段落和表子项，每个返回值都是Paragraph的实例
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


def _feature_extraction(path: str, cmd: tuple) -> pd.DataFrame:
    # 对.docx文件进行特征提取，用于分类器分类
    if cmd is None:
        h1type, h2type, h3type = -1, -1, -1
    elif len(cmd) < 3:
        h1type, h2type, h3type = cmd + (-1,) * (3 - len(cmd))
    elif len(cmd) > 3:
        raise ValueError(
            "cmd必须是一个不超过三个元素的元组，分别指定一二三级标题的样式类型，如不属于任何预先设定的任何一类，则默认为-1")
    else:
        h1type, h2type, h3type = cmd

    regs = [r"^(?:[零一二三四五六七八九壹贰叁肆伍陆柒捌玖]+[、]?)\s*",
            r"^(?!.*\b\d+年(?:[0-9]+月[0-9]+日)?\b)(?:[0-9]+(?:[\.\、])?)\s*[a-zA-Z\u4e00-\u9fa5]+",
            r"^(?:[0-9]+(?:[\.\-][0-9]+)[\.\-]?)\s*[a-zA-Z\u4e00-\u9fa5]+",
            r"^(?:[(（][零一二三四五六七八九壹贰叁肆伍陆柒捌玖]+[)）])\s*",
            r"^(?:[0-9]+(?:[\.\-][0-9]+)(?:[\.\-][0-9]+)[\.\-]?)\s*[a-zA-Z\u4e00-\u9fa5]+",
            r"^(?:[(（]?[0-9]+[）)])\s*[a-zA-Z\u4e00-\u9fa5]+",
            r"^(第([零一二三四五六七八九壹贰叁肆伍陆柒捌玖]+)章[\.]?)\s*",
            r"^(第([零一二三四五六七八九壹贰叁肆伍陆柒捌玖]+|[0-9]+)节[\.]?)\s*",
            r"^(第([零一二三四五六七八九壹贰叁肆伍陆柒捌玖]+|[0-9]+)卷[\.]?)\s*",
            r"^(第([零一二三四五六七八九壹贰叁肆伍陆柒捌玖]+|[0-9]+)部分[\.]?)\s*"
            ]
    data = []
    document = Document(path)
    for index, element in enumerate(_iter_paragraphs(document)):
        # 判断是否包含图片或表格
        if isinstance(element, docx.text.paragraph.Paragraph):
            text = element.text
            contains_image = 'graphicData' in element._element.xml
            contains_table = False
        elif isinstance(element, docx.table.Table):
            text = '\n'.join(cell.text.strip() for row in element.rows for cell in row.cells)
            contains_image = False
            contains_table = True

        # 判断样式，分别对应一二三级标题和正文
        if element.style.name == "Heading 1":
            style = 0
        elif element.style.name == "Heading 2":
            style = 1
        elif element.style.name == "Heading 3":
            style = 2
        else:
            style = 3

        # 定义段落开头类型
        def start_type(x, patterns):
            for patterns_idx, pattern in enumerate(patterns):
                if re.compile(pattern).match(x):
                    return patterns_idx
            return -1

        # 判断样式，分别对应一二三级标题和正文
        if isinstance(element, docx.text.paragraph.Paragraph):
            if h1type == -1:
                h1type = start_type(element.text, regs)
            elif h2type == -1:
                if start_type(element.text, regs) != h1type:
                    h2type = start_type(element.text, regs)
            elif h3type == -1:
                if start_type(element.text, regs) != h1type and start_type(element.text, regs) != h2type:
                    h3type = start_type(element.text, regs)

        data.append({'Text': text,
                     'Style': style,
                     'Contains Image': int(contains_image),
                     'Contains Table': int(contains_table)})

    data = [{**dict_item, 'H1 Type': h1type, 'H2 Type': h2type, 'H3 Type': h3type} for dict_item in data]
    df = pd.DataFrame(data)
    df['Text'] = df['Text'].str.strip()
    df.reset_index(drop=True, inplace=True)
    df['Text'] = df['Text'].astype(str)
    df['words'] = df['Text'].apply(len)  # 统计字数
    df['endswithdot'] = df['Text'].str.endswith('。').astype(int)  # 是否以“。”结尾
    df['contains symbol'] = df['Text'].str.contains('，' or '？' or '！' or '；').astype(int)  # 是否包含，？！：；
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


# 示例用法use
if __name__ == "__main__":
    # 功能一：文本匿名，使用时需要下载预训练模型包，链接为：http://39.96.43.154/ltp/v4/base1.tgz，解压后将路径填写到anonymized_model_path中
    anonymized_model_path = 'path/to/model'
    td = FileProcessing(anonymized_model_path)
    text = 'input text'
    anonymized_text = td.desensitization(text, anomy_p='<人名>', anomy_o='<公司>', anomy_l='<地点>')
    print(f"输入内容为：  {text}\n")
    print(f"脱敏内容为：  {anonymized_text}")


    # 功能二：一个文件夹中所有文档进行切分；若模型路径不为空，则进行智能切分，否则基于docx文件的样式进行切分。
    """
        style_cmd是正则化形式组合
        正则化形式如下：
            0:   一  一、
            1：   1  1.  1、
            2：  1.1   1.1.   1-1
            3：  （一）  (一)
            4：  1.1.1   1.1.1.  1-1-1
            5：  （1）  (1)   1）   1)
            6：  第一章 第1章
            7：  第一节 第1节
            8：  第一卷 第1卷
            9：  第一部分 第1部分
            -1： 其他，由模型判定
        例如，若标题层级依次符合：第0、1、2类正则化形式，style_cmd请输入(0,1,2)；
        若输入(0,1)，则三级标题的样式类型由模型判断；
        若为None或(-1,-1,-1)，则全部由模型判断
    """
    input_path_s = 'input/path'
    output_path_s = 'output/path'
    model_path = 'model/path'  # 可不输入，若不输入，基于段落样式拆分
    style_cmd = None
    fs = FileProcessing(model_path, input_path_s, output_path_s, style_cmd)
    data = fs.seg2json()
