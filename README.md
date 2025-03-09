# 项目简介
本项目是一个用于处理 Word 文档的自动化工具集，包含以下功能：
1. **自动编号转换**：将 Word 文档中的自动编号转换为普通文本。
2. **数据集生成**：从 Word 文档中提取特征，生成训练数据集。
3. **随机森林分类器训练**：基于生成的数据集训练随机森林分类器。
4. **文档处理**：支持文档脱敏、智能切分和基于样式的切分功能。

---

# 目录
- [需求配置](#需求配置)
- [环境](#环境)
- [代码模块说明](#代码模块说明)
  - [autonum_transfer](#autonum_transfer)
  - [preprocess](#preprocess)
  - [random_tree](#random_tree)
  - [FileProcessing](#fileprocessing)
- [使用方法](#使用方法)
- [示例](#示例)

---

# 需求配置
## 依赖库
运行本项目需要以下 Python 库：
```plaintext
pandas==2.0.3
numpy==1.26.0
python-docx==0.8.11
regex==2023.10.3
pywin32==306
joblib==1.3.2
torch==2.1.0+cu121
scikit-learn==1.3.2
ltp==4.2.13
presidio-analyzer==2.2.33
presidio-anonymizer==2.2.33
typing_extensions==4.8.0
```

## 环境
Python 版本：3.9 或更高版本
CUDA 版本：12.1（如需 GPU 支持）
操作系统：Windows / Linux（Linux无法执行“autonum_transfer.py”）

---

# 代码模块说明
## 1.autonum_transfer
### 功能
将 Word 文档中的自动编号转换为普通文本，避免后续处理时无法读取编号信息。

### 主要函数
docx2docm(input_path: str)
将 .docx 文件转换为 .docm 文件，以便添加宏代码。

transfer(input_path: str, output_path: str)
为 .docm 文件添加并运行宏代码，取消自动编号，最后将文件转换回 .docx 格式。

### 示例
```python
input_folder_path = "./input/folder/path"
output_folder_path = "./output/folder/path"
docx_files = [f for f in os.listdir(input_folder_path) if f.endswith('.docx')]
for file in docx_files:
    input_file_path = os.path.join(input_folder_path, file)
    output_file_path = os.path.join(output_folder_path, file)
    docx2docm(input_file_path)
    transfer(input_file_path, output_file_path)
```

## 2.preprocess
### 功能
从 Word 文档中提取特征，生成训练数据集。

### 主要函数
iter_paragraphs(parent: Document, recursive: bool = True)
遍历 Word 文档中的段落和表格，生成段落实例。

feature_extraction(path: str)
对文件夹内所有 .docx 文件进行特征提取，生成数据集。

### 示例
```python
folder_path = "./documents/folder/path"
dataset = feature_extraction(folder_path)
dataset.to_csv("./dataset/save/path", encoding='utf_8_sig')
```

## 3.random_tree
### 功能
基于生成的数据集训练随机森林分类器。

### 主要函数
train(feature_path: str, save_path: str, model_name: str)
训练随机森林分类器，并保存模型文件。

### 示例
```python
feature_path = './feature/path'
save_path = './model/save/path'
model_name = 'model_name.pkl'
train(feature_path, save_path, model_name)
```

## 4.FileProcessing
### 功能
文本脱敏：将输入的文字按照人名、公司、地点进行脱敏。
文档切分：将 Word 文档按一级、二级、三级标题和正文内容切分，并存储为 JSON 文件。

### 主要函数
desensitization(text: str, anomy_p: str, anomy_o: str, anomy_l: str, specify: bool)
对文本进行脱敏处理。

seg2json()
将文件夹内的所有 .docx 文档切分并存储为 JSON 文件。

### 示例
```python
# 文本脱敏
anonymized_model_path = 'path/to/model'
td = FileProcessing(anonymized_model_path)
text = 'input text'
anonymized_text = td.desensitization(text, anomy_p='<人名>', anomy_o='<公司>', anomy_l='<地点>')

# 文档切分
input_path_s = 'input/path'
output_path_s = 'output/path'
model_path = 'model/path'
style_cmd = None
fs = FileProcessing(model_path, input_path_s, output_path_s, style_cmd)
data = fs.seg2json()
```

---

# 许可证
本项目采用 MIT 许可证。

