"""
    文档中的自动编号会导致后续代码无法读取这些编号信息，本脚本通过调用宏代码将自动编号转为普通文本
"""


import os
import win32com.client as win32


def docx2docm(input_path: str) -> None:
    """将.docx文件转为.docm文件

    Args:
        input_path·(str): 需处理的.docx文件路径

    Return:
        None
    """
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False  # 不显示Word
    word.DisplayAlerts = False  # 不显示警告

    try:
        doc = word.Documents.Open(input_path)
        doc.SaveAs(input_path[:-5], FileFormat=13)  # 13 表示 .docm 格式
        doc.Close(SaveChanges=True)
    except Exception as ex:
        print(f'无法正常转换文件，{ex}')
    finally:
        word.Quit()


def transfer(input_path: str, output_path: str) -> None:
    """为.docm文件添加并运行宏文件，以取消自动编号；而后转换为.docx文件

    Args:
        input_path·(str): 输入.docm文件路径
        output_path·(str): 输出.docx文件路径

    Return:
        None
    """
    # 宏代码
    macro_code = '''
    Sub transfer()
    Dim kgslist As List
    For Each kgslist In ActiveDocument.Lists
    kgslist.ConvertNumbersToText
    Next
    End Sub
    '''
    word = win32.DispatchEx('Word.Application')
    word.Visible = False  # 不显示Word
    word.DisplayAlerts = False  # 不显示警告
    try:
        doc = word.Documents.Open(input_path[:-5] + '.docm')
        # 添加并运行宏
        vb_module = word.VBE.ActiveVBProject.VBComponents.Add(1)  # 1 表示标准模块
        vb_module.CodeModule.AddFromString(macro_code)
        macro_name = vb_module.Name + '.transfer'  # 宏名称为 transfer
        word.Application.Run(macro_name)

        # 删除模块以清理环境
        word.VBE.ActiveVBProject.VBComponents.Remove(vb_module)
        doc.SaveAs(output_path[:-5], FileFormat=12)  # 12 表示 .docx 格式
        doc.Close(SaveChanges=True)
    finally:
        word.Quit()


# 示例
if __name__ == "__main__":
    input_folder_path = "./input/folder/path"
    output_folder_path = "./output/folder/path"
    docx_files = [f for f in os.listdir(input_folder_path) if f.endswith('.docx')]
    for file in docx_files:
        try:
            input_file_path = os.path.join(input_folder_path, file)
            output_file_path = os.path.join(output_folder_path, file)
            docx2docm(input_file_path)
            transfer(input_file_path, output_file_path)
            os.remove(input_file_path[:-5] + '.docm')  # 移除中间步骤产生的.docm文件
            print(f'成功转换文档{file}')

        except Exception as e:
            print(f'无法转换文档：{file}，{e}')
