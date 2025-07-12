'''
Autor: yangfan
Date: 2025-07-12 19:12:45
'''
from docx import Document
from docxcompose.composer import Composer



def get_docx_files(folder_path):
    """
    获取指定文件夹下所有的docx文件路径
    
    Args:
        folder_path: 文件夹路径
        
    Returns:
        docx_files: 包含所有docx文件完整路径的列表
    """
    import os
    
    docx_files = []
    
    # 遍历文件夹
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            # 检查文件扩展名是否为.docx
            if file.lower().endswith('.docx'):
                # 获取完整文件路径并添加到列表
                full_path = os.path.join(root, file)
                docx_files.append(full_path)
                
    return docx_files


def main(files,combined_file):
    """
    合并多个docx文件为一个文件
    
    Args:
        files: 包含所有docx文件完整路径的列表
        combined_file: 合并后的文件路径
    """
    new_document = Document()
    composer = Composer(new_document)
    for file in files:
        composer.append(Document(file))
    composer.save(combined_file)

if __name__ == "__main__":
    files = get_docx_files(r"E:\e2w\大同-天津南-20250712\最终")  
    main(files,r"E:\e2w\大同-天津南-20250712\result.docx")