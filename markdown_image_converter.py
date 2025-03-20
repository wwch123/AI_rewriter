import re
import os
import argparse
from typing import Tuple

def convert_image_tags(content: str) -> Tuple[str, int]:
    """转换HTML格式的图片标签为Markdown格式
    
    Args:
        content: 包含HTML图片标签的文本内容
        
    Returns:
        Tuple[str, int]: 转换后的内容和替换次数
    """
    # 匹配HTML格式的图片标签，提取src和width属性
    pattern = r'<img.*?src="file:///([^"]*)".*?width="(\d{1,5})".*?>'
    
    def replace_func(match):
        # 获取图片路径和宽度
        path = match.group(1)
        width = match.group(2)
        
        # 提取文件名
        filename = os.path.basename(path)
        
        # 构造新的Markdown格式图片引用
        return f'![{filename}](./images/{filename}){{width={width}}}'
    
    # 使用正则表达式替换
    new_content, count = re.subn(pattern, replace_func, content, flags=re.DOTALL)
    return new_content, count

def process_file(file_path: str, backup: bool = True) -> None:
    """处理单个文件
    
    Args:
        file_path: 要处理的文件路径
        backup: 是否创建备份文件
    """
    print(f"处理文件: {file_path}")
    
    # 读取文件内容
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as e:
        print(f"✗ 读取文件失败: {str(e)}")
        return
    
    # 转换内容
    new_content, count = convert_image_tags(content)
    
    if count == 0:
        print("✓ 未发现需要转换的图片标签")
        return
    
    # 创建备份
    if backup:
        backup_path = file_path + '.bak'
        try:
            with open(backup_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"✓ 已创建备份文件: {backup_path}")
        except Exception as e:
            print(f"✗ 创建备份文件失败: {str(e)}")
            return
    
    # 保存修改后的内容
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f"✓ 已完成转换，共替换了 {count} 处图片标签")
    except Exception as e:
        print(f"✗ 保存文件失败: {str(e)}")

def process_directory(directory: str, backup: bool = True) -> None:
    """处理目录中的所有markdown文件
    
    Args:
        directory: 要处理的目录路径
        backup: 是否创建备份文件
    """
    print(f"处理目录: {directory}")
    
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith(('.md', '.markdown')):
                file_path = os.path.join(root, file)
                process_file(file_path, backup)

def main():
    parser = argparse.ArgumentParser(description='转换Markdown文件中的HTML图片标签为Markdown格式')
    parser.add_argument('path', help='要处理的文件或目录路径')
    parser.add_argument('--no-backup', action='store_true', help='不创建备份文件')
    
    args = parser.parse_args()
    path = os.path.abspath(args.path)
    
    if not os.path.exists(path):
        print(f"✗ 错误: 路径不存在: {path}")
        return
    
    if os.path.isfile(path):
        process_file(path, not args.no_backup)
    else:
        process_directory(path, not args.no_backup)

if __name__ == "__main__":
    main() 