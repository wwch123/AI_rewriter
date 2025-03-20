import os
import argparse

def process_file(file_path: str, backup: bool = True) -> None:
    """处理单个文件
    
    Args:
        file_path: 要处理的文件路径
        backup: 是否创建备份文件
    """
    print(f"处理文件: {file_path}")
    
    # 获取图片目录的绝对路径
    md_dir = os.path.dirname(os.path.abspath(file_path))
    image_dir = os.path.join(md_dir, 'images').replace('\\', '/')
    
    # 读取文件内容
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as e:
        print(f"✗ 读取文件失败: {str(e)}")
        return
    
    # 替换图片路径
    new_content = content.replace('./images/', f'{image_dir}/')
    
    if new_content == content:
        print("✓ 未发现需要转换的图片路径")
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
        print(f"✓ 已完成转换，文件已更新")
    except Exception as e:
        print(f"✗ 保存文件失败: {str(e)}")

def main():
    parser = argparse.ArgumentParser(description='将Markdown文件中的相对图片路径转换为绝对路径')
    parser.add_argument('path', help='要处理的Markdown文件路径')
    parser.add_argument('--no-backup', action='store_true', help='不创建备份文件')
    
    args = parser.parse_args()
    path = os.path.abspath(args.path)
    
    if not os.path.exists(path):
        print(f"✗ 错误: 路径不存在: {path}")
        return
    
    if os.path.isfile(path):
        process_file(path, not args.no_backup)
    else:
        print(f"✗ 错误: 请指定一个文件而不是目录: {path}")

if __name__ == "__main__":
    main() 