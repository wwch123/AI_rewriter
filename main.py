import os
from dotenv import load_dotenv
from content_rewriter import ContentRewriter

def main():
    # 加载 .env 文件
    load_dotenv()
    
    # 检查环境变量
    if 'ZHIPU_API_KEY' not in os.environ and 'TONGYI_API_KEY' not in os.environ:
        print("错误：未设置API密钥环境变量，请在.env文件中设置ZHIPU_API_KEY或TONGYI_API_KEY")
        return

    # 创建重写器
    rewriter = ContentRewriter()

    # 处理输入文件
    input_file = os.path.join("input", "example.docx")  # 使用相对路径，指向input目录下的示例文件
    
    # 检查输入文件是否存在
    if not os.path.exists(input_file):
        print(f"错误：输入文件不存在 ({input_file})，请将要处理的文件放在input目录")
        return
        
    try:
        rewriter.rewrite_content(input_file)
        print("处理完成！")
    except Exception as e:
        print(f"处理文件时出错: {str(e)}")

if __name__ == "__main__":
    main() 