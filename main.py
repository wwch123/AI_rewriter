import os
from dotenv import load_dotenv
from content_rewriter import ContentRewriter

def main():
    # 加载 .env 文件
    load_dotenv()
    
    # 检查环境变量
    if 'ZHIPU_API_KEY' not in os.environ:
        print("错误：未设置ZHIPU_API_KEY环境变量")
        return

    # 创建重写器
    rewriter = ContentRewriter()

    # 处理输入文件
    input_file = "F:\other_projects\content_rewriter_v0\input\吴王晨晖 2022210508.docx"
    try:
        rewriter.rewrite_content(input_file)
        print("处理完成！")
    except Exception as e:
        print(f"处理文件时出错: {str(e)}")

if __name__ == "__main__":
    main() 