from document_extractor import DocumentExtractor
import os

def test_document_extraction():
    # 初始化提取器
    extractor = DocumentExtractor()
    
    # 测试文件路径
    test_file = "F:\other_projects\content_rewriter_v0\input\TEST_1.docx"  # 请将你的测试文档放在这里
    
    # 确保文件存在
    if not os.path.exists(test_file):
        print(f"错误：找不到测试文件 {test_file}")
        return

    print("\n" + "="*50)
    print(f"开始测试文件：{test_file}")
    print("="*50)

    try:
        # 提取内容
        result = extractor.extract_content(test_file)
        
        # 打印文档结构
        print("\n文档结构:")
        print("-"*30)
        for heading in result['structure']:
            indent = "  " * (heading['level'] - 1)
            print(f"{indent}{'#'*heading['level']} {heading['text']}")

        # 打印内容块信息
        print("\n内容块详情:")
        print("-"*30)
        for idx, block in enumerate(result['content_blocks'], 1):
            print(f"\n[Block {idx}]")
            print(f"类型: {block['type']}")
            
            if block['type'] == 'text':
                print(f"内容: {block['content'][:100]}..." if len(block['content']) > 100 else f"内容: {block['content']}")
                print(f"格式信息:")
                print(f"  - 样式名称: {block['format_info']['style_name']}")
                print(f"  - 对齐方式: {block['format_info']['alignment']}")
                print(f"  - 行间距: {block['format_info']['line_spacing']}")
                
            elif block['type'] == 'heading':
                print(f"内容: {block['content']}")
                print(f"级别: {block['level']}")
                print(f"格式信息: {block['format_info']}")
                
            elif block['type'] == 'image':
                print(f"图片路径: {block['content']}")
                print(f"格式信息:")
                print(f"  - 是否内嵌: {block['format_info']['is_inline']}")
                print(f"  - 宽度: {block['format_info']['width']}")
                print(f"  - 高度: {block['format_info']['height']}")

        # 打印统计信息
        print("\n统计信息:")
        print("-"*30)
        content_types = {}
        for block in result['content_blocks']:
            content_types[block['type']] = content_types.get(block['type'], 0) + 1
        
        for content_type, count in content_types.items():
            print(f"{content_type}: {count} 个")

    except Exception as e:
        print(f"\n测试过程中出现错误：")
        print(f"错误类型: {type(e).__name__}")
        print(f"错误信息: {str(e)}")
        raise

if __name__ == "__main__":
    test_document_extraction() 