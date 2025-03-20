#!/usr/bin/env python3
"""
内容重写工具启动脚本
"""
import sys
import os
import traceback

def check_dependencies():
    """检查依赖是否安装"""
    try:
        import PyQt5
        import docx
        import fitz
        from dotenv import load_dotenv
        return True
    except ImportError as e:
        print(f"错误: 缺少必要的依赖项 - {str(e)}")
        print("请运行: pip install -r requirements.txt")
        return False

def main():
    """主函数"""
    print("正在启动内容重写工具...")
    
    # 检查依赖项
    if not check_dependencies():
        input("按任意键退出...")
        return
    
    try:
        # 导入GUI模块
        from gui import ContentRewriterGUI, QApplication
        
        # 启动应用
        app = QApplication(sys.argv)
        app.setStyle('Fusion')
        
        window = ContentRewriterGUI()
        window.show()
        
        sys.exit(app.exec_())
    except Exception as e:
        print(f"启动失败: {str(e)}")
        print("\n详细错误信息:")
        traceback.print_exc()
        input("\n按任意键退出...")

if __name__ == "__main__":
    main() 