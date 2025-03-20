import os
import dashscope
from dashscope import Generation
from typing import Dict, List, Optional
import json
import time

class TongYiAPI:
    def __init__(self, max_retries: int = 5):
        self.api_key = os.environ.get("TONGYI_API_KEY", "")
        if not self.api_key:
            raise ValueError("未设置TONGYI_API_KEY环境变量")
        dashscope.api_key = self.api_key
        self.max_retries = max_retries

    def rewrite_text(self, text: str) -> str:
        """重写文本内容，确保返回JSON格式的结果"""
        if not text or not isinstance(text, str):
            return ""
        
        for attempt in range(self.max_retries):
            try:
                print(f"\n正在进行第 {attempt + 1} 次重写尝试...")
                response = Generation.call(
                    model='qwen-plus',
                    prompt=f"""请重写以下文本，从语言风格、表达方式、逻辑结构等方面进行重写，内容要改写，但是改写前后字数要基本一致。请严格按照JSON格式返回，格式为{{"重写结果": "重写后的内容"}}：原文：{text}""",
                    result_format='text',
                )
                
                if response.status_code == 200 and response.output:
                    print("\nAPI响应成功，正在提取结果...")
                    # 尝试解析JSON
                    result = self._extract_json_result(response.output.text)
                    if result:
                        print("\n提取JSON结果成功!")
                        return result  # 返回提取后的结果，而不是原始响应
                    else:
                        print('#'*100)
                        print(f"\nJSON提取失败，原始响应：\n{response.output.text}")
                        print('#'*100)

                print(f"\n第{attempt + 1}次尝试未获得正确格式的响应，等待重试...")
                time.sleep(1)  # 添加延迟避免请求过快
                
            except Exception as e:
                print(f"\nAPI调用出错 (尝试 {attempt + 1}/{self.max_retries}): {str(e)}")
                if attempt < self.max_retries - 1:
                    time.sleep(1)
                continue
        
        print("\n所有重试都失败，返回原文")
        return text  # 如果所有重试都失败，返回原文

    def _extract_json_result(self, text: str) -> Optional[str]:
        """从响应中提取包含重写结果的JSON内容"""
        try:
            # 1. 首先尝试直接解析整个文本
            try:
                # 处理可能的markdown代码块
                text = text.strip()
                if text.startswith('```') and text.endswith('```'):
                    # 移除markdown代码块标记
                    text = '\n'.join(text.split('\n')[1:-1])  # 保留换行符
                
                # 尝试直接解析
                try:
                    data = json.loads(text)
                    if "重写结果" in data:
                        return data["重写结果"]
                except:
                    # 如果直接解析失败，尝试查找JSON部分
                    start = text.find('{')
                    end = text.rfind('}')
                    if start != -1 and end != -1:
                        json_text = text[start:end+1]
                        try:
                            data = json.loads(json_text)
                            if "重写结果" in data:
                                return data["重写结果"]
                        except:
                            pass

            except Exception as e:
                print(f"第一阶段解析失败: {e}")
                pass

            # 2. 使用正则表达式查找JSON内容
            import re
            
            # 定义更精确的JSON模式
            json_patterns = [
                # 完整的JSON对象模式（双引号）
                r'\{\s*"重写结果"\s*:\s*"((?:[^"\\]|\\.|\\u[0-9a-fA-F]{4})*?)"\s*\}',
                
                # 完整的JSON对象模式（单引号）
                r"\{\s*'重写结果'\s*:\s*'((?:[^'\\]|\\.|\\u[0-9a-fA-F]{4})*?)'\s*\}",
                
                # 带有可能的转义字符的模式（双引号）
                r'\{\s*"重写结果"\s*:\s*"([^"]*?)"\s*\}',
                
                # 带有可能的转义字符的模式（单引号）
                r"\{\s*'重写结果'\s*:\s*'([^']*?)'\s*\}",
                
                # 最宽松的模式（双引号）
                r'\{[^{]*?"重写结果"\s*:\s*"(.*?)"\s*\}',
                
                # 最宽松的模式（单引号）
                r"\{[^{]*?'重写结果'\s*:\s*'(.*?)'\s*\}"
            ]
            
            for pattern in json_patterns:
                matches = re.finditer(pattern, text, re.DOTALL)
                for match in matches:
                    try:
                        # 提取捕获组中的内容
                        content = match.group(1)
                        if content:
                            return content
                    except:
                        continue
            
            # 3. 如果上述方法都失败，尝试最基本的提取
            try:
                start = text.find('"重写结果"') + len('"重写结果"')
                if start != -1:
                    start = text.find('"', start) + 1
                    if start != -1:
                        end = text.find('"', start)
                        if end != -1:
                            return text[start:end]
            except:
                pass
            
            return None
        except Exception as e:
            print(f"提取JSON时出错: {e}")
            return None


def main():
    try:

        
        print("\n=== 运行文本重写测试 ===")
        # 创建 TongYiAPI 实例
        api = TongYiAPI()
        
        # 测试文本
        test_text = "人工智能正在改变我们的生活方式。"
        print("原始文本:", test_text)
        
        # 调用重写方法
        rewritten_text = api.rewrite_text(test_text)
        print("重写后文本:", rewritten_text)
        
        # 测试空文本
        empty_text = ""
        print("\n测试空文本:")
        result = api.rewrite_text(empty_text)
        print("空文本结果:", result)
        
    except Exception as e:
        print(f"测试过程中出错: {str(e)}")

if __name__ == "__main__":
    main()

# 确保类被正确导出
__all__ = ['TongYiAPI'] 