from zhipuai import ZhipuAI
import os
from typing import Dict, List, Optional
import sys
import io
import json
import time

class ZhipuAPI:
    def __init__(self, max_retries: int = 3):
        self.api_key = "50ea3c031b07edc77a6a640ccb1526d1.NUhtei288b3OrwF4"
        if not self.api_key:
            raise ValueError("未设置ZHIPU_API_KEY环境变量")
        self.client = ZhipuAI(api_key=self.api_key)
        self.max_retries = max_retries

    def rewrite_text(self, text: str) -> str:
        """重写文本内容，确保返回JSON格式的结果"""
        if not text or not isinstance(text, str):
            return ""
        
        for attempt in range(self.max_retries):
            try:
                response = self.client.chat.completions.create(
                    model="GLM-4-AirX",
                    messages=[
                        {"role": "user", "content": f"""请重写以下文本，从语言风格、表达方式、逻辑结构等方面进行重写。请严格按照JSON格式返回，格式为{{"重写结果": "重写后的内容"}}：
请严格按照JSON格式返回，格式为{{"重写结果": "重写后的内容"}}：

原文：{text}"""}
                    ]
                )
                
                result = self._extract_json_result(response.choices[0].message.content)
                if result:
                    print("\nJSON提取成功!")
                    return result  # 返回提取后的结果
                else:
                    print(f"JSON提取失败，原始响应：{response.choices[0].message.content}")
                
                print(f"第{attempt + 1}次尝试未获得正确格式的响应，等待重试...")
                time.sleep(1)
                
            except Exception as e:
                print(f"API调用出错 (尝试 {attempt + 1}/{self.max_retries}): {str(e)}")
                if attempt < self.max_retries - 1:
                    time.sleep(1)
                continue
        
        return text  # 如果所有重试都失败，返回原文

    def _extract_json_result(self, text: str) -> Optional[str]:
        """从响应中提取包含重写结果的JSON内容"""
        try:
            # 首先尝试直接解析整个文本
            try:
                data = json.loads(text)
                if "重写结果" in data:
                    return data["重写结果"]
            except:
                pass

            # 查找文本中的JSON格式内容
            import re
            json_pattern = r'\{[^{}]*\}'
            matches = re.finditer(json_pattern, text)
            
            for match in matches:
                try:
                    data = json.loads(match.group())
                    if "重写结果" in data:
                        return data["重写结果"]
                except:
                    continue
            
            return None
        except:
            return None

    def test_extract_json(self):
        """测试JSON提取功能的各种场景"""
        api = ZhipuAPI()
        test_cases = [
            # 测试用例1：标准单行JSON
            {
                "input": '{"重写结果": "测试文本"}',
                "expected": "测试文本"
            },
            # 测试用例2：JSON嵌入在其他文本中
            {
                "input": "这是一段文本 {'重写结果': '测试文本'} 后面还有内容",
                "expected": "测试文本"
            },
            # 测试用例3：多行格式化JSON
            {
                "input": """
                {
                    "重写结果": "多行
                    测试文本"
                }
                """,
                "expected": "多行\n                测试文本"
            },
            # 测试用例4：包含多个JSON，应该返回第一个有效的
            {
                "input": '{"其他": "无关内容"} {"重写结果": "正确内容"} {"重写结果": "错误内容"}',
                "expected": "正确内容"
            },
            # 测试用例5：无效JSON
            {
                "input": "这不是一个JSON文本",
                "expected": None
            }
        ]

        for i, test_case in enumerate(test_cases, 1):
            result = self._extract_json_result(test_case["input"])
            expected = test_case["expected"]
            success = result == expected
            print(f"\n测试用例 {i}:")
            print(f"输入: {test_case['input']}")
            print(f"期望: {expected}")
            print(f"实际: {result}")
            print(f"结果: {'通过' if success else '失败'}")

__all__ = ['ZhipuAPI']

if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

    try:
        # 运行JSON提取测试
        print("=== 运行JSON提取测试 ===")
        api = ZhipuAPI()
        api.test_extract_json()
        
        # 运行原有的重写测试
        print("\n=== 运行文本重写测试 ===")
        test_text = "人工智能正在改变我们的生活方式。它让很多任务变得更简单，效率更高。"
        rewritten_text = api.rewrite_text(test_text)
        print(f"重写后的文本: {rewritten_text}")
    except Exception as e:
        print(f"测试过程中出现错误: {str(e)}")
