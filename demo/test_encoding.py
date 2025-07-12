#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
编码测试脚本
验证中文显示是否正常
"""

import os
import sys

def test_encoding():
    """测试编码"""
    print("🔍 编码测试开始...")
    print("=" * 40)
    
    # 测试基本信息
    print(f"操作系统: {os.name}")
    print(f"Python版本: {sys.version}")
    print(f"默认编码: {sys.getdefaultencoding()}")
    
    # 测试中文字符
    test_strings = [
        "✅ 中文显示测试",
        "📁 文件路径测试",
        "🎯 项目功能测试",
        "🚀 性能测试",
        "📊 数据统计"
    ]
    
    print("\n📝 中文字符测试:")
    for i, text in enumerate(test_strings, 1):
        print(f"{i}. {text}")
    
    # 测试特殊字符
    print("\n🎨 特殊字符测试:")
    print("进度条: |████████████████████████████████| 100%")
    print("状态: ✅ 成功 | ❌ 失败 | ⚠️ 警告")
    
    print("\n✅ 编码测试完成！")
    
    if os.name == 'nt':
        print("\n💡 Windows用户提示:")
        print("- 如果看到乱码，请使用 run_demo.bat 运行演示")
        print("- 或者确保控制台使用UTF-8编码")

if __name__ == "__main__":
    test_encoding() 