#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MCP-Sheet-Parser 演示启动器
提供简单的演示入口，快速体验项目功能
"""

import os
import sys
import webbrowser
import subprocess
from pathlib import Path

def print_banner():
    """打印项目横幅"""
    banner = """
    ╔══════════════════════════════════════════════════════════════╗
    ║                                                              ║
    ║  🎯 MCP-Sheet-Parser 演示系统                                ║
    ║                                                              ║
    ║  专业的多格式表格解析与HTML转换工具                          ║
    ║                                                              ║
    ║  ✨ 支持11种格式 | 🎨 4种主题 | 🚀 高效转换                  ║
    ║                                                              ║
    ╚══════════════════════════════════════════════════════════════╝
    """
    print(banner)

def print_menu():
    """打印菜单选项"""
    menu = """
    📋 请选择演示选项:
    
    1️⃣  快速体验 - 基础功能演示
    2️⃣  完整演示 - 运行所有演示
    3️⃣  创建示例 - 生成示例文件
    4️⃣  查看文档 - 打开使用指南
    5️⃣  项目总结 - 查看实现状态
    6️⃣  打开主页 - 查看主展示页面
    0️⃣  退出演示
    
    """
    print(menu)

def quick_demo():
    """快速体验演示"""
    print("\n🚀 开始快速体验演示...")
    print("=" * 50)
    
    # 检查是否有示例文件
    sample_file = "示例文件/excel/basic_sample.xlsx"
    if not os.path.exists(sample_file):
        print("⚠️  示例文件不存在，正在创建...")
        create_samples()
    
    if os.path.exists(sample_file):
        print("✅ 找到示例文件，开始转换演示...")
        
        # 创建输出目录
        os.makedirs("demo/静态展示", exist_ok=True)
        
        # 运行基础转换
        output_file = "demo/静态展示/quick_demo.html"
        command = f"python main.py {sample_file} -o {output_file} --theme default"
        
        print(f"🔄 执行命令: {command}")
        
        try:
            # 设置环境变量以修复编码问题
            env = os.environ.copy()
            if os.name == 'nt':
                env['PYTHONIOENCODING'] = 'utf-8'
                env['PYTHONLEGACYWINDOWSSTDIO'] = 'utf-8'
            
            result = subprocess.run(
                command, 
                shell=True, 
                capture_output=True, 
                text=True, 
                encoding='utf-8',
                errors='replace',
                env=env
            )
            if result.returncode == 0:
                print("✅ 转换成功！")
                print(f"📁 输出文件: {output_file}")
                
                # 打开结果
                if os.path.exists(output_file):
                    webbrowser.open(f'file://{os.path.abspath(output_file)}')
                    print("🌐 已在浏览器中打开结果")
            else:
                print("❌ 转换失败")
                print(f"错误信息: {result.stderr}")
        except Exception as e:
            print(f"❌ 执行出错: {e}")
    else:
        print("❌ 无法创建示例文件")

def full_demo():
    """完整演示"""
    print("\n🎯 开始完整演示...")
    print("=" * 50)
    
    # 运行完整演示脚本
    demo_script = Path(__file__).parent / "演示脚本" / "run_demos.py"
    
    if demo_script.exists():
        print("🔄 运行完整演示脚本...")
        try:
            # 使用编码修复脚本运行
            fix_script = Path(__file__).parent / "fix_encoding.py"
            if fix_script.exists():
                subprocess.run([sys.executable, str(fix_script), str(demo_script)], check=True)
            else:
                # 直接运行，但设置环境变量
                env = os.environ.copy()
                if os.name == 'nt':
                    env['PYTHONIOENCODING'] = 'utf-8'
                    env['PYTHONLEGACYWINDOWSSTDIO'] = 'utf-8'
                subprocess.run([sys.executable, str(demo_script)], check=True, env=env)
        except subprocess.CalledProcessError as e:
            print(f"❌ 演示脚本执行失败: {e}")
        except Exception as e:
            print(f"❌ 执行出错: {e}")
    else:
        print("❌ 找不到演示脚本")

def create_samples():
    """创建示例文件"""
    print("\n📝 创建示例文件...")
    print("=" * 50)
    
    # 运行创建示例文件脚本
    script_path = Path(__file__).parent / "演示脚本" / "create_samples.py"
    
    if script_path.exists():
        try:
            subprocess.run([sys.executable, str(script_path)], check=True)
            print("✅ 示例文件创建完成")
        except subprocess.CalledProcessError as e:
            print(f"❌ 创建示例文件失败: {e}")
        except Exception as e:
            print(f"❌ 执行出错: {e}")
    else:
        print("❌ 找不到创建示例文件脚本")

def open_docs():
    """打开文档"""
    print("\n📚 打开文档...")
    print("=" * 50)
    
    docs = [
        ("使用指南", "demo/文档/使用指南.md"),
        ("项目实现总结", "demo/项目实现总结.md"),
        ("展示规划", "demo/展示规划.md")
    ]
    
    for name, path in docs:
        if os.path.exists(path):
            try:
                # 尝试用默认程序打开
                os.startfile(path) if os.name == 'nt' else subprocess.run(['xdg-open', path])
                print(f"✅ 已打开: {name}")
            except Exception as e:
                print(f"⚠️  无法打开 {name}: {e}")
        else:
            print(f"❌ 文件不存在: {path}")

def show_summary():
    """显示项目总结"""
    print("\n📊 项目实现总结")
    print("=" * 50)
    
    summary_file = "demo/项目实现总结.md"
    if os.path.exists(summary_file):
        try:
            with open(summary_file, 'r', encoding='utf-8') as f:
                content = f.read()
                print(content)
        except Exception as e:
            print(f"❌ 读取总结文件失败: {e}")
    else:
        print("❌ 总结文件不存在")

def open_homepage():
    """打开主展示页面"""
    print("\n🏠 打开主展示页面...")
    print("=" * 50)
    
    homepage = "demo/静态展示/index.html"
    if os.path.exists(homepage):
        try:
            webbrowser.open(f'file://{os.path.abspath(homepage)}')
            print("✅ 已在浏览器中打开主展示页面")
        except Exception as e:
            print(f"❌ 无法打开页面: {e}")
    else:
        print("❌ 主展示页面不存在")

def check_environment():
    """检查环境"""
    print("🔍 检查运行环境...")
    
    # 修复Windows编码问题
    if os.name == 'nt':
        try:
            os.system('chcp 65001 > nul')
            os.environ['PYTHONIOENCODING'] = 'utf-8'
            os.environ['PYTHONLEGACYWINDOWSSTDIO'] = 'utf-8'
            print("✅ Windows编码设置完成")
        except Exception as e:
            print(f"⚠️ 编码设置失败: {e}")
    
    # 检查Python版本
    python_version = sys.version_info
    print(f"✅ Python版本: {python_version.major}.{python_version.minor}.{python_version.micro}")
    
    # 检查项目文件
    if not os.path.exists("main.py"):
        print("❌ 请在项目根目录运行此脚本")
        return False
    
    # 检查必要目录
    os.makedirs("demo/静态展示", exist_ok=True)
    os.makedirs("demo/动态演示", exist_ok=True)
    os.makedirs("demo/文档", exist_ok=True)
    
    print("✅ 环境检查完成")
    return True

def main():
    """主函数"""
    print_banner()
    
    if not check_environment():
        sys.exit(1)
    
    while True:
        print_menu()
        
        try:
            choice = input("请输入选项 (0-6): ").strip()
            
            if choice == '0':
                print("\n👋 感谢使用MCP-Sheet-Parser演示系统！")
                break
            elif choice == '1':
                quick_demo()
            elif choice == '2':
                full_demo()
            elif choice == '3':
                create_samples()
            elif choice == '4':
                open_docs()
            elif choice == '5':
                show_summary()
            elif choice == '6':
                open_homepage()
            else:
                print("❌ 无效选项，请重新选择")
            
            input("\n按回车键继续...")
            
        except KeyboardInterrupt:
            print("\n\n👋 用户中断，退出演示系统")
            break
        except Exception as e:
            print(f"\n❌ 发生错误: {e}")
            input("按回车键继续...")

if __name__ == "__main__":
    main() 