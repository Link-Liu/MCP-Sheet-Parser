#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Windows编码修复脚本
解决Windows控制台中文乱码问题
"""

import os
import sys
import subprocess

def fix_windows_encoding():
    """修复Windows控制台编码"""
    if os.name == 'nt':  # Windows系统
        try:
            # 设置控制台代码页为UTF-8
            os.system('chcp 65001 > nul')
            
            # 设置环境变量
            os.environ['PYTHONIOENCODING'] = 'utf-8'
            os.environ['PYTHONLEGACYWINDOWSSTDIO'] = 'utf-8'
            
            # 设置标准输出编码
            if hasattr(sys.stdout, 'reconfigure'):
                sys.stdout.reconfigure(encoding='utf-8')
            if hasattr(sys.stderr, 'reconfigure'):
                sys.stderr.reconfigure(encoding='utf-8')
            
            print("✅ Windows编码设置完成")
            return True
        except Exception as e:
            print(f"⚠️ 编码设置失败: {e}")
            return False
    else:
        print("✅ 非Windows系统，无需特殊编码设置")
        return True

def run_with_encoding_fix(script_path, *args):
    """使用编码修复运行脚本"""
    if not fix_windows_encoding():
        print("❌ 编码修复失败，继续运行...")
    
    # 构建命令
    cmd = [sys.executable, script_path] + list(args)
    
    try:
        # 设置环境变量
        env = os.environ.copy()
        if os.name == 'nt':
            env['PYTHONIOENCODING'] = 'utf-8'
            env['PYTHONLEGACYWINDOWSSTDIO'] = 'utf-8'
        
        # 运行脚本
        result = subprocess.run(
            cmd,
            env=env,
            encoding='utf-8',
            errors='replace'
        )
        
        return result.returncode
    except Exception as e:
        print(f"❌ 运行脚本失败: {e}")
        return 1

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python fix_encoding.py <script_path> [args...]")
        sys.exit(1)
    
    script_path = sys.argv[1]
    args = sys.argv[2:]
    
    exit_code = run_with_encoding_fix(script_path, *args)
    sys.exit(exit_code) 