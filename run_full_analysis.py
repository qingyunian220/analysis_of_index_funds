#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
指数基金分析完整流程脚本
========================
功能：筛选基金 → 生成网页截图和 PDF 报告

使用方法:
    python run_full_analysis.py
"""

import subprocess
import sys
import os


def check_dependencies():
    """检查必要的依赖是否已安装"""
    required_packages = ['pandas', 'akshare', 'openpyxl', 'flask', 'playwright']
    missing = []

    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing.append(package)

    if missing:
        print(f"警告：以下依赖包未安装：{missing}")
        print("请先运行：pip install {' '.join(missing)}")
        return False
    return True


def check_input_file():
    """检查输入文件是否存在"""
    if not os.path.exists('index-fund.xlsx'):
        print("错误：找不到 index-fund.xlsx 文件")
        print("请先运行 fund_app.py 获取基金数据")
        return False
    return True


def run_fund_selection():
    """运行基金筛选脚本"""
    print("\n" + "=" * 60)
    print("步骤 1/2: 运行基金筛选")
    print("=" * 60 + "\n")

    result = subprocess.run(
        [sys.executable, 'filtered_fund_selection.py'],
        encoding='utf-8',
        errors='ignore'
    )

    if result.returncode != 0:
        print(f"基金筛选失败，退出码：{result.returncode}")
        return False

    # 检查筛选结果文件是否生成
    if not os.path.exists('filtered_top_funds_selection.xlsx'):
        print("错误：筛选结果文件未生成")
        return False

    return True


def run_report_generation():
    """运行报告生成脚本"""
    print("\n" + "=" * 60)
    print("步骤 2/2: 生成报告截图和 PDF")
    print("=" * 60 + "\n")

    result = subprocess.run(
        [sys.executable, 'fund_analysis_website.py', '--mode', 'auto'],
        encoding='utf-8',
        errors='ignore'
    )

    if result.returncode != 0:
        print(f"报告生成失败，退出码：{result.returncode}")
        return False

    return True


def main():
    """主函数"""
    print("=" * 60)
    print("指数基金分析系统 - 完整流程")
    print("=" * 60)

    # 检查依赖
    if not check_dependencies():
        return

    # 检查输入文件
    if not check_input_file():
        return

    # 步骤 1: 运行基金筛选
    if not run_fund_selection():
        return

    # 步骤 2: 生成报告
    if not run_report_generation():
        return

    # 完成
    print("\n" + "=" * 60)
    print("分析流程完成!")
    print("=" * 60)
    print("\n生成的文件:")
    if os.path.exists('filtered_top_funds_selection.xlsx'):
        print("  - filtered_top_funds_selection.xlsx (筛选结果)")
    if os.path.exists('fund_analysis_screenshot.png'):
        print("  - fund_analysis_screenshot.png (网页截图)")
    if os.path.exists('fund_analysis_report.pdf'):
        print("  - fund_analysis_report.pdf (PDF 报告)")
    print("=" * 60)


if __name__ == '__main__':
    main()