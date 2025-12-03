import pandas as pd
import numpy as np
from flask import Flask, render_template, jsonify
import os
import webbrowser
import threading
import time
import asyncio
from playwright.async_api import async_playwright
import argparse


app = Flask(__name__, template_folder='templates', static_folder='static')


def load_fund_data():
    """
    加载基金数据用于网页显示
    """
    try:
        # 读取筛选后的基金数据
        file_path = "filtered_top_funds_selection.xlsx"
        if not os.path.exists(file_path):
            print(f"文件 {file_path} 不存在，请先运行筛选程序")
            return []
        
        df = pd.read_excel(file_path)
        funds_data = []
        
        for _, row in df.iterrows():
            # 格式化基金代码为6位
            fund_code = str(row['基金代码']) if not pd.isna(row['基金代码']) else ''
            # 确保基金代码是6位数字，不足6位前面补0
            fund_code = fund_code.zfill(6)
            
            fund_info = {
                'index_type': row['指数类型'] if not pd.isna(row['指数类型']) else '',
                'fund_name': row['基金简称'] if not pd.isna(row['基金简称']) else '',
                'fund_code': fund_code,
                'scale': str(row['最新规模']) if not pd.isna(row['最新规模']) else 'N/A',
                'turnover': str(row['换手率']) if not pd.isna(row['换手率']) else 'N/A',
                'concentration': str(row['前10大重仓股占比']) if not pd.isna(row['前10大重仓股占比']) else 'N/A',
                'return_1m': round(float(row['近1月超额']), 2) if not pd.isna(row['近1月超额']) else 0,
                'return_3m': round(float(row['近3月超额']), 2) if not pd.isna(row['近3月超额']) else 0,
                'return_6m': round(float(row['近6月超额']), 2) if not pd.isna(row['近6月超额']) else 0,
                'return_1y': round(float(row['近1年超额']), 2) if not pd.isna(row['近1年超额']) else 0,
                'score': float(row['综合得分']) if not pd.isna(row['综合得分']) else 0
            }
            funds_data.append(fund_info)
        
        # 按指数类型排序，再按综合得分排序
        funds_data.sort(key=lambda x: (x['index_type'], -x['score']))
        return funds_data
        
    except Exception as e:
        print(f"加载基金数据时出错: {e}")
        return []


@app.route('/')
def index():
    """
    主页面路由
    """
    return render_template('fund_analysis.html')


@app.route('/api/funds')
def get_funds():
    """
    获取基金数据的API接口
    """
    funds_data = load_fund_data()
    return jsonify(funds_data)


def open_browser():
    """
    延迟打开浏览器以确保服务器已启动
    """
    time.sleep(3)
    webbrowser.open('http://localhost:5000')


def run_website():
    """
    启动Web服务器
    """
    print("=" * 50)
    print("指数基金筛选分析结果展示网站")
    print("=" * 50)
    print("1. 网站将在本地运行，地址: http://localhost:5000")
    print("2. 请确保端口5000未被占用")
    print("3. 网站启动后将自动打开浏览器")
    print("4. 按 Ctrl+C 停止服务器")
    print("=" * 50)
    
    # 在单独的线程中打开浏览器
    threading.Timer(1.0, open_browser).start()
    
    # 启动Flask应用
    app.run(host='localhost', port=5000, debug=False, use_reloader=False)


async def capture_website_screenshot():
    """
    使用Playwright截取网页屏幕截图并保存为PNG和PDF格式
    """
    async with async_playwright() as p:
        # 启动浏览器
        browser = await p.chromium.launch()
        page = await browser.new_page()
        
        # 访问网站
        print("正在访问网站...")
        try:
            await page.goto('http://localhost:5000', wait_until='networkidle')
        except Exception as e:
            print(f"无法访问网站: {e}")
            print("请确保网站已在 http://localhost:5000 运行")
            await browser.close()
            return
        
        # 等待页面加载完成
        await page.wait_for_timeout(5000)
        
        # 截取整个页面的PNG截图
        print("正在截取PNG截图...")
        await page.screenshot(path='fund_analysis_screenshot.png', full_page=True)
        print("PNG截图已保存为 fund_analysis_screenshot.png")
        
        # 生成PDF报告（横向页面，更大尺寸）
        print("正在生成PDF报告...")
        await page.pdf(
            path='fund_analysis_report.pdf', 
            format='A4', 
            print_background=True,
            landscape=True,  # 横向页面
            margin={"top": "1cm", "bottom": "1cm", "left": "1cm", "right": "1cm"}
        )
        print("PDF报告已保存为 fund_analysis_report.pdf")
        
        await browser.close()


def take_screenshots():
    """
    截取网站屏幕截图
    """
    print("正在生成基金分析报告...")
    
    # 尝试截取网站截图（需要网站正在运行）
    try:
        asyncio.run(capture_website_screenshot())
    except Exception as e:
        print(f"截图过程中出现错误: {e}")
        print("请确保网站正在运行后再尝试截图")


def main():
    parser = argparse.ArgumentParser(description='指数基金筛选分析网站')
    parser.add_argument('--mode', choices=['run', 'screenshot'], 
                        default='run', help='运行模式: run (启动网站) 或 screenshot (截取网站截图)')
    
    args = parser.parse_args()
    
    if args.mode == 'run':
        run_website()
    elif args.mode == 'screenshot':
        take_screenshots()


if __name__ == '__main__':
    main()