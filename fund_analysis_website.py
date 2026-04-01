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

        # 获取列名映射（处理可能的编码问题）
        cols = df.columns.tolist()
        # 根据列的位置获取数据，避免列名编码问题
        # 列顺序：指数类型，基金代码，基金简称，最新规模，换手率，前 10 大重仓股占比，近 1 月超额，近 3 月超额，近 6 月超额，近 1 年超额，综合得分
        col_map = {
            'index_type': cols[0],
            'fund_code': cols[1],
            'fund_name': cols[2],
            'scale': cols[3],
            'turnover': cols[4],
            'concentration': cols[5],
            'return_1m': cols[6],
            'return_3m': cols[7],
            'return_6m': cols[8],
            'return_1y': cols[9],
            'score': cols[10]
        }

        for _, row in df.iterrows():
            # 格式化基金代码为 6 位
            fund_code = str(row[col_map['fund_code']]) if not pd.isna(row[col_map['fund_code']]) else ''
            # 确保基金代码是 6 位数字，不足 6 位前面补 0
            fund_code = fund_code.zfill(6)

            fund_info = {
                'index_type': str(row[col_map['index_type']]) if not pd.isna(row[col_map['index_type']]) else '',
                'fund_name': str(row[col_map['fund_name']]) if not pd.isna(row[col_map['fund_name']]) else '',
                'fund_code': fund_code,
                'scale': str(row[col_map['scale']]) if not pd.isna(row[col_map['scale']]) else 'N/A',
                'turnover': str(row[col_map['turnover']]) if not pd.isna(row[col_map['turnover']]) else 'N/A',
                'concentration': str(row[col_map['concentration']]) if not pd.isna(row[col_map['concentration']]) else 'N/A',
                'return_1m': round(float(row[col_map['return_1m']]), 2) if not pd.isna(row[col_map['return_1m']]) else 0,
                'return_3m': round(float(row[col_map['return_3m']]), 2) if not pd.isna(row[col_map['return_3m']]) else 0,
                'return_6m': round(float(row[col_map['return_6m']]), 2) if not pd.isna(row[col_map['return_6m']]) else 0,
                'return_1y': round(float(row[col_map['return_1y']]), 2) if not pd.isna(row[col_map['return_1y']]) else 0,
                'score': float(row[col_map['score']]) if not pd.isna(row[col_map['score']]) else 0
            }
            funds_data.append(fund_info)

        # 按指数类型排序，再按综合得分排序
        funds_data.sort(key=lambda x: (x['index_type'], -x['score']))
        return funds_data

    except Exception as e:
        print(f"加载基金数据时出错：{e}")
        import traceback
        traceback.print_exc()
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
    获取基金数据的 API 接口
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
    启动 Web 服务器
    """
    print("=" * 50)
    print("指数基金筛选分析结果展示网站")
    print("=" * 50)
    print("1. 网站将在本地运行，地址：http://localhost:5000")
    print("2. 请确保端口 5000 未被占用")
    print("3. 网站启动后将自动打开浏览器")
    print("4. 按 Ctrl+C 停止服务器")
    print("=" * 50)

    # 在单独的线程中打开浏览器
    threading.Timer(1.0, open_browser).start()

    # 启动 Flask 应用
    app.run(host='localhost', port=5000, debug=False, use_reloader=False)


async def capture_website_screenshot():
    """
    使用 Playwright 截取网页屏幕截图并保存为 PNG 和 PDF 格式
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
            print(f"无法访问网站：{e}")
            print("请确保网站已在 http://localhost:5000 运行")
            await browser.close()
            return

        # 等待页面加载完成
        await page.wait_for_timeout(5000)

        # 截取整个页面的 PNG 截图
        print("正在截取 PNG 截图...")
        await page.screenshot(path='fund_analysis_screenshot.png', full_page=True)
        print("PNG 截图已保存为 fund_analysis_screenshot.png")

        # 生成 PDF 报告（横向页面，更大尺寸）
        print("正在生成 PDF 报告...")
        await page.pdf(
            path='fund_analysis_report.pdf',
            format='A4',
            print_background=True,
            landscape=True,  # 横向页面
            margin={"top": "1cm", "bottom": "1cm", "left": "1cm", "right": "1cm"}
        )
        print("PDF 报告已保存为 fund_analysis_report.pdf")

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
        print(f"截图过程中出现错误：{e}")
        print("请确保网站正在运行后再尝试截图")


def run_server_in_background():
    """
    在后台启动 Flask 服务器
    """
    def run_server():
        app.run(host='localhost', port=5000, debug=False, use_reloader=False)

    server_thread = threading.Thread(target=run_server, daemon=True)
    server_thread.start()
    return server_thread


def generate_report_auto():
    """
    一键生成报告：启动服务 → 等待加载 → 截图 → 生成 PDF → 关闭服务
    """
    print("=" * 60)
    print("指数基金分析报告生成")
    print("=" * 60)

    data_file = "filtered_top_funds_selection.xlsx"

    if not os.path.exists(data_file):
        print(f"错误：数据文件 {data_file} 不存在")
        print("请先运行 filtered_fund_selection.py 生成筛选数据")
        return

    print("1. 正在启动 Web 服务器...")
    server_thread = run_server_in_background()

    print("2. 等待服务器启动...")
    time.sleep(5)

    print("3. 正在生成截图和 PDF 报告...")

    try:
        asyncio.run(capture_website_screenshot())
        print("=" * 60)
        print("报告生成完成!")
        print("输出文件:")
        print("  - fund_analysis_screenshot.png (网页截图)")
        print("  - fund_analysis_report.pdf (PDF 报告)")
        print("=" * 60)
    except Exception as e:
        print(f"生成报告时出错：{e}")
    finally:
        print("4. 正在关闭服务...")
        time.sleep(1)


def main():
    parser = argparse.ArgumentParser(description='指数基金筛选分析网站')
    parser.add_argument('--mode', choices=['run', 'screenshot', 'auto'],
                        default='run', help='运行模式：run (启动网站) / screenshot (截取网站截图) / auto (一键生成报告)')

    args = parser.parse_args()

    if args.mode == 'run':
        run_website()
    elif args.mode == 'screenshot':
        take_screenshots()
    elif args.mode == 'auto':
        generate_report_auto()


if __name__ == '__main__':
    main()
