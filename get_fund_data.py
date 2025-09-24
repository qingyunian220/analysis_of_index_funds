from flask import Flask, jsonify, request, render_template
import pandas as pd
import os
import json
import numpy as np

app = Flask(__name__)
app.config['JSON_AS_ASCII'] = False

# 设置Excel文件路径
EXCEL_FILE = '基金C份额收益率排名_20250924.xlsx'

@app.route('/')
def index():
    return render_template('fund_ranking.html')

@app.route('/get_fund_data')
def get_fund_data():
    # 获取参数
    sheet_name = request.args.get('sheet', '沪深300基金')
    page = int(request.args.get('page', 1))
    limit = int(request.args.get('limit', 20))
    
    # 检查文件是否存在
    if not os.path.exists(EXCEL_FILE):
        return jsonify({'code': 1, 'msg': 'Excel文件不存在', 'count': 0, 'data': []})
    
    try:
        # 读取Excel文件中的指定工作表
        df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
        
        # 处理数据
        # 将所有列名中的空格去掉
        df.columns = df.columns.str.replace(' ', '', regex=False)
        
        # 处理 NaN 值，将其替换为 None（在 JSON 中会变成 null）
        df = df.replace({np.nan: None})
        
        # 处理日期格式，确保能正确序列化
        date_columns = ['日期']
        for col in date_columns:
            if col in df.columns:
                df[col] = df[col].astype(str)
        
        # 处理基金代码，确保为6位字符串格式
        fund_code_columns = ['基金代码']
        for col in fund_code_columns:
            if col in df.columns:
                # 确保基金代码为字符串并格式化为6位
                df[col] = df[col].apply(lambda x: str(int(x)).zfill(6) if isinstance(x, (int, float)) and not pd.isna(x) else str(x).zfill(6) if x is not None else '')
        # 分页处理
        total = len(df)
        start = (page - 1) * limit
        end = start + limit
        df_page = df.iloc[start:end]
        
        # 转换为字典列表
        data = df_page.to_dict('records')
        
        # 手动构建 JSON 响应以确保中文正确显示
        response_data = {
            'code': 0,
            'msg': 'success',
            'count': total,
            'data': data
        }
        
        # 使用 ensure_ascii=False 参数确保中文正确显示
        response_json = json.dumps(response_data, ensure_ascii=False)
        
        # 返回响应
        return app.response_class(
            response=response_json,
            status=200,
            mimetype='application/json'
        )
        
    except Exception as e:
        response_data = {'code': 1, 'msg': str(e), 'count': 0, 'data': []}
        response_json = json.dumps(response_data, ensure_ascii=False)
        return app.response_class(
            response=response_json,
            status=200,
            mimetype='application/json'
        )

if __name__ == '__main__':
    app.run(debug=True, port=5000)