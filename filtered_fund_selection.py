import pandas as pd
import numpy as np

def meets_criteria(fund_row):
    """
    检查基金是否满足筛选条件
    
    Args:
        fund_row (Series): 基金数据行
    
    Returns:
        bool: 是否满足条件
    """
    try:
        # 条件1: 最新规模小于40亿
        scale_condition = True
        if '最新规模' in fund_row and not pd.isna(fund_row['最新规模']):
            # 解析规模数据，例如 "39.78亿元"
            scale_str = str(fund_row['最新规模']).replace('亿元', '')
            if scale_str.replace('.', '').isdigit():
                scale = float(scale_str)
                scale_condition = scale < 40
        
        # 条件2: 换手率大于200%
        turnover_condition = True
        if '换手率' in fund_row and not pd.isna(fund_row['换手率']):
            # 解析换手率数据，例如 "176.45%"
            turnover_str = str(fund_row['换手率']).replace('%', '')
            if turnover_str.replace('.', '').isdigit():
                turnover = float(turnover_str)
                turnover_condition = turnover > 200
        
        # 条件3: 前10大重仓股占比小于40%
        concentration_condition = True
        if '前10大重仓股占比' in fund_row and not pd.isna(fund_row['前10大重仓股占比']):
            # 解析重仓股占比数据，例如 "46.56%"
            concentration_str = str(fund_row['前10大重仓股占比']).replace('%', '')
            if concentration_str.replace('.', '').isdigit():
                concentration = float(concentration_str)
                concentration_condition = concentration < 40
        
        return scale_condition and turnover_condition and concentration_condition
        
    except Exception as e:
        # 如果解析过程中出现异常，默认认为满足条件
        print(f"解析基金条件时出错: {e}")
        return True

def calculate_10point_score(fund_row, cols):
    """
    根据时间衰减权重和负值惩罚计算10分制综合得分
    
    Args:
        fund_row (Series): 基金数据行
        cols (list): 超额收益列名列表
    
    Returns:
        float: 10分制综合得分（考虑负值惩罚）
    """
    # 定义时间衰减权重（越近期权重越高）
    weights = {
        '近1月超额': 0.3,   # 权重最高，因为最能反映当前状态
        '近3月超额': 0.25,
        '近6月超额': 0.25,
        '近1年超额': 0.2    # 权重相对较低，因为时间较久远
    }
    
    weighted_sum = 0
    weight_sum = 0
    
    for col in cols:
        if col in fund_row and not pd.isna(fund_row[col]):
            value = fund_row[col]
            weight = weights[col]
            
            # 如果超额收益为负，则施加惩罚
            if value < 0:
                # 对负值进行非线性惩罚，数值越负惩罚越重
                penalty_factor = 1 + abs(value) / 50  # 惩罚系数，负值越大惩罚越重
                adjusted_value = value * penalty_factor
            else:
                adjusted_value = value
            
            weighted_sum += adjusted_value * weight
            weight_sum += weight
    
    # 如果没有任何有效数据，返回NaN
    if weight_sum == 0:
        return np.nan
    
    # 计算原始加权得分
    raw_score = weighted_sum / weight_sum
    
    # 将得分转换为10分制
    # 假设合理的得分区间为 [-10, 20]，映射到 [0, 10]
    min_score = -10
    max_score = 20
    score_10 = max(0, min(10, 10 * (raw_score - min_score) / (max_score - min_score)))
    
    return round(score_10, 2)

def select_top_funds_from_sheet(sheet_df, index_name):
    """
    从单个工作表中选择满足条件且表现最好的三只基金
    
    Args:
        sheet_df (DataFrame): 工作表数据
        index_name (str): 指数名称
    
    Returns:
        DataFrame: 选出的基金数据
    """
    # 定义要分析的超额收益列
    excess_return_cols = ['近1月超额', '近3月超额', '近6月超额', '近1年超额']
    
    # 检查是否存在这些列
    available_cols = [col for col in excess_return_cols if col in sheet_df.columns]
    
    if not available_cols:
        print(f"{index_name} 工作表中没有找到超额收益列")
        return pd.DataFrame()
    
    # 移除所有超额收益列都为空值的行
    filtered_df = sheet_df.dropna(subset=available_cols, how='all')
    
    if len(filtered_df) == 0:
        print(f"{index_name} 工作表中没有有效的数据")
        return pd.DataFrame()
    
    # 应用筛选条件
    print(f"{index_name} 工作表中共有 {len(filtered_df)} 只基金，开始应用筛选条件...")
    
    # 应用筛选条件
    condition_met_df = filtered_df[filtered_df.apply(meets_criteria, axis=1)]
    
    print(f"满足条件的基金数量: {len(condition_met_df)}")
    
    if len(condition_met_df) == 0:
        print(f"{index_name} 工作表中没有满足筛选条件的基金")
        return pd.DataFrame()
    
    # 计算10分制综合得分（考虑负值惩罚）
    condition_met_df = condition_met_df.copy()
    condition_met_df['综合得分'] = condition_met_df.apply(
        lambda row: calculate_10point_score(row, available_cols), axis=1
    )
    
    # 移除综合得分为空的行
    condition_met_df = condition_met_df.dropna(subset=['综合得分'])
    
    if len(condition_met_df) == 0:
        print(f"{index_name} 工作表中没有可以计算得分的数据")
        return pd.DataFrame()
    
    # 按综合得分降序排列，取前三名（或所有满足条件的基金，如果不足三个）
    top_funds = condition_met_df.nlargest(3, '综合得分')
    
    # 添加指数名称列
    top_funds['指数类型'] = index_name
    
    # 选择需要显示的列
    display_cols = ['指数类型', '基金代码', '基金简称', '最新规模', '换手率', '前10大重仓股占比'] + available_cols + ['综合得分']
    result = top_funds[display_cols].reset_index(drop=True)
    
    return result

def select_top_funds(file_path):
    """
    从Excel文件的所有"_超额"工作表中选择满足条件且表现最好的基金
    
    Args:
        file_path (str): Excel文件路径
    
    Returns:
        DataFrame: 所有指数类型中选出的基金
    """
    try:
        # 读取所有工作表
        all_sheets = pd.read_excel(file_path, sheet_name=None)
        
        # 筛选出带"_超额"后缀的工作表
        excess_sheets = {name: df for name, df in all_sheets.items() if name.endswith('_超额')}
        
        print(f"找到 {len(excess_sheets)} 个超额收益工作表:")
        for sheet_name in excess_sheets.keys():
            print(f"- {sheet_name}")
        
        # 存储所有选出的基金
        all_top_funds = []
        
        # 为每个工作表选择满足条件的基金
        for sheet_name, sheet_data in excess_sheets.items():
            # 提取指数名称（去除"_超额"后缀）
            index_name = sheet_name.replace('_超额', '')
            print(f"\n正在处理 {sheet_name}...")
            
            # 选择该工作表中满足条件且表现最好的基金
            top_funds = select_top_funds_from_sheet(sheet_data, index_name)
            
            if not top_funds.empty:
                all_top_funds.append(top_funds)
                print(f"已选出 {len(top_funds)} 只基金")
            else:
                print(f"未能从 {sheet_name} 中选出满足条件的基金")
        
        # 合并所有结果
        if all_top_funds:
            final_result = pd.concat(all_top_funds, ignore_index=True)
            return final_result
        else:
            print("未选出任何满足条件的基金")
            return pd.DataFrame()
            
    except FileNotFoundError:
        print(f"错误: 找不到文件 {file_path}")
        return pd.DataFrame()
    except Exception as e:
        print(f"处理文件时发生错误: {e}")
        return pd.DataFrame()

def show_filter_criteria():
    """
    显示筛选条件说明
    """
    print("基金筛选条件:")
    print("="*50)
    print("1. 最新规模 < 40亿")
    print("2. 换手率 > 200%")
    print("3. 前10大重仓股占比 < 40%")
    print("="*50)

def main():
    """主函数"""
    file_path = "index-fund.xlsx"
    
    # 显示筛选条件
    show_filter_criteria()
    
    print("\n开始根据筛选条件从各指数类型的超额收益基金中选择最佳基金...")
    top_funds = select_top_funds(file_path)
    
    if not top_funds.empty:
        print("\n" + "="*120)
        print("满足筛选条件的指数类型超额收益基金（10分制评分）")
        print("="*120)
        
        # 按综合得分降序排列
        top_funds_sorted = top_funds.sort_values('综合得分', ascending=False)
        
        # 显示结果
        for _, fund in top_funds_sorted.iterrows():
            print(f"\n【{fund['指数类型']}】{fund['基金简称']} ({fund['基金代码']})")
            print(f"   规模: {fund['最新规模']}  换手率: {fund['换手率']}  重仓股占比: {fund['前10大重仓股占比']}")
            print(f"   近1月超额: {fund['近1月超额']:.2f}%  "
                  f"近3月超额: {fund['近3月超额']:.2f}%  "
                  f"近6月超额: {fund['近6月超额']:.2f}%  "
                  f"近1年超额: {fund['近1年超额']:.2f}%")
            print(f"   综合得分: {fund['综合得分']}/10")
        
        # 保存结果到Excel文件
        output_file = "filtered_top_funds_selection.xlsx"
        top_funds_sorted.to_excel(output_file, index=False)
        print(f"\n结果已保存到 {output_file}")
    else:
        print("未能选出满足筛选条件的基金")

if __name__ == "__main__":
    main()