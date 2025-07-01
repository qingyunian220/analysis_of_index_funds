import pandas as pd
import akshare as ak
import time
from tqdm import tqdm

# 获取所有基金基础信息
fund_open_fund_rank_em_df = ak.fund_open_fund_rank_em(symbol="全部")

return_columns = ['近1月', '近3月', '近6月', '近1年', '今年来']

def fetch_fund_data(fund_name):
    fund_df = fund_open_fund_rank_em_df[fund_open_fund_rank_em_df["基金简称"].str.contains(fund_name, na=False)]
    fund_df = fund_df[fund_df["基金简称"].str.contains("C", na=False)]
    
    exclude_keywords = ["红利", "基本面", "价值", "非银", "成长", "低波动"]
    for keyword in tqdm(exclude_keywords):
        fund_df = fund_df[~fund_df["基金简称"].str.contains(keyword, na=False)]
    
    fund_df["成立时间"] = ""
    fund_df["最新规模"] = ""
    
    for idx, row in tqdm(fund_df.iterrows(), total=fund_df.shape[0]):
        code = row["基金代码"]
        try:
            info = ak.fund_individual_basic_info_xq(symbol=code)
            if "成立时间" in info["item"].values:
                fund_df.at[idx, "成立时间"] = info.loc[info["item"] == "成立时间", "value"].values[0]
            if "最新规模" in info["item"].values:
                fund_df.at[idx, "最新规模"] = info.loc[info["item"] == "最新规模", "value"].values[0]
            time.sleep(0.1)
        except Exception as e:
            print(f"基金代码{code}查询失败: {e}")
    
    return fund_df.sort_values(by='近6月', ascending=False)



def highlight_top_50_all_columns(df):
    # 创建一个样式DataFrame，默认为空字符串（无样式）
    styles = pd.DataFrame('', index=df.index, columns=df.columns)
    
    # 记录每个基金在多少个收益率列中进入前10
    top_count = {idx: 0 for idx in df.index}
    
    # 对每个收益率列，标记其前10名
    for col in return_columns:
        if col in df.columns:
            # 获取该列排序后的前10个索引
            top_10_idx = df[col].nlargest(10).index
            # 将对应位置设置为黄色背景
            styles.loc[top_10_idx, col] = 'background-color: yellow'
            # 更新每个基金进入前10的次数
            for idx in top_10_idx:
                top_count[idx] += 1
    
    # 对至少有4列进入前10的基金，将基金简称设置为金黄色背景
    for idx, count in top_count.items():
        if count >= 4:  # 至少有4列进入前10
            # 标注基金简称为金黄色
            styles.loc[idx, '基金简称'] = 'background-color: gold'
    
    return styles

def save_to_excel(writer, fund_df, sheet_name):
    if not fund_df.empty:  # 确保DataFrame不为空
        styled_df = fund_df.style.apply(highlight_top_50_all_columns, axis=None)
        styled_df.to_excel(writer, sheet_name=sheet_name, index=False)

# 创建一个ExcelWriter对象
with pd.ExcelWriter('所有基金C份额收益率排名.xlsx', engine='openpyxl') as writer:
    fund_types = ["沪深300", "中证500", "中证1000", "中证2000"]
    for fund_type in fund_types:
        fund_df = fetch_fund_data(fund_type)
        save_to_excel(writer, fund_df, f'{fund_type}基金')

    # 自适应列宽
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        for column in worksheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

print("已将所有C份额基金的排序结果保存为'所有基金C份额收益率排名.xlsx'，每个时间段的前10名标黄，至少有4个时间段进入前10的基金其简称标金黄色。") 