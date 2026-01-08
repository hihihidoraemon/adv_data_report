import streamlit as st
import pandas as pd
import numpy as np
import re
import os
from datetime import datetime, date, timedelta
import base64
from io import BytesIO
import tempfile

# ==================== Streamlit页面配置（必须放在最前面） ====================
st.set_page_config(
    page_title="网盟日报输出",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== 样式配置 ====================
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
    .warning-box {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
    .stProgress > div > div > div > div {
        background-color: #1f77b4;
    }
    .upload-area {
        border: 2px dashed #ccc;
        border-radius: 10px;
        padding: 30px;
        text-align: center;
        margin: 20px 0;
        background-color: #f9f9f9;
    }
</style>
""", unsafe_allow_html=True)

# ==================== 核心处理函数 ====================
def process_daily_report_web(uploaded_file, progress_bar=None, status_text=None):
    """
    网页版处理日报Excel数据的主函数
    """
    
    # 更新进度
    if progress_bar and status_text:
        progress_bar.progress(5)
        status_text.text("📁 正在读取Excel文件...")
    
    # ====================== 1、导入Excel数据 ======================
    try:
        sheet1_all_data = pd.read_excel(uploaded_file, sheet_name='1--all data')
        sheet3_advertiser = pd.read_excel(uploaded_file, sheet_name='3--匹配广告主')
        sheet4_reject = pd.read_excel(uploaded_file, sheet_name='4--reject事件')
        sheet2_reject_rule = pd.read_excel(uploaded_file, sheet_name='2-reject规则')
        
        if progress_bar and status_text:
            progress_bar.progress(15)
            status_text.text("✅ 成功读取Excel文件，开始数据处理...")
            
    except Exception as e:
        raise Exception(f"读取文件失败：{str(e)}")
    
    # ====================== 关键优化：自动识别最新两天日期 ======================
    if progress_bar and status_text:
        progress_bar.progress(20)
        status_text.text("📅 自动识别最新两天日期...")
    
    sheet1_all_data['Date'] = sheet1_all_data['Time'].dt.date
    sheet4_reject['Date'] = sheet4_reject['Time'].dt.date
    
    # 获取所有唯一日期并排序
    all_dates = sorted(sheet1_all_data['Date'].unique(), reverse=True)
    
    if len(all_dates) < 2:
        raise Exception(f"错误：数据中仅包含 {len(all_dates)} 天数据，至少需要2天！")
    
    # 定义最新两天（核心日期变量）
    newest_date = all_dates[0]       # 最新一天
    second_newest_date = all_dates[1] # 次新一天
    
    # 生成日期显示名称
    newest_date_str = f"{newest_date.year}/{newest_date.month}/{newest_date.day}"
    second_newest_date_str = f"{second_newest_date.year}/{second_newest_date.month}/{second_newest_date.day}"
    newest_date_file_str = f"{newest_date.year}{newest_date.month:02d}{newest_date.day:02d}"
    
    date_mapping = {
        'newest': {
            'date': newest_date,
            'str': newest_date_str,
            'file_str': newest_date_file_str,
            'col_name': f"{newest_date_str} Total Revenue",
            'reject_rate_col': f"{newest_date_str} reject率(%)"
        },
        'second': {
            'date': second_newest_date,
            'str': second_newest_date_str,
            'col_name': f"{second_newest_date_str} Total Revenue",
            'reject_rate_col': f"{second_newest_date_str} reject率(%)"
        }
    }
    
    # ====================== 2、基础数据预处理 ======================
    if progress_bar and status_text:
        progress_bar.progress(30)
        status_text.text("🔧 基础数据预处理...")
    
    # 提取每个Offer ID的最新Status
    offer_status_mapping = sheet1_all_data[sheet1_all_data['Date'] == newest_date][
        ['Offer ID', 'Status']
    ].drop_duplicates(subset=['Offer ID']).fillna('Unknown')
    
    # 精准判断新旧预算
    non_newest_data = sheet1_all_data[sheet1_all_data['Date'] != newest_date].copy()
    six_days_ago = newest_date - timedelta(days=6)
    past_6_days_data = non_newest_data[non_newest_data['Date'] >= six_days_ago].copy()
    old_budget_offers = set(
        past_6_days_data[past_6_days_data['Total Revenue'] > 0]['Offer ID'].unique()
    )
    all_offers = set(sheet1_all_data['Offer ID'].unique())
    
    def judge_budget_type(offer_id):
        return '旧预算' if offer_id in old_budget_offers else '新预算'
    
    # ====================== 3、匹配广告主信息 ======================
    if progress_bar and status_text:
        progress_bar.progress(40)
        status_text.text("🔗 匹配广告主信息...")
    
    sheet1_all_data = pd.merge(
        sheet1_all_data, 
        sheet3_advertiser[['Advertiser', '二级广告主', '三级广告主']], 
        on='Advertiser', 
        how='left'
    )
    
    # ====================== 4、核心计算：Offer级别的基础数据 ======================
    if progress_bar and status_text:
        progress_bar.progress(50)
        status_text.text("📊 计算Offer级别数据...")
    
    # 提取App ID映射
    offer_app_mapping = sheet1_all_data[['Offer ID', 'App ID']].drop_duplicates(subset=['Offer ID']).fillna('')
    
    # 计算每个Offer ID在最新/次新一天的总收入
    offer_newest_revenue = sheet1_all_data[sheet1_all_data['Date'] == newest_date].groupby('Offer ID').agg({
        'Total Revenue': 'sum'
    }).reset_index()
    offer_newest_revenue.columns = ['Offer ID', date_mapping['newest']['col_name']]
    
    offer_second_revenue = sheet1_all_data[sheet1_all_data['Date'] == second_newest_date].groupby('Offer ID').agg({
        'Total Revenue': 'sum'
    }).reset_index()
    offer_second_revenue.columns = ['Offer ID', date_mapping['second']['col_name']]
    
    # 合并Offer基础数据
    offer_base_data = offer_app_mapping.copy()
    offer_base_data = pd.merge(offer_base_data, offer_status_mapping, on='Offer ID', how='left')
    offer_base_data = pd.merge(offer_base_data, offer_newest_revenue, on='Offer ID', how='left').fillna(0)
    offer_base_data = pd.merge(offer_base_data, offer_second_revenue, on='Offer ID', how='left').fillna(0)
    
    # 计算Offer级流水差
    offer_base_data['流水差（最新-次新）'] = (
        offer_base_data[date_mapping['newest']['col_name']] - 
        offer_base_data[date_mapping['second']['col_name']]
    )
    
    def calculate_offer_change_pct(row):
        prev_revenue = row[date_mapping['second']['col_name']]
        curr_revenue = row[date_mapping['newest']['col_name']]
        if prev_revenue == 0:
            return 1000.0 if curr_revenue > 0 else 0.0
        return ((curr_revenue - prev_revenue) / abs(prev_revenue)) * 100
    
    offer_base_data['变化幅度(%)'] = offer_base_data.apply(calculate_offer_change_pct, axis=1)
    offer_base_data['预算类型'] = offer_base_data['Offer ID'].apply(judge_budget_type)
    
    # 高差异Offer筛选
    high_diff_mask = (offer_base_data['流水差（最新-次新）'].abs() >= 10)
    high_diff_offers = offer_base_data[high_diff_mask]['Offer ID'].tolist()
    
    # ====================== 5、Affiliate维度精准分析 ======================
    if progress_bar and status_text:
        progress_bar.progress(60)
        status_text.text("👥 Affiliate维度分析...")
    
    offer_influence = pd.DataFrame(columns=['Offer ID', 'influence affiliate'])
    
    if high_diff_offers:
        # 按Offer ID + Affiliate + Date分组计算
        affiliate_daily_metrics = sheet1_all_data[sheet1_all_data['Offer ID'].isin(high_diff_offers)].groupby(
            ['Offer ID', 'Affiliate', 'Date']
        ).agg({
            'Total Revenue': 'sum',
            'Total Clicks': 'sum',
            'Total Conversions': 'sum'
        }).reset_index()
        
        # 分别提取最新/次新一天数据
        aff_newest = affiliate_daily_metrics[affiliate_daily_metrics['Date'] == newest_date].copy()
        aff_newest.columns = ['Offer ID', 'Affiliate', 'Date', 'Revenue_newest', 'Clicks_newest', 'Conversions_newest']
        
        aff_second = affiliate_daily_metrics[affiliate_daily_metrics['Date'] == second_newest_date].copy()
        aff_second.columns = ['Offer ID', 'Affiliate', 'Date', 'Revenue_second', 'Clicks_second', 'Conversions_second']
        
        # 合并两天数据
        aff_merged = pd.merge(
            aff_newest, aff_second, 
            on=['Offer ID', 'Affiliate'], 
            how='outer'
        ).fillna(0)
        
        # 计算差异指标
        aff_merged['Revenue_Diff'] = aff_merged['Revenue_newest'] - aff_merged['Revenue_second']
        aff_merged['Clicks_Diff'] = aff_merged['Clicks_newest'] - aff_merged['Clicks_second']
        aff_merged['Clicks_Change_Pct'] = np.where(
            aff_merged['Clicks_second'] > 0,
            (aff_merged['Clicks_Diff'] / aff_merged['Clicks_second']) * 100,
            np.where(aff_merged['Clicks_newest'] > 0, 1000.0, 0.0)
        )
        
        # CR计算
        aff_merged['CR_newest'] = np.where(
            aff_merged['Clicks_newest'] > 0,
            (aff_merged['Conversions_newest'] / aff_merged['Clicks_newest']) * 100,
            0.0
        )
        aff_merged['CR_second'] = np.where(
            aff_merged['Clicks_second'] > 0,
            (aff_merged['Conversions_second'] / aff_merged['Clicks_second']) * 100,
            0.0
        )
        aff_merged['CR_Change_Abs'] = aff_merged['CR_newest'] - aff_merged['CR_second']
        
        # 筛选有显著收入变化的Affiliate
        significant_aff = aff_merged[aff_merged['Revenue_Diff'].abs() >= 5].copy()
        significant_aff = significant_aff.sort_values(by='Revenue_Diff', ascending=False)
        
        def generate_influence_text(row):
            affiliate = row['Affiliate']
            revenue_newest = row['Revenue_newest']
            revenue_second = row['Revenue_second']
            revenue_diff = row['Revenue_Diff']
            clicks_change = row['Clicks_Change_Pct']
            cr_change = row['CR_Change_Abs']
            
            if revenue_newest > 0 and revenue_second == 0:
                return f"{affiliate} 新增产生流水 {revenue_newest:.2f} 美金"
            elif revenue_newest == 0 and revenue_second > 0:
                return f"{affiliate} 停止产生流水，减少 {revenue_second:.2f} 美金"
            else:
                if revenue_second != 0:
                    revenue_change_pct = (revenue_diff / abs(revenue_second)) * 100
                else:
                    revenue_change_pct = 1000.0 if revenue_diff > 0 else -1000.0
                
                if revenue_diff > 0:
                    base_text = f"{affiliate} 增加 {revenue_diff:.2f} 美金/{abs(revenue_change_pct):.1f}%"
                else:
                    base_text = f"{affiliate} 减少 {abs(revenue_diff):.2f} 美金/{abs(revenue_change_pct):.1f}%"
                
                reasons = []
                direction = "增加" if clicks_change > 0 else "减少"
                reasons.append(f"Total Clicks{direction}{abs(clicks_change):.1f}%")
                direction = "增加" if cr_change > 0 else "减少"
                reasons.append(f"CR{direction}{abs(cr_change):.1f}%")
                
                return f"{base_text}，对应{', '.join(reasons)}"
        
        significant_aff['influence_text'] = significant_aff.apply(generate_influence_text, axis=1)
        offer_influence = significant_aff.groupby('Offer ID')['influence_text'].apply(
            lambda x: '\n'.join(x)
        ).reset_index()
        offer_influence.columns = ['Offer ID', 'influence affiliate']
    
    # ====================== 6、生成四个核心表格 ======================
    if progress_bar and status_text:
        progress_bar.progress(70)
        status_text.text("📈 生成核心分析表格...")
    
    # 表格一：三级广告主日报表
    table1_data = sheet1_all_data[sheet1_all_data['Date'].isin([newest_date, second_newest_date])].groupby(
        ['三级广告主', 'Date']
    ).agg({
        'Total Revenue': 'sum',
        'Total Profit': 'sum'
    }).reset_index()
    
    table1 = pd.DataFrame()
    table1['三级广告主'] = table1_data['三级广告主'].unique()
    
    for date_type in ['newest', 'second']:
        current_date = date_mapping[date_type]['date']
        current_date_str = date_mapping[date_type]['str']
        temp = table1_data[table1_data['Date'] == current_date].set_index('三级广告主')
        table1[f"{current_date_str} Total Revenue"] = table1['三级广告主'].map(temp['Total Revenue']).fillna(0)
        table1[f"{current_date_str} Total Profit"] = table1['三级广告主'].map(temp['Total Profit']).fillna(0)
    
    table1 = table1[
        ['三级广告主', 
         f"{newest_date_str} Total Revenue", f"{newest_date_str} Total Profit",
         f"{second_newest_date_str} Total Revenue", f"{second_newest_date_str} Total Profit"]
    ].copy().round(2)
    
    # 表格二：高差异Offer ID详情
    if high_diff_offers:
        offer_details = sheet1_all_data[sheet1_all_data['Offer ID'].isin(high_diff_offers)][
            ['Offer ID', 'GEO', 'Advertiser']
        ].drop_duplicates(subset=['Offer ID']).reset_index(drop=True)
        
        table2 = pd.merge(offer_details, offer_base_data[
            ['Offer ID', 'App ID', 'Status', date_mapping['newest']['col_name'], 
             date_mapping['second']['col_name'], '流水差（最新-次新）', '变化幅度(%)', '预算类型']
        ], on='Offer ID', how='left')
        
        table2 = pd.merge(table2, offer_influence, on='Offer ID', how='left')
        table2['influence affiliate'] = table2['influence affiliate'].fillna('无显著变化')
        
        table2 = table2[
            ['Offer ID', 'App ID', 'Status', 'GEO', 'Advertiser',
             date_mapping['newest']['col_name'], date_mapping['second']['col_name'],
             '流水差（最新-次新）', '变化幅度(%)', '预算类型', 'influence affiliate']
        ].copy()
        
        numeric_cols_table2 = [
            date_mapping['newest']['col_name'], date_mapping['second']['col_name'],
            '流水差（最新-次新）', '变化幅度(%)'
        ]
        table2[numeric_cols_table2] = table2[numeric_cols_table2].round(2)
    else:
        table2 = pd.DataFrame(columns=[
            'Offer ID', 'App ID', 'Status', 'GEO', 'Advertiser',
            date_mapping['newest']['col_name'], date_mapping['second']['col_name'],
            '流水差（最新-次新）', '变化幅度(%)', '预算类型', 'influence affiliate'
        ])
    
     # ---------------------- 表格三：二级广告主综合报表（新增reject率） ----------------------
    print("核心新增：表格三计算二级广告主reject率...")
    table3 = pd.DataFrame()
    table3['二级广告主'] = sheet1_all_data['二级广告主'].unique()
    
    # 填充收入/利润/转化数据
    for date_type in ['newest', 'second']:
        current_date = date_mapping[date_type]['date']
        current_date_str = date_mapping[date_type]['str']
        
        temp = sheet1_all_data[sheet1_all_data['Date'] == current_date].groupby('二级广告主').agg({
            'Total Revenue': 'sum',
            'Total Profit': 'sum',
            'Total Conversions': 'sum'
        }).reset_index()
        
        table3[f"{current_date_str} Total Revenue"] = table3['二级广告主'].map(temp.set_index('二级广告主')['Total Revenue']).fillna(0)
        table3[f"{current_date_str} Total Profit"] = table3['二级广告主'].map(temp.set_index('二级广告主')['Total Profit']).fillna(0)
        table3[f"{current_date_str} Total Conversions"] = table3['二级广告主'].map(temp.set_index('二级广告主')['Total Conversions']).fillna(0)
    
    # 处理4--reject事件数据
    sheet4_reject = pd.merge(
        sheet4_reject, sheet3_advertiser[['Advertiser', '二级广告主']], 
        on='Advertiser', how='left'
    )
    sheet4_reject['New Time'] = sheet4_reject['Time'].copy()
    appnext_mask = sheet4_reject['Advertiser'].str.contains('appnext', case=False, na=False)
    sheet4_reject.loc[appnext_mask, 'New Time'] = sheet4_reject.loc[appnext_mask, 'New Time'] - timedelta(days=1)
    sheet4_reject['New Date'] = sheet4_reject['New Time'].dt.date
    sheet4_reject = pd.merge(
        sheet4_reject, sheet2_reject_rule[['Event', '是否为reject']], 
        on='Event', how='left'
    )
    
    # 填充Reject数据
    reject_stats = sheet4_reject[sheet4_reject['New Date'].isin([newest_date, second_newest_date])].groupby(
        ['New Date', '二级广告主']
    ).agg({
        '是否为reject': lambda x: (x == True).sum()
    }).reset_index()
    
    for date_type in ['newest', 'second']:
        current_date = date_mapping[date_type]['date']
        current_date_str = date_mapping[date_type]['str']
        
        temp = reject_stats[reject_stats['New Date'] == current_date].set_index('二级广告主')
        table3[f"{current_date_str} Total reject"] = table3['二级广告主'].map(temp['是否为reject']).fillna(0)
    
    # ========== 核心新增：计算二级广告主reject率 ==========
    def calculate_reject_rate(row, date_str):
        """
        计算reject率：reject / (conversions + reject)
        分母为0时返回0，避免除以0错误
        """
        conversions = row[f"{date_str} Total Conversions"]
        reject = row[f"{date_str} Total reject"]
        total = conversions + reject
        if total == 0:
            return 0.0
        return (reject / total) * 100
    
    # 计算最新/次新一天的reject率
    table3[date_mapping['newest']['reject_rate_col']] = table3.apply(
        lambda x: calculate_reject_rate(x, newest_date_str), axis=1
    ).round(2)
    
    table3[date_mapping['second']['reject_rate_col']] = table3.apply(
        lambda x: calculate_reject_rate(x, second_newest_date_str), axis=1
    ).round(2)
    
    # 调整列顺序并格式化
    table3 = table3[
        ['二级广告主', 
         f"{newest_date_str} Total Revenue", f"{newest_date_str} Total Profit",
         f"{second_newest_date_str} Total Revenue", f"{second_newest_date_str} Total Profit",
         f"{newest_date_str} Total Conversions", f"{newest_date_str} Total reject", date_mapping['newest']['reject_rate_col'],
         f"{second_newest_date_str} Total Conversions", f"{second_newest_date_str} Total reject", date_mapping['second']['reject_rate_col']]
    ].copy()
    
    numeric_cols_table3 = [f"{newest_date_str} Total Revenue", f"{newest_date_str} Total Profit",
                          f"{second_newest_date_str} Total Revenue", f"{second_newest_date_str} Total Profit",
                          date_mapping['newest']['reject_rate_col'], date_mapping['second']['reject_rate_col']]
    table3[numeric_cols_table3] = table3[numeric_cols_table3].round(2)
    
    int_cols_table3 = [f"{newest_date_str} Total Conversions", f"{newest_date_str} Total reject",
                      f"{second_newest_date_str} Total Conversions", f"{second_newest_date_str} Total reject"]
    table3[int_cols_table3] = table3[int_cols_table3].astype(int)
    
    # ---------------------- 表格四：Affiliate综合报表（新增reject率） ----------------------
    print("核心新增：表格四计算Affiliate reject率...")
    table4 = pd.DataFrame()
    table4['Affiliate'] = sheet1_all_data['Affiliate'].unique()
    
    # 动态填充两天的收入/利润/转化数据
    for date_type in ['newest', 'second']:
        current_date = date_mapping[date_type]['date']
        current_date_str = date_mapping[date_type]['str']
        
        daily_data = sheet1_all_data[sheet1_all_data['Date'] == current_date].groupby('Affiliate').agg({
            'Total Revenue': 'sum',
            'Total Profit': 'sum',
            'Total Conversions': 'sum',
            '二级广告主': lambda x: x.mode()[0] if not x.mode().empty else ''
        }).reset_index()
        
        table4[f"{current_date_str} Total Revenue"] = table4['Affiliate'].map(daily_data.set_index('Affiliate')['Total Revenue']).fillna(0)
        table4[f"{current_date_str} Total Profit"] = table4['Affiliate'].map(daily_data.set_index('Affiliate')['Total Profit']).fillna(0)
        table4[f"{current_date_str} Total Conversions"] = table4['Affiliate'].map(daily_data.set_index('Affiliate')['Total Conversions']).fillna(0)
        table4[f"{current_date_str} 二级广告主"] = table4['Affiliate'].map(daily_data.set_index('Affiliate')['二级广告主']).fillna('')
    
    # 合并二级广告主信息
    def merge_advertisers(row):
        adv1 = row[f"{second_newest_date_str} 二级广告主"]
        adv2 = row[f"{newest_date_str} 二级广告主"]
        advs = set()
        if adv1 and adv1 != '0':
            advs.add(str(adv1))
        if adv2 and adv2 != '0':
            advs.add(str(adv2))
        return '; '.join(advs)
    
    table4['二级广告主'] = table4.apply(merge_advertisers, axis=1)
    
    # 填充Reject数据
    reject_long = pd.melt(
        table3[['二级广告主', f"{newest_date_str} Total reject", f"{second_newest_date_str} Total reject"]],
        id_vars=['二级广告主'],
        var_name='Date',
        value_name='Total reject'
    )
    reject_long['Date'] = reject_long['Date'].str.extract(r'(\d{4}/\d{1,2}/\d{1,2})')
    
    def get_affiliate_reject(row, target_date_str):
        if not row['二级广告主']:
            return 0
        total_reject = 0
        for adv in row['二级广告主'].split('; '):
            adv = adv.strip()
            reject_val = reject_long[
                (reject_long['二级广告主'] == adv) & 
                (reject_long['Date'] == target_date_str)
            ]['Total reject'].sum()
            total_reject += reject_val
        return total_reject
    
    # 添加reject列
    table4[f"{newest_date_str} Total reject"] = table4.apply(
        lambda x: get_affiliate_reject(x, newest_date_str), axis=1
    ).astype(int)
    
    table4[f"{second_newest_date_str} Total reject"] = table4.apply(
        lambda x: get_affiliate_reject(x, second_newest_date_str), axis=1
    ).astype(int)
    
    # ========== 核心新增：计算Affiliate reject率 ==========
    table4[date_mapping['newest']['reject_rate_col']] = table4.apply(
        lambda x: calculate_reject_rate(x, newest_date_str), axis=1
    ).round(2)
    
    table4[date_mapping['second']['reject_rate_col']] = table4.apply(
        lambda x: calculate_reject_rate(x, second_newest_date_str), axis=1
    ).round(2)
    
    # 调整列顺序并格式化
    table4 = table4[
        ['Affiliate', 
         f"{newest_date_str} Total Revenue", f"{newest_date_str} Total Profit",
         f"{second_newest_date_str} Total Revenue", f"{second_newest_date_str} Total Profit",
         f"{newest_date_str} Total Conversions", f"{newest_date_str} Total reject", date_mapping['newest']['reject_rate_col'],
         f"{second_newest_date_str} Total Conversions", f"{second_newest_date_str} Total reject", date_mapping['second']['reject_rate_col'],
         '二级广告主']
    ].copy()
    
    table4 = table4.fillna(0)
    numeric_cols_table4 = [f"{newest_date_str} Total Revenue", f"{newest_date_str} Total Profit",
                          f"{second_newest_date_str} Total Revenue", f"{second_newest_date_str} Total Profit",
                          date_mapping['newest']['reject_rate_col'], date_mapping['second']['reject_rate_col']]
    table4[numeric_cols_table4] = table4[numeric_cols_table4].round(2)
    
    int_cols_table4 = [f"{newest_date_str} Total Conversions", f"{newest_date_str} Total reject",
                      f"{second_newest_date_str} Total Conversions", f"{second_newest_date_str} Total reject"]
    table4[int_cols_table4] = table4[int_cols_table4].astype(int)
    table4 = table4.sort_values('Affiliate').reset_index(drop=True)


    
    if progress_bar and status_text:
        progress_bar.progress(90)
        status_text.text("💾 准备下载文件...")
    
    # 返回所有结果
    results = {
        'table1': table1,
        'table2': table2,
        'table3': table3,
        'table4': table4,
        'newest_date_str': newest_date_str,
        'newest_date_file_str': newest_date_file_str,
        'stats': {
            '高差异Offer数量': len(high_diff_offers),
            '旧预算Offer数量': len(old_budget_offers),
            '新预算Offer数量': len(all_offers - old_budget_offers)
        }
    }
    
    if progress_bar and status_text:
        progress_bar.progress(100)
        status_text.text("🎉 分析完成！")
    
    return results

# ==================== 文件下载功能 ====================
def get_excel_download_link(results):
    """生成Excel文件下载链接"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        results['table1'].to_excel(writer, sheet_name='表格一_三级广告主日报表', index=False)
        results['table2'].to_excel(writer, sheet_name='表格二_高差异Offer ID详情', index=False)
        results['table3'].to_excel(writer, sheet_name='表格三_二级广告主综合报表', index=False)
        results['table4'].to_excel(writer, sheet_name='表格四_Affiliate综合报表', index=False)
    
    output.seek(0)
    b64 = base64.b64encode(output.read()).decode()
    filename = f"{results['newest_date_file_str']}日报分析结论.xlsx"
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">📥 下载完整分析报告</a>'
    return href

# ==================== Streamlit主界面 ====================
def main():
    st.markdown('<div class="main-header">📊网盟日报分析</div>', unsafe_allow_html=True)
    
    # 侧边栏
    with st.sidebar:
        st.header("📋 使用说明")
        st.markdown("""
        **无需安装任何软件，直接在网页中使用！**
        
        ### 使用步骤：
        1. 上传Excel数据文件
        2. 系统自动分析Offer数据  
        3. 查看分析结果并下载报告
        
        ### 支持功能：
        - ✅ 自动识别最新两天日期
        - ✅ 高差异Offer智能分析
        - ✅ Affiliate维度精准分析
        - ✅ 新旧预算自动判断
        - ✅ 一键下载完整报告
        """)
        
        st.header("⚙️ 分析规则")
        st.info("""
        - 高差异筛选：流水差绝对值≥10美金
        - Affiliate分析：收入变化≥5美金
        - 预算判断：过去6天收入>0=旧预算，否则新预算
        """)
        
        st.header("📊 文件要求")
        st.success("""
        确保Excel包含以下工作表：
        - 1--all data
        - 3--匹配广告主  
        - 4--reject事件
        - 2-reject规则
        """)
    
    # 主内容区 - 文件上传
    st.markdown("### 📤 第一步：上传Excel文件")
    
    uploaded_file = st.file_uploader(
        "选择Excel文件（支持.xlsx格式）",
        type=['xlsx'],
        help="请上传包含Offer数据的完整Excel文件"
    )
    
    if uploaded_file is not None:
        # 显示文件信息
        file_details = {
            "文件名": uploaded_file.name,
            "文件类型": uploaded_file.type,
            "文件大小": f"{uploaded_file.size / 1024:.2f} KB"
        }
        
        col1, col2 = st.columns([2, 1])
        with col1:
            st.json(file_details)
        
        # 数据预览
        with st.expander("📖 数据预览（前5行）", expanded=False):
            try:
                df_preview = pd.read_excel(uploaded_file, sheet_name='1--all data')
                st.dataframe(df_preview.head(), use_container_width=True)
                st.success(f"✅ 数据格式正确，共 {len(df_preview)} 行记录")
            except Exception as e:
                st.error(f"❌ 数据预览失败：{str(e)}")
        
        # 开始分析按钮
        if st.button("🚀 开始分析数据", type="primary", use_container_width=True):
            # 创建进度条
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # 处理数据
            with st.spinner("数据分析中，请稍候..."):
                try:
                    results = process_daily_report_web(uploaded_file, progress_bar, status_text)
                    
                    # 显示分析结果摘要
                    st.markdown("### 📈 分析结果摘要")
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("高差异Offer数量", results['stats']['高差异Offer数量'])
                    with col2:
                        st.metric("旧预算Offer", results['stats']['旧预算Offer数量'])
                    with col3:
                        st.metric("新预算Offer", results['stats']['新预算Offer数量'])
                    
                    # 结果显示标签页
                    tab1, tab2, tab3, tab4 = st.tabs([
                        "📊 二级广告主报表", 
                        "✅ 高差异Offer详情", 
                        "👥 二级广告主报表", 
                        "🔍 Affiliate报表"
                    ])
                    
                    with tab1:
                        st.dataframe(results['table1'], use_container_width=True)
                    
                    with tab2:
                        st.dataframe(results['table2'], use_container_width=True)
                    
                    with tab3:
                        st.dataframe(results['table3'], use_container_width=True)
                    
                    with tab4:
                        st.dataframe(results['table4'], use_container_width=True)
                    
                    # 下载功能
                    st.markdown("### 📥 下载分析报告")
                    st.markdown(get_excel_download_link(results), unsafe_allow_html=True)
                    
                    st.success("🎉 分析完成！点击上方链接下载完整报告")
                    
                except Exception as e:
                    st.error(f"❌ 分析过程中出现错误：{str(e)}")
                    st.code(str(e))
    
    else:
        # 欢迎界面
        st.markdown("""
        <div class="upload-area">
            <h3>🌐 欢迎使用Offer数据分析系统</h3>
            <p>请上传Excel文件开始分析，系统将自动处理并生成完整分析报告</p>
        </div>
        """, unsafe_allow_html=True)
        
        # 功能说明
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            ### ✨ 核心功能
            - **智能日期识别**：自动识别最新两天数据
            - **高差异分析**：精准筛选重要变化Offer
            - **预算类型判断**：自动区分新旧预算
            - **Affiliate分析**：详细分析每个流量方贡献
            """)
        
        with col2:
            st.markdown("""
            ### 📋 输出内容
            - 表格一：流水总结
            - 表格二：高差异Offer ID详情  
            - 表格三：广告主综合报表
            - 表格四：流量综合报表
            - 完整Excel报告一键下载
            """)

if __name__ == "__main__":
    main()
