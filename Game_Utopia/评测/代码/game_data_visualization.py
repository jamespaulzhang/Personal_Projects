import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
from matplotlib import rcParams
import numpy as np
import re

# 设置中文字体
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']  # 使用Arial Unicode MS
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

def extract_rating_percentage(rating_str):
    """从评价字符串中提取百分比数值"""
    if pd.isna(rating_str) or rating_str == '无评价':
        return None
    
    # 使用正则表达式提取百分比数字
    match = re.search(r'(\d+\.?\d*)%', str(rating_str))
    if match:
        return float(match.group(1))
    return None

def load_and_preprocess_data(filepath):
    """加载并预处理数据"""
    df = pd.read_excel(filepath, engine='openpyxl')
    
    # 数据预处理
    df = df.dropna(how='all')  # 去除空白行
    
    # 处理评价数据 - 提取百分比数值
    df['评价百分比'] = df['商店评价'].apply(extract_rating_percentage)
    
    # 填充可能缺失的值
    df['评价百分比'] = df['评价百分比'].fillna(0)
    
    return df

def plot_publisher_distribution(df):
    """绘制发行商分布图"""
    plt.figure(figsize=(12, 8))
    publisher_counts = df['游戏发行商'].value_counts().head(15)
    
    # 使用渐变色
    colors = sns.color_palette("viridis", len(publisher_counts))
    
    ax = sns.barplot(x=publisher_counts.values, y=publisher_counts.index, palette=colors)
    plt.title('Top 15 游戏发行商 (按游戏数量)', fontsize=16)
    plt.xlabel('游戏数量', fontsize=12)
    plt.ylabel('发行商', fontsize=12)
    
    # 在柱子上添加数值
    for p in ax.patches:
        width = p.get_width()
        ax.text(width + 1, p.get_y() + p.get_height()/2., 
                f'{int(width)}', 
                ha='left', va='center', fontsize=10)
    
    plt.savefig('1_publisher_distribution.png', dpi=300, bbox_inches='tight')
    plt.show()

def plot_game_type_distribution(df):
    """绘制游戏类型分布图"""
    # Split the game types and count occurrences
    game_types = df['游戏类型'].dropna().str.split(',').explode().str.strip()
    type_counts = game_types.value_counts().head(15)

    plt.figure(figsize=(12, 8))

    # Create a bar plot for the top 15 game types
    ax = sns.barplot(x=type_counts.values, y=type_counts.index, palette="viridis")

    plt.title('Top 15 游戏类型 (按游戏数量)', fontsize=16)
    plt.xlabel('游戏数量', fontsize=12)
    plt.ylabel('游戏类型', fontsize=12)

    # Add value labels on the bars
    for p in ax.patches:
        width = p.get_width()
        ax.text(width + 1, p.get_y() + p.get_height()/2.,
                f'{int(width)}',
                ha='left', va='center', fontsize=10)

    plt.savefig('2_game_type_distribution.png', dpi=300, bbox_inches='tight')
    plt.show()

def plot_rating_distribution(df):
    """绘制评分分布图"""
    plt.figure(figsize=(12, 6))
    
    # 使用核密度估计
    ax = sns.histplot(df['评价百分比'], bins=20, kde=True, color='skyblue', 
                     edgecolor='white', linewidth=0.5)
    
    # 添加中位数和平均线
    median_rating = df['评价百分比'].median()
    mean_rating = df['评价百分比'].mean()
    
    ax.axvline(median_rating, color='red', linestyle='--', linewidth=2, 
               label=f'中位数: {median_rating:.1f}%')
    ax.axvline(mean_rating, color='green', linestyle='-', linewidth=2, 
               label=f'平均数: {mean_rating:.1f}%')
    
    plt.title('商店评价分布 (百分比)', fontsize=16)
    plt.xlabel('好评百分比', fontsize=12)
    plt.ylabel('游戏数量', fontsize=12)
    plt.legend()
    plt.savefig('3_rating_distribution.png', dpi=300, bbox_inches='tight')
    plt.show()

def plot_interactive_rating_by_genre(df):
    """使用 Plotly 创建交互式图表按游戏类型分析评分"""
    # Split the game types and explode them into separate rows
    genres_df = df.copy()
    genres_df['游戏类型'] = genres_df['游戏类型'].str.split(',')
    genres_df = genres_df.explode('游戏类型')
    genres_df['游戏类型'] = genres_df['游戏类型'].str.strip()

    # Calculate the average rating for each game type
    genre_ratings = genres_df.groupby('游戏类型')['评价百分比'].mean().reset_index()
    genre_ratings = genre_ratings.sort_values('评价百分比', ascending=False)

    # Create an interactive bar plot
    fig = px.bar(genre_ratings, x='评价百分比', y='游戏类型',
                 title='各游戏类型的平均好评百分比',
                 labels={'评价百分比': '平均好评百分比', '游戏类型': '游戏类型'},
                 orientation='h')

    fig.update_layout(
        yaxis=dict(autorange="reversed"),  # Reverse the y-axis to have the highest rating on top
        height=800,  # Adjust height to fit more categories
        xaxis_title='平均好评百分比',
        yaxis_title='游戏类型'
    )

    fig.show()
    fig.write_html('interactive_rating_by_genre.html')

def plot_review_status(df):
    """绘制测评状态分布"""
    plt.figure(figsize=(10, 6))
    
    status_counts = df['测评状态'].value_counts()
    
    # 创建水平条形图
    ax = sns.barplot(x=status_counts.values, y=status_counts.index, 
                    palette=['#4CAF50', '#FFC107', '#F44336'],  # 绿色、黄色、红色
                    edgecolor='black', linewidth=0.5)
    
    plt.title('测评状态分布', fontsize=16)
    plt.xlabel('数量', fontsize=12)
    plt.ylabel('状态', fontsize=12)
    
    # 添加数值标签
    for p in ax.patches:
        width = p.get_width()
        ax.text(width + 1, p.get_y() + p.get_height()/2., 
                f'{int(width)}', 
                ha='left', va='center', fontsize=10)
    
    plt.savefig('5_review_status.png', dpi=300, bbox_inches='tight')
    plt.show()

def plot_reviewer_activity(df):
    """绘制测评人员活跃度"""
    # 收集所有测评人员
    reviewers = pd.concat([df['测评人员1'], df['测评人员2'], 
                         df['测评人员3'], df['测评人员4']]).dropna()
    
    plt.figure(figsize=(14, 8))
    reviewer_counts = reviewers.value_counts().head(20)
    
    # 创建条形图
    ax = sns.barplot(x=reviewer_counts.values, y=reviewer_counts.index, 
                    palette='magma', edgecolor='black', linewidth=0.5)
    
    plt.title('Top 20 最活跃测评人员', fontsize=16)
    plt.xlabel('参与测评次数', fontsize=12)
    plt.ylabel('测评人员', fontsize=12)
    
    # 添加数值标签
    for p in ax.patches:
        width = p.get_width()
        ax.text(width + 0.2, p.get_y() + p.get_height()/2., 
                f'{int(width)}', 
                ha='left', va='center', fontsize=10)
    
    plt.savefig('6_reviewer_activity.png', dpi=300, bbox_inches='tight')
    plt.show()

def create_interactive_charts(df):
    """创建交互式图表"""
    # 1. 发行商与平均评分的交互式散点图
    publisher_ratings = df.groupby('游戏发行商').agg(
        game_count=('评价百分比', 'size'),
        avg_rating=('评价百分比', 'mean')
    ).reset_index().sort_values('game_count', ascending=False).head(50)
    
    fig1 = px.scatter(
        publisher_ratings,
        x='avg_rating',
        y='game_count',
        color='avg_rating',
        size='game_count',
        hover_name='游戏发行商',
        title='发行商分析: 游戏数量 vs 平均好评百分比',
        labels={
            'avg_rating': '平均好评百分比',
            'game_count': '游戏数量'
        },
        color_continuous_scale=px.colors.sequential.Viridis
    )
    fig1.update_layout(
        hovermode='closest',
        xaxis_title='平均好评百分比',
        yaxis_title='游戏数量',
        coloraxis_colorbar=dict(title='平均好评百分比')
    )
    fig1.write_html('interactive_publisher_ratings.html')
    
    # 2. 游戏类型与发放类型的交互式热力图
    cross_tab = pd.crosstab(df['游戏类型'], df['游戏发放类型'])
    fig2 = px.imshow(
        cross_tab,
        labels=dict(x="发放类型", y="游戏类型", color="数量"),
        x=cross_tab.columns,
        y=cross_tab.index,
        title='游戏类型 vs 发放类型分布',
        color_continuous_scale='Blues'
    )
    fig2.update_xaxes(side="top")
    fig2.write_html('interactive_genre_distribution.html')
    
    return fig1, fig2

def generate_all_visualizations(filepath):
    """生成所有可视化图表"""
    # 加载数据
    df = load_and_preprocess_data(filepath)
    
    print("数据加载完成，开始生成可视化图表...")
    
    # 静态图表
    plot_publisher_distribution(df)
    plot_game_type_distribution(df)
    plot_rating_distribution(df)
    plot_interactive_rating_by_genre(df)
    plot_review_status(df)
    plot_reviewer_activity(df)
    
    # 交互式图表
    print("正在生成交互式图表...")
    fig1, fig2 = create_interactive_charts(df)
    
    print("所有可视化图表已生成并保存到当前目录!")

if __name__ == '__main__':
    # 替换为你的Excel文件路径
    excel_file = '评测汇总.xlsx'
    generate_all_visualizations(excel_file)
