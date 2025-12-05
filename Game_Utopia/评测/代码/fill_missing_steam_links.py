import pandas as pd
import re
from urllib.parse import urlparse

def extract_app_id_from_review_url(url):
    """
    从个人评价链接中提取appid
    示例输入: https://steamcommunity.com/profiles/76561199472684338/recommended/3190340/
    返回: 3190340
    """
    if pd.isna(url) or not isinstance(url, str):
        return None
    
    # 尝试多种匹配模式以确保提取成功
    patterns = [
        r'/recommended/(\d+)',  # 标准推荐链接
        r'/reviews/(\d+)',      # 可能的其他格式
        r'app/(\d+)',           # 直接包含appid的情况
        r'/(\d+)/?$'            # 纯数字结尾
    ]
    
    for pattern in patterns:
        match = re.search(pattern, url)
        if match:
            return match.group(1)
    
    return None

def generate_store_url(appid):
    """根据appid生成Steam商店链接"""
    if appid:
        return f"https://store.steampowered.com/app/{appid}/"
    return None

def fill_missing_store_links(input_file, output_file):
    """
    主处理函数:
    1. 读取Excel文件
    2. 从"个人评价链接1"提取appid
    3. 为缺失的"游戏商店链接"生成并填入正确的链接
    """
    # 读取Excel文件
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"无法读取Excel文件: {e}")
        return False
    
    # 确保列存在
    required_columns = ['游戏商店链接', '个人评价链接1']
    for col in required_columns:
        if col not in df.columns:
            print(f"错误: 列 '{col}' 不存在")
            return False
    
    # 统计初始缺失值
    initial_missing = df['游戏商店链接'].isna().sum()
    print(f"初始缺失的游戏商店链接数量: {initial_missing}")
    
    # 处理每一行
    filled_count = 0
    for index, row in df.iterrows():
        # 只有当游戏商店链接为空时才处理
        if pd.isna(row['游戏商店链接']) or row['游戏商店链接'].strip() == '':
            review_url = row['个人评价链接1']
            if pd.isna(review_url) or review_url.strip() == '':
                continue
                
            # 从评价链接提取appid
            appid = extract_app_id_from_review_url(review_url)
            if appid:
                # 生成商店链接并填入
                store_url = generate_store_url(appid)
                if store_url:
                    df.at[index, '游戏商店链接'] = store_url
                    filled_count += 1
                    print(f"已填充行 {index + 1}: {store_url}")
    
    # 保存结果
    try:
        df.to_excel(output_file, index=False)
        print(f"\n处理完成! 成功填充 {filled_count} 个缺失的游戏商店链接")
        print(f"处理后仍缺失的游戏商店链接数量: {df['游戏商店链接'].isna().sum()}")
        print(f"结果已保存到: {output_file}")
        return True
    except Exception as e:
        print(f"保存文件时出错: {e}")
        return False

if __name__ == "__main__":
    # 配置输入输出文件路径
    input_excel = "评测汇总.xlsx"  # 替换为你的输入文件路径
    output_excel = "评测汇总_已填充商店链接.xlsx"  # 输出文件路径
    
    print("开始处理缺失的游戏商店链接...")
    success = fill_missing_store_links(input_excel, output_excel)
    
    if success:
        print("\n操作成功完成!")
    else:
        print("\n处理过程中出现错误，请检查输出信息。")
