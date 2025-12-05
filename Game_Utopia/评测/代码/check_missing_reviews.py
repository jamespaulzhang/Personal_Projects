import pandas as pd
import numpy as np

def check_missing_reviews(input_file, output_file):
    """
    检查拖稿情况的脚本:
    1. 读取Excel文件
    2. 检查每个游戏的每个测评人员是否有对应的评价链接
    3. 生成拖稿报告
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(input_file)
        print("文件读取成功!")
    except Exception as e:
        print(f"无法读取Excel文件: {e}")
        return False
    
    # 检查必要列是否存在
    required_columns = [
        '测评游戏（商店全名）', '游戏发放类型', '游戏类型', '商店评价', 
        '游戏剩余情况', '游戏商店链接', '测评状态'
    ]
    
    # 动态检测测评人员列 (测评人员1-4)
    reviewer_columns = [col for col in df.columns if col.startswith('测评人员')]
    link_columns = [col for col in df.columns if col.startswith('个人评价链接')]
    
    if len(reviewer_columns) != len(link_columns):
        print("警告: 测评人员列和个人评价链接列数量不匹配!")
    
    # 创建拖稿记录列表
    missing_reviews = []
    
    # 遍历每一行(每个游戏)
    for index, row in df.iterrows():
        game_name = row['测评游戏（商店全名）']
        
        # 检查每个测评人员
        for i in range(len(reviewer_columns)):
            reviewer_col = f'测评人员{i+1}'
            link_col = f'个人评价链接{i+1}'
            
            # 如果测评人员列有值但链接为空
            if pd.notna(row[reviewer_col]) and pd.isna(row[link_col]):
                missing_reviews.append({
                    '游戏名称': game_name,
                    '游戏商店链接': row['游戏商店链接'],
                    '拖稿人员': row[reviewer_col],
                    '对应链接列': link_col,
                    '行号': index + 2  # Excel行号从1开始，header占1行
                })
    
    # 如果没有拖稿情况
    if not missing_reviews:
        print("恭喜！没有发现拖稿情况。")
        return True
    
    # 创建拖稿报告DataFrame
    report_df = pd.DataFrame(missing_reviews)
    
    # 按拖稿人员分组统计，并收集游戏名称
    stats_df = report_df.groupby('拖稿人员').agg({
        '游戏名称': ['count', lambda x: ', '.join(x)],
        '行号': lambda x: ', '.join(map(str, x))
    }).reset_index()
    
    # 重命名列
    stats_df.columns = ['拖稿人员', '拖稿数量', '涉及游戏', '行号']
    
    # 保存结果
    try:
        with pd.ExcelWriter(output_file) as writer:
            report_df.to_excel(writer, sheet_name='详细拖稿记录', index=False)
            stats_df.to_excel(writer, sheet_name='拖稿统计', index=False)
        
        print(f"\n发现 {len(missing_reviews)} 条拖稿记录:")
        print(stats_df.to_string(index=False))
        print(f"\n详细报告已保存到: {output_file}")
        return True
    except Exception as e:
        print(f"保存文件时出错: {e}")
        return False

if __name__ == "__main__":
    # 配置输入输出文件路径
    input_excel = "评测汇总.xlsx"  # 替换为你的输入文件路径
    output_excel = "拖稿报告.xlsx"  # 输出文件路径
    
    print("开始检查拖稿情况...")
    success = check_missing_reviews(input_excel, output_excel)
    
    if success:
        print("\n检查完成!")
    else:
        print("\n处理过程中出现错误，请检查输出信息。")
