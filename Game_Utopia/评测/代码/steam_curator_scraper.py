import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import random
import re
import urllib3

# 禁用SSL警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def get_curator_games(curator_id, output_file='curator_games.csv'):
    """
    爬取指定Steam鉴赏家评测过的所有游戏名称
    
    参数:
        curator_id: 鉴赏家ID (例如45337284)
        output_file: 输出文件名
    """
    base_url = f"https://store.steampowered.com/curator/{curator_id}/"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
        "Accept-Language": "zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Referer": "https://store.steampowered.com/"
    }
    
    # 添加cookies以绕过可能的年龄验证
    cookies = {
        'birthtime': '0',
        'mature_content': '1',
        'lastagecheckage': '1-January-1980'
    }
    
    games = []
    page = 1
    
    print(f"开始爬取鉴赏家 {curator_id} 的游戏评测...")
    
    while True:
        # 构建分页URL
        url = f"{base_url}?page={page}"
        
        try:
            print(f"正在处理第 {page} 页...")
            response = requests.get(url, headers=headers, cookies=cookies, timeout=30, verify=False)
            response.raise_for_status()
            
            # 检查是否被重定向到年龄验证页面
            if "agecheck" in response.url:
                print("检测到年龄验证，尝试绕过...")
                session = requests.Session()
                session.cookies.update(cookies)
                age_check_url = response.url
                form_data = {
                    'ageDay': '1',
                    'ageMonth': 'January',
                    'ageYear': '1980',
                    'snr': '1_agecheck_agecheck__age-gate'
                }
                response = session.post(age_check_url, data=form_data, headers=headers, timeout=30, verify=False)
                response.raise_for_status()
            
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # 新的选择器 - 尝试多种可能的元素选择方式
            game_blocks = soup.select('div.curator_review, div.recommendation')
            
            # 如果没有找到游戏块，尝试备用选择器
            if not game_blocks:
                game_blocks = soup.find_all('div', {'class': lambda x: x and 'review' in x.lower()})
            
            # 如果还是没有找到，说明可能到达最后一页
            if not game_blocks:
                print("已到达最后一页，爬取完成！")
                break
            
            # 提取每个游戏的信息
            for block in game_blocks:
                try:
                    # 尝试多种方式获取游戏名称
                    title_elem = block.select_one('div.curator_review_title a, a.title')
                    if not title_elem:
                        continue
                    
                    game_name = title_elem.text.strip()
                    game_url = title_elem.get('href', '')
                    
                    # 获取评测日期
                    date_elem = block.select_one('div.curator_review_date, div.posted_date')
                    review_date = date_elem.text.strip() if date_elem else "未知日期"
                    
                    # 提取游戏ID
                    app_id = None
                    match = re.search(r'/app/(\d+)', game_url)
                    if match:
                        app_id = match.group(1)
                    
                    games.append({
                        '游戏名称': game_name,
                        '游戏ID': app_id,
                        '游戏链接': game_url,
                        '评测日期': review_date
                    })
                except Exception as e:
                    print(f"处理游戏块时出错: {str(e)}")
                    continue
            
            # 检查是否还有下一页
            next_page = soup.select_one('a.pagebtn[href*="page="]:contains(">")')
            if not next_page:
                print("已到达最后一页，爬取完成！")
                break
            
            # 随机延迟，避免触发反爬
            time.sleep(random.uniform(2, 5))
            page += 1
            
        except Exception as e:
            print(f"处理第 {page} 页时出错: {str(e)}")
            print("可能已到达最后一页或页面结构发生变化。")
            break
    
    # 保存结果到CSV文件
    if games:
        df = pd.DataFrame(games)
        df.to_csv(output_file, index=False, encoding='utf-8-sig')
        print(f"\n成功爬取 {len(games)} 个游戏评测！")
        print(f"结果已保存到: {output_file}")
        
        # 打印前5条记录预览
        print("\n前5条记录预览:")
        print(df.head().to_string(index=False))
    else:
        print("没有找到任何游戏评测。")

if __name__ == "__main__":
    # 游戏乌托邦鉴赏家ID
    curator_id = 45337284
    get_curator_games(curator_id)
