import requests
import pandas as pd
import re
from time import sleep
import random
from bs4 import BeautifulSoup
import urllib3
import hashlib

# Disable all warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 百度翻译API配置
BAIDU_APP_ID = '20250530002369816'
BAIDU_SECRET_KEY = 'SuRrNE3RiYhTApBuH4it'

def clean_dataframe(df):
    """Ensure data types are correct"""
    str_columns = ['游戏类型', '商店评价', '游戏商店链接', '游戏发行商']
    for col in str_columns:
        if col in df.columns:
            df[col] = df[col].astype(str).replace('nan', '')
    return df

def extract_app_id(url):
    """Improved appid extraction function"""
    if pd.isna(url) or not isinstance(url, str):
        return None
    match = re.search(r'/app/(\d+)', url)
    return match.group(1) if match else None

def translate_with_baidu(text, from_lang='en', to_lang='zh'):
    """使用百度翻译API翻译文本"""
    if not text.strip():
        return text
    
    salt = random.randint(32768, 65536)
    sign = hashlib.md5((BAIDU_APP_ID + text + str(salt) + BAIDU_SECRET_KEY).encode()).hexdigest()
    
    url = "https://fanyi-api.baidu.com/api/trans/vip/translate"
    params = {
        'q': text,
        'from': from_lang,
        'to': to_lang,
        'appid': BAIDU_APP_ID,
        'salt': salt,
        'sign': sign
    }
    
    try:
        response = requests.get(url, params=params, timeout=10)
        result = response.json()
        if 'trans_result' in result:
            return ' '.join([item['dst'] for item in result['trans_result']])
        return text
    except Exception as e:
        print(f"翻译失败: {str(e)}")
        return text

def get_adult_cookie(age=21):
    """生成成人内容验证cookie"""
    return {
        'birthtime': '0',
        'mature_content': '1',
        'lastagecheckage': f'1-January-{2023-age}',
        'wants_mature_content': '1',
        'steamAgeVerified': '1'
    }

def get_publisher_info(session, app_id):
    """获取游戏发行商信息"""
    url = f"https://store.steampowered.com/api/appdetails?appids={app_id}&l=schinese"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"
    }
    cookies = get_adult_cookie()
    
    try:
        response = session.get(url, headers=headers, cookies=cookies, timeout=30)
        response.raise_for_status()
        data = response.json()
        
        if str(app_id) in data and data[str(app_id)]['success']:
            game_data = data[str(app_id)]['data']
            # 优先获取发行商(publishers)，如果没有则获取开发商(developers)
            publishers = game_data.get('publishers', [])
            if not publishers:
                publishers = game_data.get('developers', [])
            
            # 如果有多个发行商，用逗号分隔
            if publishers:
                return ', '.join(publishers)
        
        return "未知发行商"
    except Exception as e:
        print(f"❌ 获取发行商信息失败: {str(e)[:100]}")
        return "获取失败"

def scrape_steam_page(session, app_id):
    """从Steam商店页面爬取用户定义的标签 - 改进版，特别处理成人内容"""
    url = f"https://store.steampowered.com/app/{app_id}/"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
        "Accept-Language": "zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7",
    }
    
    # 设置成人内容cookie
    cookies = get_adult_cookie()
    
    try:
        # 第一次尝试：带成人cookie请求
        response = session.get(url, headers=headers, cookies=cookies, timeout=30)
        response.raise_for_status()
        
        # 检查是否被年龄验证拦截
        if "agecheck" in response.url:
            # 提交年龄验证表单
            soup = BeautifulSoup(response.text, 'html.parser')
            form = soup.find('form', {'id': 'agecheck_form'})
            if form:
                data = {
                    'ageDay': '1',
                    'ageMonth': 'January',
                    'ageYear': '1980',
                    'snr': '1_agecheck_agecheck__age-gate',
                    'sessionid': session.cookies.get('sessionid', '')
                }
                # 获取表单action URL
                action_url = form.get('action', url)
                if not action_url.startswith('http'):
                    action_url = f"https://store.steampowered.com{action_url}"
                
                # 提交年龄验证
                response = session.post(action_url, data=data, headers=headers, cookies=cookies, timeout=30)
                response.raise_for_status()
        
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 方法1：尝试获取流行标签
        tags = []
        tags_section = soup.select('div.glance_tags.popular_tags a')
        for tag in tags_section:
            tag_text = tag.text.strip()
            if tag_text and tag_text not in tags:
                tags.append(tag_text)
        
        # 方法2：如果标签太少，尝试从详情部分获取
        if len(tags) < 3:
            details_section = soup.find('div', class_='details_block')
            if details_section:
                for b in details_section.find_all('b'):
                    if '类型' in b.text or 'Genre' in b.text or '标签' in b.text or 'Tag' in b.text:
                        genre_text = b.next_sibling.strip()
                        if genre_text:
                            tags.extend([g.strip() for g in genre_text.split(',') if g.strip()])
        
        # 方法3：如果还是太少，尝试从游戏描述中提取关键词
        if len(tags) < 3:
            description = soup.find('div', class_='game_description_snippet')
            if description:
                desc_text = description.get_text().strip()
                if desc_text and len(desc_text.split()) > 2:
                    tags.append(desc_text)
        
        # 方法4：尝试从API获取更多标签
        if len(tags) < 3:
            api_url = f"https://store.steampowered.com/api/appdetails?appids={app_id}&l=schinese"
            api_response = session.get(api_url, headers=headers, cookies=cookies, timeout=30)
            if api_response.status_code == 200:
                api_data = api_response.json()
                if str(app_id) in api_data and api_data[str(app_id)]['success']:
                    data = api_data[str(app_id)]['data']
                    # 获取官方分类
                    if 'genres' in data:
                        tags.extend([g['description'] for g in data['genres']])
                    # 获取开发商/发行商信息
                    if 'developers' in data:
                        tags.extend(data['developers'])
                    if 'publishers' in data:
                        tags.extend(data['publishers'])
        
        return list(set(tags))[:15]  # 限制最多15个标签避免过多
    
    except Exception as e:
        print(f"❌ 爬取失败: {str(e)[:100]}")
        return []

def get_steam_game_info(session, app_id, api_key):
    """Fetch game info using Steam Web API and scraping - 改进版，特别处理成人内容"""
    user_tags = scrape_steam_page(session, app_id)
    details_url = f"https://store.steampowered.com/api/appdetails"
    details_params = {"appids": app_id, "l": "schinese"}  # 优先请求中文数据

    try:
        # 设置成人内容cookie
        cookies = get_adult_cookie()
        
        details_response = session.get(details_url, params=details_params, cookies=cookies, timeout=30)
        details_response.raise_for_status()
        details_data = details_response.json()

        game_data = details_data.get(str(app_id), {}).get("data", {}) if details_data.get(str(app_id), {}).get("success") else {}
        api_genres = [genre["description"] for genre in game_data.get("genres", [])]

        # 检查是否有中文数据
        has_chinese = 'schinese' in game_data.get('supported_languages', '').lower()
        
        # 合并所有标签
        all_tags = list(set(user_tags + api_genres))
        
        # 如果没有中文数据，则翻译英文标签
        if not has_chinese and all_tags:
            translated_tags = []
            for tag in all_tags:
                # 简单检查是否是中文字符
                if not any('\u4e00' <= char <= '\u9fff' for char in tag):
                    translated = translate_with_baidu(tag)
                    translated_tags.append(translated)
                else:
                    translated_tags.append(tag)
            all_tags = translated_tags
        
        genres_str = ', '.join(all_tags) if all_tags else '未知'

        # 获取评价信息
        reviews_url = f"https://store.steampowered.com/appreviews/{app_id}"
        reviews_params = {
            "json": 1, 
            "key": api_key, 
            "language": "schinese",
            "filter": "all",  # 获取所有评价
            "purchase_type": "all"  # 包括非Steam购买的评价
        }

        reviews_response = session.get(reviews_url, params=reviews_params, cookies=cookies, timeout=30)
        reviews_response.raise_for_status()
        reviews_data = reviews_response.json()

        review_summary = reviews_data.get("query_summary", {})
        total_positive = review_summary.get('total_positive', 0)
        total_negative = review_summary.get('total_negative', 0)

        total_reviews = total_positive + total_negative
        if total_reviews > 0:
            positive_rate = total_positive / total_reviews
            review_text = f"{positive_rate:.1%} 好评 ({total_positive} 好评, {total_negative} 差评)"
        else:
            review_text = "无评价"

        # 获取发行商信息
        publisher = get_publisher_info(session, app_id)

        return {
            '游戏类型': genres_str,
            '商店评价': review_text,
            '游戏商店链接': f"https://store.steampowered.com/app/{app_id}/",
            '游戏发行商': publisher
        }

    except Exception as e:
        print(f"❌ 请求失败: {str(e)[:100]}")
        return {
            '游戏类型': ', '.join(user_tags) if user_tags else '未知',
            '商店评价': "无评价",
            '游戏商店链接': f"https://store.steampowered.com/app/{app_id}/",
            '游戏发行商': "获取失败"
        }

# Your Steam Web API Key
api_key = "578A56730A0159A5AF01CEA6B9075902"

# Read and clean data
df = pd.read_excel("评测汇总.xlsx")
df = clean_dataframe(df)

# 确保有游戏发行商列
if '游戏发行商' not in df.columns:
    df['游戏发行商'] = ''

# Use a session to reuse TCP connections
with requests.Session() as session:
    for index, row in df.iterrows():
        need_update = (row['游戏类型'] in ['', '未知', 'nan'] or
                      row['商店评价'] in ['', '无评价', 'nan'] or
                      row['游戏发行商'] in ['', 'nan'])

        if not need_update:
            continue

        print(f"\n=== Processing row {index+1} ===")
        print(f"Game: {row['测评游戏（商店全名）']}")

        app_id = extract_app_id(row['游戏商店链接'])
        if not app_id:
            print("⏩ No valid appid, skipping")
            continue

        if info := get_steam_game_info(session, app_id, api_key):
            print(f"✅ Successfully retrieved - Type: {info['游戏类型']} | Reviews: {info['商店评价']} | Publisher: {info['游戏发行商']}")

            updates = {}
            if row['游戏类型'] in ['', '未知', 'nan']:
                updates['游戏类型'] = info['游戏类型']
            if row['商店评价'] in ['', '无评价', 'nan']:
                updates['商店评价'] = info['商店评价']
            if pd.isna(row['游戏商店链接']) or row['游戏商店链接'] == '':
                updates['游戏商店链接'] = info['游戏商店链接']
            if pd.isna(row['游戏发行商']) or row['游戏发行商'] in ['', 'nan']:
                updates['游戏发行商'] = info['游戏发行商']

            for col, value in updates.items():
                df.at[index, col] = value
        else:
            print("❌ Retrieval failed")

        # 增加延迟以避免触发反爬
        sleep_time = random.uniform(2, 5)
        sleep(sleep_time)

# Save results
output_path = "评测汇总_带发行商.xlsx"
df.to_excel(output_path, index=False)
print(f"\n🎉 Processing complete! Results saved to: {output_path}")
