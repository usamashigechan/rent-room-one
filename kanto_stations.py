import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import time
from datetime import datetime

def scrape_kanto_stations():
    """
    関東全域の都県、路線、駅ごとの物件数のみをスクレイピング
    """
    # 関東各県の路線ページURL
    prefectures = {
        "東京都": "https://suumo.jp/chintai/tokyo/ensen/",
        "神奈川県": "https://suumo.jp/chintai/kanagawa/ensen/",
        "埼玉県": "https://suumo.jp/chintai/saitama/ensen/",
        "千葉県": "https://suumo.jp/chintai/chiba/ensen/",
        "群馬県": "https://suumo.jp/chintai/gumma/ensen/",
        "茨城県": "https://suumo.jp/chintai/ibaraki/ensen/",
        "栃木県": "https://suumo.jp/chintai/tochigi/ensen/"
    }
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }
    
    all_stations = []
    
    print("関東全域の駅別物件数をスクレイピング開始")
    print("=" * 50)
    
    for pref_name, pref_url in prefectures.items():
        print(f"\n{pref_name} 処理中...")
        
        try:
            # 都道府県の路線ページを取得
            response = requests.get(pref_url, headers=headers, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")
            
            # 路線リンクを取得
            route_links = get_route_links(soup)
            print(f"  {len(route_links)}路線を発見")
            
            # 各路線の駅データを取得
            for i, (route_name, route_url) in enumerate(route_links, 1):
                print(f"  [{i:2d}] {route_name:30s}", end=" ")
                
                stations = get_stations_from_route(route_url, route_name, pref_name, headers)
                
                if stations:
                    all_stations.extend(stations)
                    print(f"→ {len(stations)}駅")
                else:
                    print("→ 0駅")
                
                time.sleep(1)  # サーバー負荷軽減
            
            print(f"  ✅ {pref_name} 完了")
            
        except Exception as e:
            print(f"  ❌ {pref_name} エラー: {e}")
        
        time.sleep(2)  # 都道府県間の間隔
    
    return all_stations

def get_route_links(soup):
    """路線ページから路線リンクを取得"""
    route_links = []
    
    # 複数の方法で路線リンクを探す
    selectors = [
        "li.searchitem a",
        "div.searchitem a",
        "a[href*='/ensen/']"
    ]
    
    for selector in selectors:
        links = soup.select(selector)
        
        for link in links:
            href = link.get('href', '')
            text = link.text.strip()
            
            # 有効な路線リンクかチェック
            if (href and '/ensen/' in href and 
                text and len(text) >= 3 and len(text) <= 30 and
                not any(word in text for word in ['路線', 'ページ', 'SUUMO', '検索', 'ヘルプ'])):
                
                # 完全URLを作成
                if href.startswith('/'):
                    href = f"https://suumo.jp{href}"
                
                route_links.append((text, href))
        
        if route_links:
            break
    
    # 重複除去
    return list(set(route_links))

def get_stations_from_route(route_url, route_name, pref_name, headers):
    """路線ページから駅と物件数を取得"""
    try:
        response = requests.get(route_url, headers=headers, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        
        stations = []
        
        # 方法1: 標準的なsearchitem構造
        items = soup.select("li.searchitem")
        for item in items:
            link = item.select_one("a")
            count_span = item.select_one("span.searchitem-list-value")
            
            if link and count_span:
                station_name = clean_station_name(link.text.strip())
                count = count_span.text.strip()
                
                if is_valid_station_name(station_name):
                    stations.append({
                        '都県名': pref_name,
                        '路線名': route_name,
                        '駅名': station_name,
                        '物件数': count
                    })
        
        # 方法2: 物件数から逆算（方法1で取得できない場合）
        if not stations:
            count_spans = soup.select("span.searchitem-list-value")
            for span in count_spans:
                count = span.text.strip()
                
                # 親要素を辿って駅名リンクを探す
                current = span.parent
                for _ in range(3):
                    if current is None:
                        break
                    
                    link = current.select_one("a")
                    if link and ('/eki/' in link.get('href', '') or 'eki' in link.get('href', '')):
                        station_name = clean_station_name(link.text.strip())
                        
                        if is_valid_station_name(station_name):
                            stations.append({
                                '都県名': pref_name,
                                '路線名': route_name,
                                '駅名': station_name,
                                '物件数': count
                            })
                        break
                    
                    current = current.parent
        
        return stations
        
    except Exception:
        return []

def clean_station_name(name):
    """駅名をクリーンアップ"""
    if not name:
        return ""
    
    # 不要な文字を除去
    name = re.sub(r'[\[\]（）\(\)]', '', name)
    name = re.sub(r'駅$', '', name)  # 末尾の「駅」を除去
    name = name.strip()
    
    return name

def is_valid_station_name(name):
    """有効な駅名かチェック"""
    if not name or len(name) < 2 or len(name) > 15:
        return False
    
    invalid_keywords = [
        '路線', 'ページ', '一覧', '検索', 'SUUMO', 'エリア',
        '詳細', '条件', 'もっと', 'ヘルプ', '地図'
    ]
    
    return not any(keyword in name for keyword in invalid_keywords)

def save_to_csv(stations_data):
    """CSVファイルに保存"""
    if not stations_data:
        print("保存するデータがありません")
        return
    
    # DataFrameに変換
    df = pd.DataFrame(stations_data)
    
    # 重複除去
    df = df.drop_duplicates(subset=['都県名', '路線名', '駅名'])
    
    # CSV保存
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f'関東駅別物件数_{timestamp}.csv'
    df.to_csv(filename, index=False, encoding='utf-8-sig')
    
    print(f"\n💾 CSV保存完了: {filename}")
    print(f"📊 データ数: {len(df)}駅")
    
    return df, filename

def show_results(df):
    """結果を表示"""
    print(f"\n📋 データサンプル (先頭10行):")
    print(df.head(10).to_string(index=False))
    
    print(f"\n🗾 都県別駅数:")
    for pref, count in df['都県名'].value_counts().items():
        print(f"   {pref}: {count}駅")
    
    # 物件数の多い駅トップ10
    try:
        df_copy = df.copy()
        df_copy['物件数_数値'] = df_copy['物件数'].str.replace(',', '').str.replace('件', '').astype(int, errors='ignore')
        top_stations = df_copy.nlargest(10, '物件数_数値')
        
        print(f"\n🏆 物件数トップ10駅:")
        for i, (_, row) in enumerate(top_stations.iterrows(), 1):
            print(f"   {i:2d}. {row['都県名']} {row['路線名']} {row['駅名']} ({row['物件数']})")
    except:
        pass

# メイン実行
if __name__ == "__main__":
    print("🚆 関東全域 駅別物件数スクレイピングツール")
    print("出力: 都県名, 路線名, 駅名, 物件数")
    
    # スクレイピング実行
    print("\n🚀 スクレイピング開始...")
    stations_data = scrape_kanto_stations()
    
    if stations_data:
        print(f"\n✅ スクレイピング完了: {len(stations_data)}駅のデータを取得")
        
        # CSV保存
        df, filename = save_to_csv(stations_data)
        
        # 結果表示
        show_results(df)
        
        print(f"\n🎉 処理完了! '{filename}' に保存されました")
    
    else:
        print("\n❌ データ取得に失敗しました")
        
        # サンプルデータを保存
        sample_data = [
            {'都県名': '東京都', '路線名': 'JR山手線', '駅名': '新宿', '物件数': '12,450'},
            {'都県名': '東京都', '路線名': 'JR山手線', '駅名': '渋谷', '物件数': '10,230'},
            {'都県名': '神奈川県', '路線名': '東急大井町線', '駅名': '二子新地', '物件数': '4,300'},
            {'都県名': '神奈川県', '路線名': 'JR東海道本線', '駅名': '横浜', '物件数': '15,230'}
        ]
        
        df = pd.DataFrame(sample_data)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f'関東駅別物件数_sample_{timestamp}.csv'
        df.to_csv(filename, index=False, encoding='utf-8-sig')
        
        print(f"サンプルデータを保存: {filename}")
        print(df.to_string(index=False))