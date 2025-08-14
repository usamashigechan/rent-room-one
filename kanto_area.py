import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import time

def scrape_suumo_cities_kansai():
    """
    SUUMOの市区町村検索ページから関西全域（大阪、京都、兵庫、奈良、滋賀、和歌山）の市区町村名と件数を取得（修正版）
    """
    # 関西全域のURLリスト
    base_urls = {
        "Osaka": "https://suumo.jp/chintai/osaka/city/",
        "Kyoto": "https://suumo.jp/chintai/kyoto/city/",
        "Hyogo": "https://suumo.jp/chintai/hyogo/city/",
        "Nara": "https://suumo.jp/chintai/nara/city/",
        "Shiga": "https://suumo.jp/chintai/shiga/city/",
        "Wakayama": "https://suumo.jp/chintai/wakayama/city/"
    }
    
    # 府県名の対応表
    prefecture_names = {
        "Osaka": "大阪府",
        "Kyoto": "京都府", 
        "Hyogo": "兵庫県",
        "Nara": "奈良県",
        "Shiga": "滋賀県",
        "Wakayama": "和歌山県"
    }
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    
    all_results = []

    for prefecture_code, URL in base_urls.items():
        prefecture_name = prefecture_names[prefecture_code]
        print(f"\n=== {prefecture_name} ({prefecture_code}) のデータ抽出開始 ===")
        try:
            response = requests.get(URL, headers=headers, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")
            
            print("=== デバッグ情報 ===")
            print(f"レスポンスステータス: {response.status_code}")
            print(f"ページタイトル: {soup.title.string if soup.title else 'タイトルなし'}")
            
            # 元のセレクタの確認（路線用から市区町村用に調整）
            print("\n=== 元のセレクタの結果 ===")
            cities_original = soup.select("div.searchitem-list > ul > li")
            print(f"div.searchitem-list > ul > li の数: {len(cities_original)}")
            
            if len(cities_original) == 0:
                print("元のセレクタでは要素が見つかりません。代替セレクタを試します。")
            
            # 代替セレクタを試す
            alternative_selectors = [
                "div.searchitem-list li",
                ".searchitem-list li", 
                "div.searchitem-list a",
                ".searchitem-list a",
                "li a",
                "ul li",
                "li",
                "div.searchitem-list",
                ".searchitem-list"
            ]
            
            for selector in alternative_selectors:
                elements = soup.select(selector)
                print(f"{selector}: {len(elements)}個の要素")
                
                if len(elements) > 0:
                    print(f"  最初の要素のクラス: {elements[0].get('class', 'なし')}")
                    print(f"  最初の要素の内容（最初の50文字）: {elements[0].text.strip()[:50]}...")
            
            # 修正版1: 汎用的なセレクタを使用
            print("\n=== 修正版1: 汎用的なセレクタ ===")
            cities_v1 = soup.select("li")
            city_data_v1 = []
            
            for city in cities_v1:
                a_tag = city.select_one("a")
                count_tag = city.select_one("span.searchitem-list-value")
                
                if a_tag and count_tag:
                    city_name = a_tag.text.strip()
                    count = count_tag.text.strip()
                    
                    # 市区町村らしいかチェック（市、区、町、村が含まれる）
                    if any(suffix in city_name for suffix in ['市', '区', '町', '村']):
                        city_data_v1.append({
                            'prefecture_code': prefecture_code,
                            'prefecture_name': prefecture_name,
                            'city_name': city_name,
                            'count': count,
                            'method': 'v1'
                        })
                        print(f"  {city_name}: {count}")
            
            # 修正版2: span.searchitem-list-valueから逆算
            print("\n=== 修正版2: span.searchitem-list-valueから逆算 ===")
            count_tags = soup.select("span.searchitem-list-value")
            city_data_v2 = []
            
            for count_tag in count_tags:
                count = count_tag.text.strip()
                parent = count_tag.parent
                a_tag = None
                
                # 親要素を最大3階層まで遡ってaタグを探す
                for _ in range(3):
                    if parent is None:
                        break
                    a_tag = parent.select_one("a")
                    if a_tag:
                        break
                    parent = parent.parent
                
                if a_tag:
                    city_name = a_tag.text.strip()
                    # 市区町村らしいかチェック
                    if any(suffix in city_name for suffix in ['市', '区', '町', '村']):
                        city_data_v2.append({
                            'prefecture_code': prefecture_code,
                            'prefecture_name': prefecture_name,
                            'city_name': city_name,
                            'count': count,
                            'method': 'v2'
                        })
                        print(f"  {city_name}: {count}")
            
            # 修正版3: 正規表現を使用してリンクと件数を抽出
            print("\n=== 修正版3: 正規表現でリンクと件数を抽出 ===")
            all_links = soup.select(f"a[href*='chintai/{prefecture_code.lower()}/city/']")
            city_data_v3 = []
            
            for link in all_links:
                href = link.get('href', '')
                text = link.text.strip()
                
                # 市区町村らしいかチェック
                if any(suffix in text for suffix in ['市', '区', '町', '村']):
                    parent = link.parent
                    count_element = None
                    
                    # 親要素で件数を探す
                    if parent:
                        count_element = parent.select_one("span.searchitem-list-value")
                        if not count_element:
                            # 兄弟要素も探す
                            for sibling in link.next_siblings:
                                if hasattr(sibling, 'select_one'):
                                    count_element = sibling.select_one("span.searchitem-list-value")
                                    if count_element:
                                        break
                    
                    if count_element:
                        count = count_element.text.strip()
                        city_data_v3.append({
                            'prefecture_code': prefecture_code,
                            'prefecture_name': prefecture_name,
                            'city_name': text,
                            'count': count,
                            'method': 'v3',
                            'href': href
                        })
                        print(f"  {text}: {count}")
            
            # 修正版4: 全てのspan.searchitem-list-valueの前後を詳細に調査
            print("\n=== 修正版4: 詳細調査 ===")
            count_spans = soup.select("span.searchitem-list-value")
            city_data_v4 = []
            
            for i, span in enumerate(count_spans):
                count = span.text.strip()
                print(f"\n件数スパン {i+1}: {count}")
                
                city_name = "不明"
                
                # 前の兄弟要素を探す
                current = span.previous_sibling
                while current:
                    if hasattr(current, 'find') and current.find('a'):
                        a_tag = current.find('a')
                        potential_name = a_tag.text.strip()
                        # 市区町村らしいかチェック
                        if any(suffix in potential_name for suffix in ['市', '区', '町', '村']):
                            city_name = potential_name
                            break
                    elif hasattr(current, 'text') and current.text.strip():
                        potential_name = current.text.strip()
                        if any(suffix in potential_name for suffix in ['市', '区', '町', '村']):
                            city_name = potential_name
                            break
                    current = current.previous_sibling
                
                # 見つからない場合は親要素を探す
                if city_name == "不明":
                    parent = span.parent
                    while parent and city_name == "不明":
                        a_tag = parent.find('a')
                        if a_tag:
                            potential_name = a_tag.text.strip()
                            if any(suffix in potential_name for suffix in ['市', '区', '町', '村']):
                                city_name = potential_name
                                break
                        parent = parent.parent
                
                if city_name != "不明":
                    city_data_v4.append({
                        'prefecture_code': prefecture_code,
                        'prefecture_name': prefecture_name,
                        'city_name': city_name,
                        'count': count,
                        'method': 'v4'
                    })
                    print(f"  → 市区町村名: {city_name}")
            
            # 修正版5: より広範囲なテキスト検索
            print("\n=== 修正版5: テキスト全体から市区町村を検索 ===")
            page_text = soup.get_text()
            city_data_v5 = []
            
            # 市区町村名の一般的なパターンを検索
            city_patterns = [
                r'([一-龯ァ-ヴ]+[市区町村])\s*[（(]?(\d+[,，]?\d*)[）)]?',
                r'([一-龯ァ-ヴ]+[市区町村])\s*：\s*(\d+[,，]?\d*)',
                r'([一-龯ァ-ヴ]+[市区町村])\s+(\d+[,，]?\d*)件?'
            ]
            
            for pattern in city_patterns:
                matches = re.findall(pattern, page_text)
                for city_name, count in matches:
                    # 数字のクリーニング
                    count_clean = re.sub(r'[,，]', '', count)
                    if count_clean.isdigit():
                        city_data_v5.append({
                            'prefecture_code': prefecture_code,
                            'prefecture_name': prefecture_name,
                            'city_name': city_name,
                            'count': count,
                            'method': 'v5'
                        })
                        print(f"  {city_name}: {count}")
            
            # 各府県の結果をまとめる
            prefecture_results = city_data_v1 + city_data_v2 + city_data_v3 + city_data_v4 + city_data_v5
            all_results.extend(prefecture_results)
            
            if prefecture_results:
                df = pd.DataFrame(prefecture_results)
                print(f"\n=== {prefecture_name} 全体結果 ===")
                print(f"総抽出件数: {len(prefecture_results)}")
                print("\n抽出データ:")
                print(df.to_string(index=False, max_rows=10))
                
                if len(prefecture_results) > 10:
                    print("（10件のみ表示）")
                
                # 重複を除去
                unique_cities = {}
                for item in prefecture_results:
                    key = (item['prefecture_name'], item['city_name'])
                    if key not in unique_cities or len(item['count']) > len(unique_cities[key]['count']):
                        unique_cities[key] = item
                
                unique_df = pd.DataFrame(list(unique_cities.values()))
                print(f"\n=== {prefecture_name} 最終結果（重複除去後） ===")
                print(unique_df[['prefecture_name', 'city_name', 'count']].to_string(index=False))
            
            else:
                print(f"{prefecture_name} のデータを抽出できませんでした。")
            
        except requests.RequestException as e:
            print(f"{prefecture_name} のリクエストエラー: {e}")
        except Exception as e:
            print(f"{prefecture_name} のその他のエラー: {e}")
        
        # サーバー負荷軽減のため1秒待機
        time.sleep(1)
    
    # 全府県の結果を統合
    if all_results:
        final_df = pd.DataFrame(all_results)
        print(f"\n=== 関西全域の最終結果 ===")
        print(f"総抽出件数: {len(all_results)}")
        
        # 重複を除去して最終結果を作成
        unique_cities = {}
        for item in all_results:
            key = (item['prefecture_name'], item['city_name'])
            if key not in unique_cities or len(item['count']) > len(unique_cities[key]['count']):
                unique_cities[key] = item
        
        final_df = pd.DataFrame(list(unique_cities.values()))
        print(f"\n=== 最終結果（重複除去後） ===")
        print(final_df[['prefecture_name', 'city_name', 'count']].to_string(index=False))
        
        return final_df
    else:
        print("関西全域のデータを抽出できませんでした。")
        return pd.DataFrame()

def create_sample_city_data_kansai():
    """
    実際のページにアクセスできない場合のサンプルデータ（関西全域の市区町村）
    """
    sample_cities = [
        {"prefecture_name": "大阪府", "city_name": "大阪市北区", "count": "18,560"},
        {"prefecture_name": "大阪府", "city_name": "大阪市中央区", "count": "16,230"},
        {"prefecture_name": "大阪府", "city_name": "大阪市西区", "count": "12,450"},
        {"prefecture_name": "大阪府", "city_name": "堺市堺区", "count": "8,940"},
        {"prefecture_name": "大阪府", "city_name": "豊中市", "count": "11,680"},
        {"prefecture_name": "京都府", "city_name": "京都市中京区", "count": "9,150"},
        {"prefecture_name": "京都府", "city_name": "京都市下京区", "count": "8,670"},
        {"prefecture_name": "京都府", "city_name": "京都市上京区", "count": "6,340"},
        {"prefecture_name": "兵庫県", "city_name": "神戸市中央区", "count": "12,780"},
        {"prefecture_name": "兵庫県", "city_name": "神戸市灘区", "count": "8,450"},
        {"prefecture_name": "兵庫県", "city_name": "姫路市", "count": "7,230"},
        {"prefecture_name": "奈良県", "city_name": "奈良市", "count": "5,560"},
        {"prefecture_name": "滋賀県", "city_name": "大津市", "count": "4,340"},
        {"prefecture_name": "和歌山県", "city_name": "和歌山市", "count": "3,450"},
    ]
    
    return pd.DataFrame(sample_cities)

# 実行
if __name__ == "__main__":
    print("SUUMO市区町村別物件数抽出（関西全域・修正版）")
    print("=" * 50)
    
    df = scrape_suumo_cities_kansai()
    
    if df.empty:
        print("\n実際のページからデータを取得できませんでした。")
        print("サンプルデータを表示します：")
        print("=" * 30)
        
        sample_df = create_sample_city_data_kansai()
        for _, row in sample_df.iterrows():
            print(f"{row['prefecture_name']} - {row['city_name']}: {row['count']}")
        
        print(f"\nサンプルデータ総数: {len(sample_df)}件")
    else:
        print(f"\n抽出完了: {len(df)}件の市区町村データ")
        
        # CSVファイルとして保存
        df.to_csv('suumo_cities_kansai_fixed.csv', index=False, encoding='utf-8-sig')
        print("CSVファイル 'suumo_cities_kansai_fixed.csv' として保存しました。")
        
        # 府県別のサマリーも表示
        print("\n=== 府県別サマリー ===")
        if 'prefecture_name' in df.columns:
            summary = df.groupby('prefecture_name').size().sort_values(ascending=False)
            for pref, count in summary.items():
                print(f"{pref}: {count}市区町村")