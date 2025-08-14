import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import time

def scrape_suumo_routes_fixed():
    """
    SUUMOの路線検索ページから関東全域（東京、神奈川、埼玉、千葉、群馬、茨城、栃木）の路線名と件数を取得（修正版）
    """
    # 関東全域のURLリスト
    base_urls = {
        "Tokyo": "https://suumo.jp/chintai/tokyo/ensen/",
        "Kanagawa": "https://suumo.jp/chintai/kanagawa/ensen/",
        "Saitama": "https://suumo.jp/chintai/saitama/ensen/",
        "Chiba": "https://suumo.jp/chintai/chiba/ensen/",
        "Gunma": "https://suumo.jp/chintai/gumma/ensen/",
        "Ibaraki": "https://suumo.jp/chintai/ibaraki/ensen/",
        "Tochigi": "https://suumo.jp/chintai/tochigi/ensen/"
    }
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    
    all_results = []

    for prefecture, URL in base_urls.items():
        print(f"\n=== {prefecture} のデータ抽出開始 ===")
        try:
            response = requests.get(URL, headers=headers, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")
            
            print("=== デバッグ情報 ===")
            print(f"レスポンスステータス: {response.status_code}")
            print(f"ページタイトル: {soup.title.string if soup.title else 'タイトルなし'}")
            
            # 元のセレクタの確認
            print("\n=== 元のセレクタの結果 ===")
            routes_original = soup.select("div.searchitem-list > ul > li")
            print(f"div.searchitem-list > ul > li の数: {len(routes_original)}")
            
            if len(routes_original) == 0:
                print("元のセレクタでは要素が見つかりません。代替セレクタを試します。")
            
            # 代替セレクタを試す
            alternative_selectors = [
                "div.searchitem-list li",
                ".searchitem-list li",
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
            routes_v1 = soup.select("li")
            route_data_v1 = []
            
            for route in routes_v1:
                a_tag = route.select_one("a")
                count_tag = route.select_one("span.searchitem-list-value")
                
                if a_tag and count_tag:
                    route_name = a_tag.text.strip()
                    count = count_tag.text.strip()
                    route_data_v1.append({
                        'prefecture': prefecture,
                        'route_name': route_name,
                        'count': count,
                        'method': 'v1'
                    })
                    print(f"{route_name}: {count}")
            
            # 修正版2: span.searchitem-list-valueから逆算
            print("\n=== 修正版2: span.searchitem-list-valueから逆算 ===")
            count_tags = soup.select("span.searchitem-list-value")
            route_data_v2 = []
            
            for count_tag in count_tags:
                count = count_tag.text.strip()
                parent = count_tag.parent
                a_tag = None
                
                for _ in range(3):
                    if parent is None:
                        break
                    a_tag = parent.select_one("a")
                    if a_tag:
                        break
                    parent = parent.parent
                
                if a_tag:
                    route_name = a_tag.text.strip()
                    route_data_v2.append({
                        'prefecture': prefecture,
                        'route_name': route_name,
                        'count': count,
                        'method': 'v2'
                    })
                    print(f"{route_name}: {count}")
            
            # 修正版3: 正規表現を使用してリンクと件数を抽出
            print("\n=== 修正版3: 正規表現でリンクと件数を抽出 ===")
            all_links = soup.select(f"a[href*='chintai/{prefecture.lower()}/ensen/']")
            route_data_v3 = []
            
            for link in all_links:
                href = link.get('href', '')
                text = link.text.strip()
                
                parent = link.parent
                count_element = None
                
                if parent:
                    count_element = parent.select_one("span.searchitem-list-value")
                    if not count_element:
                        for sibling in link.next_siblings:
                            if hasattr(sibling, 'select_one'):
                                count_element = sibling.select_one("span.searchitem-list-value")
                                if count_element:
                                    break
                
                if count_element:
                    count = count_element.text.strip()
                    route_data_v3.append({
                        'prefecture': prefecture,
                        'route_name': text,
                        'count': count,
                        'method': 'v3',
                        'href': href
                    })
                    print(f"{text}: {count}")
            
            # 修正版4: 全てのspan.searchitem-list-valueの前後を詳細に調査
            print("\n=== 修正版4: 詳細調査 ===")
            count_spans = soup.select("span.searchitem-list-value")
            route_data_v4 = []
            
            for i, span in enumerate(count_spans):
                count = span.text.strip()
                print(f"\n件数スパン {i+1}: {count}")
                
                prev_sibling = span.previous_sibling
                route_name = "不明"
                
                current = span.previous_sibling
                while current:
                    if hasattr(current, 'find') and current.find('a'):
                        a_tag = current.find('a')
                        route_name = a_tag.text.strip()
                        break
                    elif hasattr(current, 'text') and current.text.strip():
                        route_name = current.text.strip()
                        break
                    current = current.previous_sibling
                
                if route_name == "不明":
                    parent = span.parent
                    while parent and route_name == "不明":
                        a_tag = parent.find('a')
                        if a_tag:
                            route_name = a_tag.text.strip()
                            break
                        parent = parent.parent
                
                route_data_v4.append({
                    'prefecture': prefecture,
                    'route_name': route_name,
                    'count': count,
                    'method': 'v4'
                })
                print(f"  → 路線名: {route_name}")
            
            # 各都道府県の結果をまとめる
            prefecture_results = route_data_v1 + route_data_v2 + route_data_v3 + route_data_v4
            all_results.extend(prefecture_results)
            
            if prefecture_results:
                df = pd.DataFrame(prefecture_results)
                print(f"\n=== {prefecture} 全体結果 ===")
                print(f"総抽出件数: {len(prefecture_results)}")
                print("\n抽出データ:")
                print(df.to_string(index=False))
                
                # 重複を除去
                unique_routes = {}
                for item in prefecture_results:
                    key = (item['prefecture'], item['route_name'])
                    if key not in unique_routes or len(item['count']) > len(unique_routes[key]['count']):
                        unique_routes[key] = item
                
                unique_df = pd.DataFrame(list(unique_routes.values()))
                print(f"\n=== {prefecture} 最終結果（重複除去後） ===")
                print(unique_df[['prefecture', 'route_name', 'count']].to_string(index=False))
            
            else:
                print(f"{prefecture} のデータを抽出できませんでした。")
            
        except requests.RequestException as e:
            print(f"{prefecture} のリクエストエラー: {e}")
        except Exception as e:
            print(f"{prefecture} のその他のエラー: {e}")
        
        # サーバー負荷軽減のため1秒待機
        time.sleep(1)
    
    # 全都道府県の結果を統合
    if all_results:
        final_df = pd.DataFrame(all_results)
        print(f"\n=== 関東全域の最終結果 ===")
        print(f"総抽出件数: {len(all_results)}")
        
        # 重複を除去して最終結果を作成
        unique_routes = {}
        for item in all_results:
            key = (item['prefecture'], item['route_name'])
            if key not in unique_routes or len(item['count']) > len(unique_routes[key]['count']):
                unique_routes[key] = item
        
        final_df = pd.DataFrame(list(unique_routes.values()))
        print(f"\n=== 最終結果（重複除去後） ===")
        print(final_df[['prefecture', 'route_name', 'count']].to_string(index=False))
        
        return final_df
    else:
        print("関東全域のデータを抽出できませんでした。")
        return pd.DataFrame()

def create_sample_data():
    """
    実際のページにアクセスできない場合のサンプルデータ（関東全域）
    """
    sample_routes = [
        {"prefecture": "Tokyo", "route_name": "JR山手線", "count": "45,230"},
        {"prefecture": "Tokyo", "route_name": "JR中央線", "count": "38,940"},
        {"prefecture": "Tokyo", "route_name": "東京メトロ丸ノ内線", "count": "32,150"},
        {"prefecture": "Kanagawa", "route_name": "JR東海道本線", "count": "29,870"},
        {"prefecture": "Kanagawa", "route_name": "小田急小田原線", "count": "25,430"},
        {"prefecture": "Saitama", "route_name": "JR京浜東北線", "count": "22,150"},
        {"prefecture": "Saitama", "route_name": "JR埼京線", "count": "19,670"},
        {"prefecture": "Chiba", "route_name": "JR総武本線", "count": "18,340"},
        {"prefecture": "Chiba", "route_name": "京成本線", "count": "16,780"},
        {"prefecture": "Gunma", "route_name": "JR高崎線", "count": "12,450"},
        {"prefecture": "Ibaraki", "route_name": "JR常磐線", "count": "15,230"},
        {"prefecture": "Tochigi", "route_name": "JR東北本線", "count": "13,560"},
    ]
    
    return pd.DataFrame(sample_routes)

# 実行
if __name__ == "__main__":
    print("SUUMO路線検索データ抽出（関東全域・修正版）")
    print("=" * 50)
    
    df = scrape_suumo_routes_fixed()
    
    if df.empty:
        print("\n実際のページからデータを取得できませんでした。")
        print("サンプルデータを表示します：")
        print("=" * 30)
        
        sample_df = create_sample_data()
        for _, row in sample_df.iterrows():
            print(f"{row['prefecture']} - {row['route_name']}: {row['count']}")
        
        print(f"\nサンプルデータ総数: {len(sample_df)}件")
    else:
        print(f"\n抽出完了: {len(df)}件の路線データ")
        
        # CSVファイルとして保存
        df.to_csv('suumo_routes_kanto_fixed.csv', index=False, encoding='utf-8-sig')
        print("CSVファイル 'suumo_routes_kanto_fixed.csv' として保存しました。")