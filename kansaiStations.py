import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import time

def scrape_suumo_kansai_lines_stations():
    """
    SUUMOの関西全府県の路線ページから府県名、路線名（<a>タグ）、駅名（<span>タグ）、件数（span.searchitem-list-value、(X,XXX)形式を整数に変換）を取得
    """
    BASE_URL = "https://suumo.jp"
    PREFECTURES = [
        {"name": "大阪府", "code": "osaka"},
        {"name": "京都府", "code": "kyoto"},
        {"name": "兵庫県", "code": "hyogo"},
        {"name": "奈良県", "code": "nara"},
        {"name": "滋賀県", "code": "shiga"},
        {"name": "和歌山県", "code": "wakayama"}
    ]
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    
    all_results = []
    
    for prefecture in PREFECTURES:
        prefecture_name = prefecture["name"]
        prefecture_code = prefecture["code"]
        lines_url = f"https://suumo.jp/chintai/{prefecture_code}/ensen/"
        
        print(f"\n=== {prefecture_name} の路線一覧ページ ({lines_url}) ===")
        try:
            response = requests.get(lines_url, headers=headers, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")
            
            print(f"レスポンスステータス: {response.status_code}")
            print(f"ページタイトル: {soup.title.string if soup.title else 'タイトルなし'}")
            
            # 路線名の<a>タグを取得
            line_links = soup.select(f"a[href*='chintai/{prefecture_code}/en_']")
            print(f"路線リンクの数: {len(line_links)}")
            
            if not line_links:
                print(f"{prefecture_name}: 路線リンクが見つかりません。")
                continue
            
            for link in line_links:
                line_name = link.text.strip()
                line_url = BASE_URL + link.get('href', '')
                print(f"\n=== 路線: {line_name} ({line_url}) ===")
                
                # 各路線のページを取得
                try:
                    line_response = requests.get(line_url, headers=headers, timeout=10)
                    line_response.raise_for_status()
                    line_soup = BeautifulSoup(line_response.text, "html.parser")
                    
                    print(f"  レスポンスステータス: {line_response.status_code}")
                    print(f"  ページタイトル: {line_soup.title.string if line_soup.title else 'タイトルなし'}")
                    
                    # Method v1: 汎用的なセレクタを使用
                    stations = line_soup.select("li")
                    print(f"  li の数: {len(stations)}")
                    
                    if len(stations) == 0:
                        print(f"  {line_name}: セレクタで要素が見つかりません。")
                        continue
                    
                    for station in stations:
                        span_tag = station.select_one("span:not(.searchitem-list-value)")
                        count_tag = station.select_one("span.searchitem-list-value")
                        
                        if span_tag and count_tag:
                            station_name = span_tag.text.strip()
                            count_str = count_tag.text.strip()
                            count_clean = re.sub(r'[(),]', '', count_str)
                            try:
                                count = int(count_clean)
                                all_results.append({
                                    'prefecture_name': prefecture_name,
                                    'line_name': line_name,
                                    'station_name': station_name,
                                    'count': count
                                })
                                print(f"    {station_name}: {count}")
                            except ValueError:
                                print(f"    件数変換エラー: {station_name} の件数 '{count_str}' は数値に変換できません")
                    
                    # サーバー負荷軽減のため少し待機
                    time.sleep(1)
                    
                except requests.RequestException as e:
                    print(f"  {line_name} のリクエストエラー: {e}")
                    continue
                except Exception as e:
                    print(f"  {line_name} のその他のエラー: {e}")
                    continue
            
        except requests.RequestException as e:
            print(f"{prefecture_name} の路線一覧ページのリクエストエラー: {e}")
            continue
        except Exception as e:
            print(f"{prefecture_name} のその他のエラー: {e}")
            continue
    
    if all_results:
        df = pd.DataFrame(all_results)
        print(f"\n=== 抽出結果 ===")
        print(f"総抽出件数: {len(df)}")
        print("\n抽出データ:")
        print(df.to_string(index=False))
        
        # 府県名、路線名、駅名でソート
        final_df = df.sort_values(by=['prefecture_name', 'line_name', 'station_name'])
        print(f"\n=== 最終結果（関西全府県・全路線・全駅） ===")
        print(final_df[['prefecture_name', 'line_name', 'station_name', 'count']].to_string(index=False))
        
        return final_df
    else:
        print("データを抽出できませんでした。")
        return pd.DataFrame()

# 実行
if __name__ == "__main__":
    print("SUUMO関西全府県・全路線・全駅データ抽出（府県名、路線名：<a>タグ、駅名：<span>タグ、整数、(X,XXX)形式対応）")
    print("=" * 50)
    
    df = scrape_suumo_kansai_lines_stations()
    
    if df.empty:
        print("\n実際のページからデータを取得できませんでした。")
    else:
        print(f"\n抽出完了: {len(df)}件の駅データ")
        
        # CSVファイルとして保存
        df.to_csv('suumo_kansai_all_lines_stations.csv', index=False, encoding='utf-8-sig')
        print("CSVファイル 'suumo_kansai_all_lines_stations.csv' として保存しました。")