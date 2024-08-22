import requests
from bs4 import BeautifulSoup
import pygsheets
from google.oauth2.service_account import Credentials
import urllib.parse

# 設定 Google Sheets API 認證
SERVICE_ACCOUNT_FILE = '$USER API.json'  # 憑證文件路徑

credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE)
client = pygsheets.authorize(service_account_file=SERVICE_ACCOUNT_FILE)
spreadsheet_id = '$Target sheet ID'
sh = client.open_by_key(spreadsheet_id)

# 年齡認證
def age_verification():
    session = requests.Session()
    age_check_url = "https://www.dmm.co.jp/age_check/=/declared=yes/?rurl=https%3A%2F%2Fwww.dmm.co.jp%2Fdigital%2Fvideoa%2F-%2Flist%2F%3Factress%3D8704%26view%3Dtext"
    session.get(age_check_url)
    return session

# 搜尋出演者ID
def search_actress_id(session, actress_name):
    search_url = f'https://actress.dmm.co.jp/-/search/=/searchstr={urllib.parse.quote(actress_name)}'
    response = session.get(search_url)
    response.encoding = 'utf-8'
    soup = BeautifulSoup(response.text, 'html.parser')
    
    actress_tag = soup.find('a', href=lambda x: x and '/-/detail/=/actress_id=' in x)
    if not actress_tag:
        return None
    actress_url = actress_tag['href']
    actress_id = actress_url.split('actress_id=')[-1].split('/')[0]
    return actress_id

# 獲取影片資料
def fetch_video_data(session, actress_id, existing_urls):
    video_data = []
    page = 1
    while True:
        list_url = f'https://www.dmm.co.jp/digital/videoa/-/list/?actress={actress_id}&view=text&page={page}'
        response = session.get(list_url)
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'html.parser')

        # 檢查是否有影片
        video_links = soup.find_all('a', href=lambda x: x and '/digital/videoa/-/detail/=/cid=' in x)
        if not video_links or response.url != list_url:
            break
        for link in video_links:
            video_url = link['href']
            if 'cid=' in video_url:
                video_page_url = f'https://www.dmm.co.jp{video_url}'
                print(f"處理網址： '{video_page_url}' 。")
                if video_page_url not in existing_urls:
                    video_page = session.get(video_page_url)
                    video_soup = BeautifulSoup(video_page.text, 'html.parser')
                    try:
                        title = video_soup.find('meta', property='og:title')['content']
                        code_td = video_soup.find('td',  string='品番：')
                        if code_td:
                            code = code_td.find_next_sibling('td').text.strip()
                        else:
                            code = 'Unknown'
                        release_date = video_soup.find('td', string='配信開始日：').find_next_sibling('td').text.strip()
                        sale_date = video_soup.find('td', string='商品発売日：').find_next_sibling('td').text.strip()
                        genres = video_soup.find('td', string='ジャンル：').find_next_sibling('td').find_all('a')
                        genres_text = [genre.text.strip() for genre in genres]
                        single_work = "単体作品" if "単体作品" in genres_text else ""
                        best_of = "ベスト・総集編" if "ベスト・総集編" in genres_text else ""
                        video_data.append({
                            'code': code,
                            'title': title,
                            'video_page_url': video_page_url,
                            'release_date': release_date,
                            'sale_date': sale_date,
                            'genres': genres_text,
                            'single_work': single_work,
                            'best_of': best_of
                        })
                    except Exception as e:
                        print(f"錯誤處理影片資料：{e}")
                        continue
                else:
                    print(f"{video_page_url} 已存在，跳過爬取。" )
        page += 1
    return video_data

# 將數字轉換為字母 (例如 1 -> A, 2 -> B, 27 -> AA)
def number_to_letter(n):
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result

# 根據E欄品番跟G欄標題是否存在值來判斷該筆資料完整
def check_row_completeness(worksheet, current_index):
    # 取得 E 欄和 G 欄的值
    e_value = worksheet.cell(f'E{current_index}').value
    g_value = worksheet.cell(f'G{current_index}').value

    # 檢查 E 欄和 F 欄是否都有值
    if e_value and g_value:
        return True  # 資料完整
    return False  # 資料不完整

# 計算工作表應該宣告多大
def calculate_worksheet_dimensions(video_data, current_row_count, current_column_count):
    total_videos = len(video_data)  # 計算總影片數
    print(f"影片:{total_videos}")
    max_genres_count = 0  # 初始化最大ジャンル數
    for video in video_data:
        genres_count = len(video['genres'])
        if genres_count > max_genres_count:
            max_genres_count = genres_count

    final_rows = current_row_count + total_videos

    if current_column_count > max_genres_count + 10:
        total_columns = current_column_count
    else:
        total_columns = 10 + max_genres_count  # 10 是固定欄位數，max_genres_count 是最大ジャンル數
    return final_rows, total_videos, total_columns

# 寫入 Google Sheets
def write_to_google_sheets(actress_name, video_data, current_row_count, current_column_count):
    # 計算作品數和欄位數
    final_rows, total_videos, total_columns = calculate_worksheet_dimensions(video_data, current_row_count, current_column_count)
    print(f"'{actress_name}' 要更新的影片總數為： {total_videos}, 欄位數: {total_columns}")
    
    if current_row_count == 0:
        final_rows = final_rows + 1
    
    final_cols = total_columns
    print(f"工作表大小應為： {final_rows}列，{final_cols}行。")

    # 檢查是否已存在該出演者的分頁
    #worksheet = sh.worksheet_by_title(actress_name)
    worksheet_names = [worksheet.title for worksheet in sh.worksheets()]    
    if actress_name in worksheet_names:
        worksheet = sh.worksheet_by_title(actress_name)
        # 更新工作表的列數和欄位數
        worksheet.resize(final_rows, final_cols)   
    else:
        worksheet = sh.add_worksheet(actress_name, final_rows, final_cols)
        worksheet.update_value('A1', 'PB路徑')
        worksheet.update_value('B1', '下載者')
        worksheet.update_value('C1', '備註')
        worksheet.update_value('D1', '通用番號')
        worksheet.update_value('E1', '品番')
        worksheet.update_value('F1', '片名')
        worksheet.update_value('G1', '影片網址')
        worksheet.update_value('H1', '配信開始日')
        worksheet.update_value('I1', '商品發售日')
        worksheet.update_value('J1', '単体作品')
        worksheet.update_value('K1', 'ベスト・総集編')

    # 凍結第一行
    worksheet.frozen_rows = 1
    
    # 取得已存在的品番
    existing_codes = worksheet.get_col(5, include_tailing_empty=False)[1:]  

    # 寫入新資料
    row_index = len(existing_codes) + 2 # 寫在最後一列，加上扣掉第一列凍結的。
    for data in video_data:
        if data['code'] not in existing_codes:
            worksheet.update_value(f'E{row_index}', data['code'])
            print(f"成功寫入品番：'{data['code']}''。")
            worksheet.update_value(f'F{row_index}', data["title"])
            print('成功寫入標題')
            worksheet.update_value(f'H{row_index}', data['release_date'])
            print('成功寫入配信日期')
            worksheet.update_value(f'I{row_index}', data['sale_date'])
            print('成功寫入發售日期')
            worksheet.update_value(f'J{row_index}', data['single_work'])
            print('成功寫入單體作品')
            worksheet.update_value(f'K{row_index}', data['best_of'])
            print('成功寫入總集編')            
    
            # 將影片網址放在該筆資料的最後寫入，如果影片網址為空，則表示此筆資料未寫入完整。
            worksheet.update_value(f'G{row_index}', data['video_page_url'])
            print('成功寫入影片網址')
            row_index += 1
        else:
            current_row_index = existing_codes.index(data['code']) + 2
            if not check_row_completeness(worksheet, current_row_index):
                print(f"第{current_row_index}列資料不完整，更新中。")
                worksheet.update_value(f'E{current_row_index}', data['code'])
                print(f"成功寫入品番：'{data['code']}''。")
                worksheet.update_value(f'F{current_row_index}', data["title"])
                print('成功寫入標題')
                worksheet.update_value(f'H{current_row_index}', data['release_date'])
                print('成功寫入配信日期')
                worksheet.update_value(f'I{current_row_index}', data['sale_date'])
                print('成功寫入發售日期')
                worksheet.update_value(f'J{current_row_index}', data['single_work'])
                print('成功寫入單體作品')
                worksheet.update_value(f'K{current_row_index}', data['best_of'])
                print('成功寫入總集編')
                # 將其他的ジャンル資料寫入J欄之後的欄位
                col_index = 12  # K欄開始
                for genre in data['genres']:
                    if genre not in ['単体作品', 'ベスト・総集編']:
                        worksheet.update_value((current_row_index, col_index), genre)
                        col_index += 1      
                worksheet.update_value(f'G{current_row_index}', data['video_page_url'])
                print('成功寫入影片網址') 
                print(f"第 {current_row_index} 列資料已完整。")

# 主程式
def main():

    # 建立會話，執行年齡認證
    session = age_verification()

    # 讀取"女優列表"分頁
    actress_list_sheet = sh.worksheet_by_title("女優列表")
    actress_names = actress_list_sheet.get_col(1, include_tailing_empty=False)[1:]  # 跳過標題列，從A2開始
    
    # 設定想排除的名稱，工作表女優列表內的A欄的值。
    exclude_actress_names = [""] 

    for i, actress_name in enumerate(actress_names, start=2):
        if actress_name not in exclude_actress_names:
            print(f"{actress_name}不存在於例外列表，開始處理。")
            actress_id_value = actress_list_sheet.cell(f'D{i}').value
            if not actress_id_value:
                print(f"女優ID: '{actress_id_value}' 不存在於Sheet中，前往DMM網站搜尋女優名： '{actress_name}' 。")
                actress_id = search_actress_id(session, actress_name)
                print(f"取得女優ID: '{actress_id}' 。")
            else:
                actress_id = actress_id_value
                print(f"取得女優ID: '{actress_id}' 。")

            worksheet_names = [worksheet.title for worksheet in sh.worksheets()]
            if actress_name in worksheet_names:
                actress_worksheet = sh.worksheet_by_title(actress_name)
                existing_urls = actress_worksheet.get_col(7, include_tailing_empty=False)[1:]  # 取得G欄影片網址
                # 取得目前女優工作表的列與欄數。
                current_row_count = actress_worksheet.rows
                current_column_count = actress_worksheet.cols
                print(f"目前工作表列數：{current_row_count}，目前工作表欄數：{current_column_count}")
            else:
                existing_urls = []

            if actress_id:
                # 更新女優列表中的出演者ID
                actress_list_sheet.update_value(f'D{i}', f'=HYPERLINK("https://www.dmm.co.jp/digital/videoa/-/list/?actress={actress_id}", "{actress_id}")')

                # 獲取影片資料
                video_data = fetch_video_data(session, actress_id, existing_urls)

                # 寫入資料到Google Sheets
                write_to_google_sheets(actress_name, video_data, current_row_count, current_column_count)

                # 將分頁女優列表中的出演者名稱建立連結到對應的分頁
                actress_name_tab = sh.worksheet_by_title(actress_name)
                actress_name_tab_gid = actress_name_tab.id
                actress_list_sheet.update_value(f'A{i}', f'=HYPERLINK("#gid={actress_name_tab_gid}", "{actress_name}")')
            else:
                print(f"查無此人：{actress_name}")
                continue
        else:
            print(f"{actress_name}存在於例外列表，跳過不處理。")
        
if __name__ == "__main__":
    main()