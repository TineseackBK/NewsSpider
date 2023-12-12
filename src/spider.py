from bs4 import BeautifulSoup
# from selenium import webdriver
# from selenium.webdriver.chrome.options import Options
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
import requests
from requests.exceptions import HTTPError
import xlwt
import json
from jsonpath import jsonpath
import time
import re
import random
from fake_useragent import UserAgent
import datetime as dt 

# 创建 Excel 文件和表单
news_book = xlwt.Workbook(encoding='UTF-8')
news_sheet1 = news_book.add_sheet('sheet1', cell_overwrite_ok=True)

news_sheet1.write(0, 0, "标题")
news_sheet1.write(0, 1, "摘要")
news_sheet1.write(0, 2, "链接")
news_sheet1.write(0, 3, "来源")
news_sheet1.write(0, 4, "时间")
news_sheet1.write(0, 5, "正文")

news_sheet1.col(0).width = 6400
news_sheet1.col(1).width = 6400
news_sheet1.col(2).width = 6400
news_sheet1.col(3).width = 3200
news_sheet1.col(4).width = 3200
news_sheet1.col(5).width = 51200

line = 1

# 默认关键词
key = "芯片制造 芯片市场"

# 起始地址池
start_urls = ["https://search.cctv.com/search.php?qtext=", "https://i.news.qq.com/gw/pc_search/result", "https://api.thepaper.cn/search/web/news"]

def main():
    global key
    key = input("输入关键词：")
    start_page = input("输入起始页数：")
    max_page = input("输入终止页数：")
    url_index = 1
    for url in start_urls:
        if url_index == 1:
            url = url + key
            cctv_news(url, int(start_page), int(max_page))
        if url_index == 2:
            tencent_news(url, int(start_page), int(max_page))
        if url_index == 3:
            paper_news(url, int(start_page), int(max_page))
        url_index += 1
    # test()

# 央视新闻
def cctv_news(url, start_page, max_page):
    for page in range(start_page, max_page+1):
        time.sleep(random.uniform(2, 5))
        url_page = url + "&type=web&page=" + str(page)
        try:
            headers = {'User-agent': str(UserAgent().random)}
            response = requests.get(url_page, headers=headers)
            response.close()
            response.encoding = 'utf-8'
            response.raise_for_status
            html = response.text

            soup = BeautifulSoup(html, 'lxml')

            # 获取搜索到的每个文章的 li
            li_tags = soup.body.find_all('li', class_='image')

            for li in li_tags:
                # 获取文章标题
                article_title = li.find('div', class_='tright').find('h3', class_='tit')
                # 获取摘要
                article_summary = li.find('div', class_='tright').find('p', class_='bre')
                # 获取来源和时间
                article_source = li.find('div', class_='tright').find('div', class_='src-tim').find('span', class_='src')
                article_time = li.find('div', class_='tright').find('div', class_='src-tim').find('span', class_='tim')
                # 获取链接
                article_link = li.find('div', class_='tright').find('h3', class_='tit').find('span').get('lanmu1')
                
                # 获取正文
                # 如果正文获取不到就开摆，之后拿摘要分析就行
                article_body = cctv_news_body(article_link)

                global line

                try:
                    news_sheet1.write(line, 0, article_title.text.lstrip().rstrip())
                    news_sheet1.write(line, 1, article_summary.text.lstrip().rstrip())
                    news_sheet1.write(line, 2, article_link.lstrip().rstrip())
                    news_sheet1.write(line, 3, article_source.text[3:].lstrip().rstrip())
                    news_sheet1.write(line, 4, article_time.text[5:].lstrip().rstrip())
                    try:
                        news_sheet1.write(line, 5, article_body.lstrip().rstrip())
                    except:
                        continue
                except:
                    print("Error When Writing CCTV News")
                finally:
                    print(f'央视新闻第{page}页，总第{line}行已完成！')
                    line += 1
            
        except HTTPError as http_err:
            print(f"HTTP error occurred: {http_err}")
            continue
        except Exception as err:
            print(f"Other error occurred: Page {page} fucked up - {err}")
            continue

# 央视新闻正文
def cctv_news_body(link):
    time.sleep(random.uniform(2, 5))
    headers = {'User-agent': str(UserAgent().random)}
    response = requests.get(link, headers=headers)
    response.raise_for_status
    response.close()
    response.encoding = 'utf-8'
    html = response.text
    soup = BeautifulSoup(html, 'lxml')

    # 哪怕都是央视新闻，也有不同的网页结构，只能多试几种了
    # 如果这一种没成功，body_text 就是空的，就试试下一种

    body_text = ''

    try:
        body_text = soup.body.find('div', class_='content_area')
    except:
        print("Error 1")

    if body_text is None:
        try:
            body_text = soup.body.find('div', class_='cnt_bd')
            body_text.h1.decompose()
            body_text.h2.decompose()
            body_text.find('div', class_='function').decompose()
        except:
            print(link, "fucked up - Error 2")

    if body_text is None:
        try:
            body_text = soup.body.find('div', class_='cont')
        except:
            print(link, "fucked up - Error 3")
        
    if body_text is None:
        try:
            body_text = soup.body.find('div', class_='text_area')
        except:
            print(link, "fucked up - Error 4")
            return

    try:
        output_str = re.sub(r'\s+', ' ', body_text.text.lstrip().rstrip())
        return output_str
    except:
        return

# 腾讯新闻
def tencent_news(url, start_page, max_page):
    for page in range(start_page, max_page+1):
        # 腾讯新闻搜索页面是动态加载的，所以要用 Selenium 模拟浏览器行为才能获取资源
        # try:
        #     # 改变策略，提高加载速度
        #     chrome_options = Options()
        #     chrome_options.page_load_strategy = 'eager'
        #     driver = webdriver.Chrome(options=chrome_options)

        #     driver.get(url_page)

        #     # 显式等待直到元素可见
        #     element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "tcopyright")))

        #     driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        #     html = driver.page_source

        #     soup = BeautifulSoup(html, 'lxml')

        #     # 获取搜索到的每个文章的 li
        #     ul = driver.find_element(By.XPATH, '//*[@id="root"]/div/div[1]/div[1]/div[2]/ul')
        #     li_tags = ul.find_elements(By.XPATH, 'li')

        # Selenium 失败，尝试从 AJAX 入手
        try:
            time.sleep(random.uniform(2, 5))
            headers = {'User-agent': str(UserAgent().random)}
            form_data = {
                "page": page,
                "query": key,
                "is_pc": "1",
                "hippy_custom_version": "18",
                "search_type": "all",
                "search_count_limit": "10",
                "appver": "15.5_qqnews_7.1.80"
            }

            response = requests.post(url, data=form_data, headers=headers)
            response.raise_for_status
            response.close()
            response.encoding = 'utf-8'

            web_json = response.json()

            # 获取搜索到的文章列表 secList，在此基础上获取 newsList
            sec_list = jsonpath(web_json, '$..secList')
            news_list = jsonpath(sec_list[0], '$..newsList')

            for news in news_list:
                news_title = news[0]['longtitle']
                news_abstract = news[0]['abstract']
                news_url = news[0]['url']
                for i in news:
                    if not isinstance(i, dict):
                        continue
                    else:
                        news_source = i['chlname']
                        break
                news_time = news[0]['time']

                # 获取正文
                news_body = tencent_news_body(news_url)

                global line

                try:
                    news_sheet1.write(line, 0, news_title.lstrip().rstrip())
                    news_sheet1.write(line, 1, news_abstract.lstrip().rstrip())
                    news_sheet1.write(line, 2, news_url.lstrip().rstrip())
                    news_sheet1.write(line, 3, news_source.lstrip().rstrip())
                    news_sheet1.write(line, 4, news_time.lstrip().rstrip())
                    news_sheet1.write(line, 5, news_body.lstrip().rstrip())           
                except:
                    print("Error When Writing Tencent News")
                finally:
                    print(f'腾讯新闻第{page}页，总第{line}行已完成！')
                    line += 1

        except HTTPError as http_err:
            print(f"HTTP error occurred: {http_err}")
            continue
        except Exception as err:
            print(f"Other error occurred: {err}")
            continue

# 腾讯新闻正文
def tencent_news_body(link):
    time.sleep(random.uniform(2, 5))
    headers = {'User-agent': str(UserAgent().random)}
    response = requests.get(link, headers=headers)
    response.raise_for_status
    response.close()
    response.encoding = 'utf-8'
    html = response.text
    soup = BeautifulSoup(html, 'lxml')

    try:
        body_text = soup.body.find('div', id='ArticleContent')
    except:
        print("Body Error 1")
        return
    
    try:
        output_str = re.sub(r'\s+', ' ', body_text.text.lstrip().rstrip())
        return output_str
    except:
        print("Body Error 2")
        return

# 澎湃新闻
def paper_news(url, start_page, max_page):
    for page in range(start_page, max_page+1):
        try:
            time.sleep(random.uniform(2, 5))
            headers = {'User-agent': str(UserAgent().random)}
            payload = {
                "pageNum": page,
                "word": key,
                "orderType": "3",
                "pageSize": "10",
                "searchType": "1"
            }

            response = requests.post(url, json=payload, headers=headers)
            response.raise_for_status
            response.close()
            response.encoding = 'utf-8'

            web_json = response.json()

            # 获取搜索到的文章列表 secList，在此基础上获取 newsList
            web_data = jsonpath(web_json, '$..data')
            news_list = jsonpath(web_data, '$..list')

            for news in news_list[0]:
                # 使用 BeautifulSoup 去除爬取的文本里的 HTML 标签
                news_title = BeautifulSoup(news['name'], 'lxml').text
                news_abstract = BeautifulSoup(news['summary'], 'lxml').text
                news_cont = news['contId']
                news_url = f'https://www.thepaper.cn/newsDetail_forward_{news_cont}'
                news_source = news['nodeInfo']['name']
                news_time = news['pubTime']

                # 获取正文
                news_body = paper_news_body(news_url)
                
                global line

                try:
                    news_sheet1.write(line, 0, news_title.lstrip().rstrip())
                    news_sheet1.write(line, 1, news_abstract.lstrip().rstrip())
                    news_sheet1.write(line, 2, news_url.lstrip().rstrip())
                    news_sheet1.write(line, 3, news_source.lstrip().rstrip())
                    news_sheet1.write(line, 4, news_time.lstrip().rstrip())
                    news_sheet1.write(line, 5, news_body.lstrip().rstrip())
                except:
                    print("Error When Writing Paper News")
                finally:
                    print(f'澎湃新闻第{page}页，总第{line}行已完成！')
                    line += 1

        except HTTPError as http_err:
            print(f"HTTP error occurred: {http_err}")
            continue
        except Exception as err:
            print(f"Other error occurred: {err}")
            continue

# 澎湃新闻正文
def paper_news_body(link):
    time.sleep(random.uniform(2, 5))
    headers = {'User-agent': str(UserAgent().random)}
    response = requests.get(link, headers=headers)
    response.raise_for_status
    response.close()
    response.encoding = 'utf-8'
    html = response.text
    soup = BeautifulSoup(html, 'lxml')

    try:
        body_text = soup.body.find('div', class_='index_cententWrap__Jv8jK')
    except:
        print("Body Error 1")
        return
    
    try:
        output_str = re.sub(r'\s+', ' ', body_text.text.lstrip().rstrip())
        return output_str
    except:
        print("Body Error 2")
        return

def test():
    headers = {'User-agent': str(UserAgent().random)}
    url = 'https://news.cctv.com/2023/11/09/ARTIc4znsKeJGbuk1d1kc0Xe231109.shtml'
    response = requests.get(url, headers=headers)
    response.close()
    response.encoding = 'utf-8'
    html = response.text
    soup = BeautifulSoup(html, 'lxml')
    body_text = soup.body.find('div', class_='text_area')
    # body_text.h1.decompose()
    # body_text.h2.decompose()
    # body_text.find('div', class_='function').decompose()
    output_str = re.sub(r'\s+', ' ', body_text.text.lstrip().rstrip())
    print(output_str)

if __name__ == "__main__":
    main()
    # 获取当前时间
    time_now = dt.datetime.now().strftime('%F')
    filename = f'.\新闻原始数据_{time_now}_{time.time()}_line{line}.xls'
    news_book.save(filename)