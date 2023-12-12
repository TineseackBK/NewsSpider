import xlwt
import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog
import difflib
import jieba
from jieba import analyse
from collections import Counter
import logging
from PIL import Image
import numpy as np
from wordcloud import WordCloud, ImageColorGenerator

jieba.setLogLevel(logging.INFO)

punctuation = r"""!"#$%&'()*+,-./:;<=>?@[\]^_`{|}~“”？，！【】（）、。：；’‘……￥·"""

# 导入文件并清洗有空值的行
def import_file(file_path):
    file = pd.read_excel(file_path, dtype=object).dropna()
    return file

# 导入时筛选年份
def import_file_filt(file_path, year):
    file = pd.read_excel(file_path, dtype=object).dropna()
    
    tmp_list = file['时间']
    year_list = []
    for i in tmp_list:
        year_list.append(i[0:4])
    file['时间'] = year_list

    file_filt = file[file['时间']==year]
    return file_filt

# 统计网页来源
def ana_source(file):
    # 把 Series 转成 list
    source_list = file['来源'].tolist()

    # 统计每个 key 出现的次数
    source_frequency = {}
    for source in source_list:
        source_frequency[source] = source_frequency.get(source, 0) + 1
    source_list = sorted(source_frequency.items(), key=lambda d: d[1], reverse=True)
    source_frequency = {}
    for source in source_list:
        source_frequency[source[0]] = source[1]

    # 把少于三次的来源合并,称作"其他"
    source_count = 0
    source_delete = []
    for source_key, source_value in source_frequency.items():
        if source_value <= 2:
            source_count += 1
            source_delete.append(source_key)

    for source in source_delete:
        del source_frequency[source]

    if source_count > 0:
        source_frequency['其他'] = source_count

    # 创建 plt 图
    x = source_frequency.values()
    y = range(len(source_frequency))
    plt.rcParams['font.sans-serif'] = ['SimHei']
    source_graph = plt.barh(y, x)
    plt.xlabel('频数')
    plt.ylabel('来源')
    plt.yticks(y, source_frequency.keys())
    plt.title('网页来源频数')
    plt.bar_label(source_graph)
    plt.show()

# 统计网页更新时间
def ana_time(file):
    # 把 Series 转成 list
    stime_list = file['时间'].tolist()

    # 统计每个年份出现的次数
    stime_frequency = {}
    for stime in stime_list:
        stime_frequency[stime[0:4]] = stime_frequency.get(stime[0:4], 0) + 1

    # 按照年份排序
    stime_list = sorted(stime_frequency.items(), key=lambda d: d[0][0:4], reverse=False)
    stime_frequency = {}
    for stime in stime_list:
        stime_frequency[stime[0]] = stime[1]

    # 创建 plt 图
    x = stime_frequency.keys()
    y = stime_frequency.values()
    plt.rcParams['font.sans-serif'] = ['SimHei']
    stime_graph = plt.plot(x, y)
    plt.xlabel('年份')
    plt.ylabel('频数')
    plt.title('网页年份频数')
    plt.show()

# 统计网页总数量
def sum(file):
    print(f'[ 网页总数量: {len(file)} ]')

# 统计相似网页
def sum_sim(file):
    web_list = file['标题']
    # 防止重复
    simi_rep = []
    for web in web_list:
        simi = {}
        for web_o in web_list:
            if web_o == web:
                continue
            simi[web_o] = similar(web, web_o)
        for key, value in simi.items():
            if value >= 0.7 and value not in simi_rep:
                simi_rep.append(value)
                print(f'[{web}] 和 [{key}] 的相似度: {value*100}%')
    # 相似网页总数
    print(f'\n[ 相似程度达到70%及以上的网页总数: {len(simi_rep)*2}, 占总数的{len(simi_rep)*2/len(file)*100}% ]')

# 字符串相似度
def similar(str1, str2):
    # 注意不要判断标点符号
    return difflib.SequenceMatcher(lambda x: x in punctuation, str1, str2).quick_ratio()

# 提取关键词
def keyword(file, filt_option):
    # jieba 分词,用自带的函数进行 TF-IDF 算法,筛选出最关键的关键词
    # 先分词,然后把分词之后的内容整合到一起,统一进行 TF-IDF 算法

    body_list = file['正文'].tolist()
    tmp_list = jieba.lcut(str(body_list))
    seg_list = []
    for tmp in tmp_list:
        tmp = ''.join(tmp.split())
        if tmp != '' and tmp != '\n' and tmp != '\n\n':
            seg_list.append(tmp)

    # 关键名词
    keywords_all = analyse.extract_tags(
        sentence = str(seg_list),
        topK = 20,
        allowPOS = ['n'],
        withWeight = False,
        withFlag = False,
    )
    keywords_all_w = analyse.extract_tags(
        sentence = str(seg_list),
        topK = 50,
        allowPOS = ['n'],
        withWeight = True,
        withFlag = False,
    )
    print(f'[ 关键名词(TF-IDF): {keywords_all} ]')

    # 关键人名
    keywords_people = analyse.extract_tags(
        sentence = str(seg_list),
        topK = 20,
        allowPOS = ['nr'],
        withWeight = False,
        withFlag = False,
    )
    keywords_people_w = analyse.extract_tags(
        sentence = str(seg_list),
        topK = 50,
        allowPOS = ['nr'],
        withWeight = True,
        withFlag = False,
    )
    print(f'[ 关键人名(TF-IDF): {keywords_people} ]')

    # 关键地名
    keywords_site = analyse.extract_tags(
        sentence = str(seg_list),
        topK = 20,
        allowPOS = ['ns'],
        withWeight = False,
        withFlag = False,
    )
    keywords_site_w = analyse.extract_tags(
        sentence = str(seg_list),
        topK = 50,
        allowPOS = ['ns'],
        withWeight = True,
        withFlag = False,
    )
    print(f'[ 关键地名(TF-IDF): {keywords_site} ]')

    # 其他专名
    keywords_others = analyse.extract_tags(
        sentence = str(seg_list),
        topK = 20,
        allowPOS = ['nz'],
        withWeight = False,
        withFlag = False,
    )
    keywords_others_w = analyse.extract_tags(
        sentence = str(seg_list),
        topK = 50,
        allowPOS = ['nz'],
        withWeight = True,
        withFlag = False,
    )
    print(f'[ 其他专名(TF-IDF): {keywords_others} ]')

    # 生成词云
    wordcloud(keywords_all_w, filt_option, wc_type=0)
    wordcloud(keywords_people_w, filt_option, wc_type=1)
    wordcloud(keywords_site_w, filt_option, wc_type=2)
    wordcloud(keywords_others_w, filt_option, wc_type=3)

    # 关键词写入文件
    if filt_option == '-1':
        with open('keywords.txt', 'w') as kwf:
            kwf.write(f'[ 关键名词(TF-IDF): {keywords_all} ]\n')
            kwf.write(f'[ 关键人名(TF-IDF): {keywords_people} ]\n')
            kwf.write(f'[ 关键地名(TF-IDF): {keywords_site} ]\n')
            kwf.write(f'[ 其他专名(TF-IDF): {keywords_others} ]\n')
    else:
        with open(f'keywords_{filt_option}.txt', 'w') as kwf:
            kwf.write(f'[ {filt_option} 年关键名词(TF-IDF): {keywords_all} ]\n')
            kwf.write(f'[ {filt_option} 年关键人名(TF-IDF): {keywords_people} ]\n')
            kwf.write(f'[ {filt_option} 年关键地名(TF-IDF): {keywords_site} ]\n')
            kwf.write(f'[ {filt_option} 年其他专名(TF-IDF): {keywords_others} ]\n')

# 词云
def wordcloud(keywords_w, filt_option, wc_type):
    keywords = {}
    for word in keywords_w:
        keywords[word[0]] = word[1]

    wc = WordCloud(
        font_path='simhei.ttf',
        background_color='white',
    )
    wc.generate_from_frequencies(keywords)
    plt.imshow(wc)
    plt.axis('off')
    # plt.show()

    file_name = ''
    if wc_type == 0:
        file_name = '词云_关键名词'
    if wc_type == 1:
        file_name = '词云_关键人名'
    if wc_type == 2:
        file_name = '词云_关键地名'
    if wc_type == 3:
        file_name = '词云_其他专名'

    if filt_option == '-1':
        pass
    else:
        file_name = file_name + f'_{filt_option}'

    file_name = file_name + '.png'

    wc.to_file(file_name)
    print(f'[ 词云已保存为 \"{file_name}\" ]')

def main():
    # 弹窗选择文件并获取路径
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()

    # 操作选择
    while True:
        filt_option = input("请选择是否只统计某一年的内容\n若是, 请输入 [2014 - 2023] 中某一年的年份\n若否, 请输入 -1\n")
        if filt_option in ['2014', '2015', '2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023']:
            filt_year = filt_option
            file = import_file_filt(file_path, filt_year)
            print(f"[ 数据现已仅局限于{filt_year}年, 请继续操作 ]")
            break
        elif filt_option == '-1':
            # 正常导入文件
            file = import_file(file_path)
            break
        else:
            print("\n[ !请输入正确的选项! ]\n")

    while True:
        try:
            option = input("请选择操作: \n[ 1 - 统计网页总数量 ]\n[ 2 - 统计相似网页 ]\n[ 3 - 统计网页来源 ]\n[ 4 - 统计网页时间 ]\n[ 5 - 提取关键词并生成词云 ]\n")

            if option == '1':
                # 统计网页总数量
                sum(file)
            elif option == '2':
                # 统计相似网页
                sum_sim(file)
            elif option == '3':
                # 统计网页来源
                ana_source(file)
            elif option == '4':
                # 统计网页时间
                ana_time(file)
            elif option == '5':
                # 提取关键词
                keyword(file, filt_option)
            else:
                print("\n[ !请输入正确的选项! ]\n")
        except EOFError:
            break

if __name__ == "__main__":
    main()