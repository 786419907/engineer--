# coding:utf-8
'''
@author:                    Kai
LoginHNNU:                  登录接口
DownloadPersonalIcon：      爬头像
DownloadPersonExam:         爬个人成绩
DownloadPersonInformation： 爬个人信息
'''
from bs4 import BeautifulSoup
import requests
import xlrd
import time
import os
import re

# 开启全局session
session = requests.session()
# 伪造请求头
headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.115 Safari/537.36'}
def LoginHNNU(account,password):
    # 构造请求信息,这里的'B1'不难发现是代表登录的标识
    data = {'xh' : account,
            'sfzh' : password,
            'B1' : '%CC%E1%BD%BB'
            }
    # post请求
    resp = session.post('http://211.70.176.123/wap/index.asp',data = data,headers = headers)
    # 请求成功
    if resp.status_code == 200:
        html_content = resp.content.decode('gb2312')
    # 拼接地址
    url1 = 'http://211.70.176.123/dbsdb/tp.asp?xh=' + str(account)
    url2 = 'http://211.70.176.123/wap/grxx.asp'
    name = DownloadPersonInformation(url2)
    DownloadPersonalIcon(url1,name)
    DownloadPersonExam(name)
def DownloadPersonalIcon(url,name):
    # 创建文件夹
    if not os.path.exists('E:/icon_folder'):
        os.makedirs('E:/icon_folder')
    # 这里必须要用同一个session来发请求，教务系统必须验证session后才能下载图片，之前遇到坑，天真的以为获取到链接就可以下载了
    icon = session.get(url)
    # E盘的icon_folder文件夹里用流写头像,这里就很喜欢python了，以前java里读写I/O至少要十来行，还要用到缓冲Buffer、关闭流之类的，手写好痛苦，现在python用2行代码就搞定
    with open('E:/icon_folder/'+ name +'.png','wb') as f:
        f.write(icon.content)
        
def DownloadPersonInformation(url):
    # 美丽汤解析
    html_content = session.get(url,headers=headers).content.decode('gb2312')
    soup = BeautifulSoup(html_content,'html.parser',from_encoding='utf-8')
    name_item = soup.find('font',attrs={'color' : 'red'})
    name = name_item.get_text()
    items_all = soup.find('table',id='table1').findAll('td',attrs={'height':'22'})
    # 存入字典
    info_dict = {'姓名' : name}
    info_dict['学号']     = items_all[2].get_text()
    info_dict['专业']     = items_all[3].get_text()
    info_dict['院系']     = items_all[6].get_text()
    info_dict['班级']     = items_all[7].get_text()
    info_dict['民族']     = items_all[10].get_text()
    info_dict['籍贯']     = items_all[11].get_text()
    info_dict['出生日期']  = items_all[14].get_text()
    info_dict['政治面貌']  = items_all[15].get_text()
    info_dict['身份证号码'] = items_all[18].get_text()
    info_dict['高考考生号'] = items_all[19].get_text()
    with open('E:/info.txt','a+',encoding = 'utf-8',errors='ignore') as f:
        for key,value in info_dict.items():
            f.write(key + ':' + value + '  ')
        f.write('\n')
    return name

def DownloadPersonExam(name):
    url = 'http://211.70.176.123/wap/cjcx.asp'
    # xn,xqss,这2个字段很好猜，这个的crumb经过网上查证是反爬虫的随机串，看来学校还是有反爬虫的措施的
    # 应对办法：先来一次普通的get请求解析得到crumb，再用crumb进行post
    #这里的'Referer'是重点啊，之前爬了好几次都是空数据，以为data里少了什么参数，对比了请求头后发现服务器都对referer做了监控，无referer或假referer都被认定为非法访问。
    headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.115 Safari/537.36','Referer' : 'http://211.70.176.123/wap/cjcx.asp'}
    html_content = session.get(url,headers = headers).content.decode('gb2312')
    soup = BeautifulSoup(html_content,'html.parser',from_encoding='utf-8')
    crumb_item = soup.find('input',attrs = {'name' : "crumb"})
    str_item = str(crumb_item)
    str_split = str_item.split('"')
    crumb = str_split[len(str_split)-2]
    # 构造请求信息
    data = {
            'xn' : '2016-2017',
            'xqss' : '1',
            'crumb' : crumb,
            'B1' : '%CC%E1%BD%BB'
            }
    html_content = session.post(url,data = data,headers = headers).content.decode('gb2312')
    # 汤美丽解析
    soup = BeautifulSoup(html_content,'html.parser',from_encoding='utf-8')
    # 获得考试科目
    scores_name_list = []
    scores_name = soup.findAll('a',href = "#")
    for score_name in scores_name:
        scores_name_list.append(score_name.get_text())
    # 正则获得分数
    scores_score_list = []
    scores_score = soup.findAll('a',href = re.compile(r'^kcpm2.asp'))
    for score_score in scores_score:
        scores_score_list.append(score_score.get_text())
    with open('E:/exam.txt','a+',encoding = 'utf-8',errors='ignore') as f:
        f.write(name+'  :  ')
        for index,value in enumerate(scores_name_list):
            f.write(value + ':' + scores_score_list[index] + '    ')
        f.write('\n')

# 测试用例
if __name__ == '__main__':
    time_start = time.time()
    data = xlrd.open_workbook('info.xls')
    table = data.sheet_by_index(0)
    nrows = table.nrows
    list_num = []
    list_card = []
    for i in range(1,nrows):
        list_num.append(table.cell(i,4).value)
    for i in range(1,nrows):
        list_card.append(table.cell(i,7).value)
    for i in range(nrows):
        try:
            LoginHNNU(list_num[i], list_card[i])
        except:
            print('爬取失败')
    time_end = time.time()
    print(str((time_end - time_start) / 60))
