import requests
import xlwt
import xlrd
import json

filePath = 'D:\新冠病毒2.xlsx'

class weiboData(object):

    def __init__(self):
        self.f = xlwt.Workbook()  #创建工作薄
        self.sheet1 = self.f.add_sheet(u'新冠病毒', cell_overwrite_ok = True)   #命名table
        self.rowsTitle = [u'编号', u'创建时间', u'发文博主', u'是否为加V博主', u'是否为长文本', u'内容']     #创建标题
        for i in range(0, len(self.rowsTitle)):
            #最后一个参数设置样式
            self.sheet1.write(0, i, self.rowsTitle[i], self.set_style('Times new Roman', 220, True))
        #Excel保存位置
        self.f.save(filePath)

    #该函数设置字体样式
    def set_style(self, name, height, bold=False):
        style = xlwt.XFStyle()  # 初始化样式
        font = xlwt.Font()  # 为样式创建字体
        font.name = name
        font.bold = bold
        font.colour_index = 2
        font.height = height
        style.font = font
        return style

    def getUrl(self):
        for page in range(100):
            url = 'https://m.weibo.cn/api/container/getIndex?containerid=100103type%3D1%26q%3D%E6%96%B0%E5%86%A0%E7%97%85%E6%AF%92&page_type=searchall&page={}'.format(page)
            self.spiderPage(url)

    def spiderPage(self, url):
        if url is None:
           return None

        try:
           data = xlrd.open_workbook(filePath)  #打开Excel文件
           table = data.sheets()[0] #通过索引顺序获取table，因为初始化时只创建了一个table，因此索引值为0
           rowCount = table.nrows  #获取行数，下次从这一行开始
           proxies = {  #使用代理IP，获取IP的方式在上一篇文章爬虫打卡4中有叙述
                'http':'http://110.73.1.47:8123'}
           user_agent="Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"
           headers = {'User-Agent': user_agent}


           m = 0
           response = requests.get(url, headers = headers, proxies = proxies)
           if response.content:
               all = json.loads(response.text)
               cards = all['data']['cards']
               for card in cards:
                   data = []
                   if card.get('mblog', None):
                       mblog = card['mblog']
                       time = mblog['created_at']
                       author = mblog['user']['screen_name']
                       isVerified = mblog['user']['verified']
                       if mblog['isLongText'] == True:
                           isLongText = True
                           text = mblog['longText']['longTextContent']
                       if mblog['isLongText'] == False:
                           isLongText = False
                           text = mblog['text']

                   #拼装成一个列表
                   data.append(rowCount + m)  # 为每条微博加序号
                   data.append(time)
                   data.append(author)
                   data.append(isVerified)
                   data.append(isLongText)
                   data.append(text)

                   for i in range(len(data)):
                       self.sheet1.write(rowCount+m,i,data[i]) #写入数据到execl中

                   m += 1   #记录行数增量
                   print(m)

           else:
               print("失败")

        except Exception as e:
               print ('出错',type(e),e)

        finally:
           self.f.save(filePath)

if __name__ == '__main__':
    weiboDemo = weiboData()
    weiboDemo.getUrl()


