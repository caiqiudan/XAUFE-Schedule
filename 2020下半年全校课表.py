import requests
from parsel import Selector
import xlwt

class Schedule():
    def __init__(self,page):
        self.url = 'http://my.xaufe.edu.cn/ufeportal/xxcx/qxkbcx'
        # 设置请求头，需要cookie
        self.headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Content-Length': '74',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Cookie': 'JSESSIONID=D474*******5AF7; JSESSIONID=3B9D865B6C*******09B4BFCD', # 请用自己的cookie
        'Host': 'my.xaufe.edu.cn',
        'Origin': 'http://my.xaufe.edu.cn',
        'Pragma': 'no-cache',
        'Referer': 'http://my.xaufe.edu.cn/ufeportal/xxcx/qxkbcx',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.25 Safari/537.36 Core/1.70.3741.400 QQBrowser/10.5.3863.400 chrome-extension',}
        # 表单
        self.page = page


    def get_save_data(self):
        book = xlwt.Workbook() # encoding='utf-8'
        sheet = book.add_sheet('全校课表')
        # 添加标题行
        head = ['星期', '课程大节', '周次', '教室', '老师', '科目', '班级']
        for h in range(len(head)):
            sheet.write(0, h, head[h])
        n = 1
        for i in range(1,self.page+1): # 遍历指定页数
            print(f'爬取第{i}页')
            form = {'xnxq': '2020-2021上学期','pageNo': f'{i}'} # 页数，共200页
            response = requests.post(self.url, headers=self.headers, data=form)

            # xpath提取内容
            text = response.text
            selector = Selector(text=text)
            week = selector.xpath('//*[@id="sample_1"]/..//tr/td[2]/text()').getall() # 星期
            time = selector.xpath('//*[@id="sample_1"]/..//tr/td[3]/text()').getall() # 课程大节
            week_time = selector.xpath('//*[@id="sample_1"]/..//tr/td[4]/text()').getall() # 周次
            classroom = selector.xpath('//*[@id="sample_1"]/..//tr/td[5]/text()').getall() # 教室
            teacher = selector.xpath('//*[@id="sample_1"]/..//tr/td[6]/text()').getall() # 老师
            subject = selector.xpath('//*[@id="sample_1"]/..//tr/td[7]/text()').getall() # 科目
            classid = selector.xpath('//*[@id="sample_1"]/..//tr/td[8]/text()').getall() # 班级
            print(classid)
            try:
                for j in range(0,50):
                    data = [week[j],time[j],week_time[j],classroom[j],teacher[j],subject[j],classid[j]]

                    # 保存文件
                    m = 0
                    for dat in data:
                        sheet.write(n,m,dat)
                        m += 1
                    n += 1
            except:
                pass

        book.save(('2020下半年全校课表.xls'))






if __name__ == '__main__':
    schedule = Schedule(200)
    schedule.get_save_data()
