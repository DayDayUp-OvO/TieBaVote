import re
import xlwt
import requests
import os
from lxml import etree


class Tieba():
    def __init__(self, tie_serialNumber):
        # 贴子序号
        self.tid = tie_serialNumber
        # 帖子链接
        self.tie_url = "https://tieba.baidu.com/p/" + \
            self.tid + "?see_lz=1&pn={}"
        # 设置请求头
        self.headers = {
            "Connection": "close",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36 Edg/92.0.902.78"
        }

    def getUrl(self, url):
        '''
        请求

        :param url:
        :return:
        '''
        return requests.get(url, self.headers)

    def getPid(self):
        '''
        获取每层的pid

        :return:
        '''
        pid_dict = {}
        # 获取所有页数链接
        res = self.getUrl(self.tie_url.format(1))
        try:
            end_url = etree.HTML(res.content.decode(
                "utf-8")).xpath("//a[contains(string(), '尾页')]/@href")[0]
        # 不知道为啥，这里爬到的是实际看到页数的两倍，但不能/2
            total_pn = int(int(end_url.split("pn=")[1]))
        except IndexError:
            total_pn = 1

        url_list = [self.tie_url.format(i) for i in range(1, total_pn + 1)]
        print("正在获取所有楼主层")
        for url_num in range(0, len(url_list)):
            res = self.getUrl(url_list[url_num])
            # print(res)
            etree_html = etree.HTML(res.content.decode("utf-8"))
            pid_list = etree_html.xpath(
                '//div[@class="l_post j_l_post l_post_bright  "]/@data-pid')
            for pid in pid_list:
                text = etree_html.xpath(
                    "//div[@id='post_content_{}']/text()".format(pid))[0]
                pid_dict[pid] = text.replace(" ", '')
            print("\r正在获取pid {:.2f}%".format(
                url_num * 100 / (len(url_list) - 1)), end='')

        print()
        return pid_dict

    def getComment(self, pid):
        '''
        获取当前层的楼中楼评分，正则过滤  只取0-10分

        :param pid:
        :return:
        '''
        total_soure = 0
        c_url = "https://tieba.baidu.com/p/comment?tid={}&pid={}".format(
            self.tid, pid)
        res = self.getUrl(c_url + "&pn=1")

        etree_html = etree.HTML(res.content.decode("utf-8"))
        try:
            end_url = etree_html.xpath(
                "//a[contains(string(), '尾页')]/@href")[0]
            total_pn = int(int(end_url.split("#")[1]))
        except IndexError:
            # print("error")
            total_pn = 1

        name_list = []
        re_filter = "[0-9]+"
        for pn in range(1, total_pn + 1):
            pn_res = self.getUrl(c_url + "&pn=" + str(pn))
            pn_etree_html = etree.HTML(pn_res.content.decode("utf-8"))
            lzl_list = pn_etree_html.xpath('//div[@class="lzl_cnt"]')
            for lzl in lzl_list:
                name = lzl.xpath("a/@username")[0]
                if name == "":
                    name = lzl.xpath("a/text()")[0]
                if name not in name_list:
                    text = lzl.xpath(
                        'span[@class="lzl_content_main"]/text()')[-1]
                    try:
                        soure = int(re.findall(re_filter, text)[0])
                        # print(soure)
                        if 0 <= soure <= 10:
                            name_list.append(lzl.xpath("a/@username"))
                            total_soure += soure
                    except:
                        pass
        return [total_soure, len(name_list)]

    def execute(self):
        '''
        执行并写入

        :return:
        '''
        # 文件
        workbook = xlwt.Workbook(encoding='utf-8')  # 创建workbook对象
        # 表单
        worksheet = workbook.add_sheet('sheet1')  # 创建工作表
        # 表头
        worksheet.write(0, 0, "标题")
        worksheet.write(0, 1, "得分")
        worksheet.write(0, 2, "有效评论人数")
        worksheet.write(0, 3, "平均分")

        num = 1

        pid_dict = tb.getPid()
        for pid in pid_dict.keys():
            print("\r正在统计楼中楼得分{:.2f}%".format(
                num * 100 / (len(pid_dict))), end='')
            get_r = self.getComment(pid)
            worksheet.write(num, 0, pid_dict[pid])
            worksheet.write(num, 1, get_r[0])
            worksheet.write(num, 2, get_r[1])
            try:
                worksheet.write(num, 3, get_r[0] / get_r[1])
            except:
                worksheet.write(num, 3, "无有效评论")
            num += 1
        print()
        workbook.save("统计.xls")
        print("统计完成，文件已生成在当前路径的文件夹下\n路径：" + os.getcwd() + "\统计.xls")


while True:
    try:
        get_id = eval(input("请输入需要统计的帖子id：\n"))
        tb = Tieba(str(get_id))
        tb.execute()
    except:
        print("发生异常，请重试！")