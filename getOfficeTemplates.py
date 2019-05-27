# 从office官网下载所有office模板
# 按office组件类别整理
from pprint import pprint
import requests
from lxml import etree
import time
import random
import os
import re

requests.DEFAULT_RETRIES = 5


# noinspection PyBroadException
class OfficeTemplate:
    # 类初始化
    def __init__(self):
        self.basePage = 'https://templates.office.com/zh-CN'
        self.baseLink = 'https://templates.office.com'
        self.headers = {
            'Host': 'templates.office.com',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36'}

    # 下网网页源文件
    def getResponse(self, url, bl_bit=0, ref=None):
        time.sleep(random.uniform(1, 1.5))
        try:
            if bl_bit == 1:
                if ref is not None:
                    self.headers['Referer'] = ref
                    print(self.headers)
                    r = requests.get(url, headers=self.headers)
                    r.raise_for_status()
                    return r.content
            elif bl_bit == 0:
                r = requests.get(url, headers=self.headers)
                r.raise_for_status()
                return r.text
        except Exception as e:
            print(e)
            return None

    # 获取目录类别及链接
    def get_categories(self):
        r = self.getResponse(self.basePage)
        if r is not None:
            sel = etree.HTML(r)
            links = [
                self.baseLink +
                eve for eve in sel.xpath('//*[@id="categories"]/div/div/p/a/@href')]
            ctg = sel.xpath('//*[@id="categories"]/div/div/p/a/text()')
            pprint(list(zip(links, ctg)))
            return list(zip(links, ctg))

    # 按页类别获取所有模板名称及链接
    def get_detail_links(self, ctg):
        link = ctg[0]
        r = self.getResponse(link)
        all_info = []
        if len(r):
            pages_num = [int(pageNum) for pageNum in re.findall(r'Page (\d+)', r)]
            if len(pages_num) != 0:
                maxPageNum = max(pages_num)
                detail_links = (
                    '{}?page={}'.format(
                        link, pagenum) for pagenum in range(
                    1, maxPageNum + 1))
                for page in detail_links:
                    try:
                        page_r = self.getResponse(page)
                        sel = etree.HTML(page_r)
                        res_links = [
                            self.baseLink +
                            eve for eve in sel.xpath("//div[@class='odcom-template-tile-anchor-wrapper']/a/@href")]
                        file_names = [eve.strip() for eve in sel.xpath(
                            "//div[@class='odcom-card-template-tile-info-title']/text()")]
                        pprint(list(zip(res_links, file_names)))
                        all_info.extend(list(zip(res_links, file_names)))

                    except Exception as e:
                        print('按页解析失败-{}'.format(page))
                return [ctg[1], all_info]
            else:
                return None
        else:
            return None

    # 逐个链接下载保存模板
    def save_files(self, detail_info):
        # print(detail_info)
        ctg = detail_info[0]
        file_link, file_name = detail_info[1]
        re_str = r"[\/\\\:\*\?\"\<\>\|]"  # '/ \ : * ? " < > |'
        file_name = re.sub(re_str, "_", file_name)
        # print('正在参数保存文件：{}'.format(file_link))
        r = self.getResponse(file_link, ref=self.headers)
        if r is not None:
            root = etree.HTML(r)
            data_app = root.xpath(
                '//*[@id="endNodeMainContent"]/div[2]/div/div[1]/section/div/a[1]/@data-app')

            file_path = os.path.join(
                r'D:\OFFICE模板', data_app[0].strip())  # 一级--APP

            file_path = os.path.join(file_path, ctg)  # -二级类别

            res_file_name = os.path.join(file_path, file_name)  # -三级名细

            # print(file_path)
            print('文件名：' + res_file_name)
            if not os.path.exists(file_path):
                os.makedirs(file_path)

            data_link = root.xpath(
                '//*[@id="endNodeMainContent"]/div[2]/div/div[1]/section/div/a[1]/@href')
            try:
                r = requests.get(data_link[0])
                file_type = r.headers['Content-Disposition'].split(r'.')[1]
                save_file_name = r"{}.{}".format(res_file_name, file_type)
                print(save_file_name)
                if not os.path.exists(save_file_name):
                    with open(save_file_name, 'wb') as f:
                        content = r.content
                        if r is not None:
                            f.write(content)
                            print('下载完成->[{}-->{}--->{}]<-'.format(data_app[0], file_name, data_link))
                            print('=' * 60)
                            print('\n')
                else:
                    print('已存在：{}'.format(file_name))
            except Exception as e:
                print(e)

    # 主下载
    def main(self):
        ctg = self.get_categories()
        for ctf_info in ctg:
            dtl = self.get_detail_links(ctf_info)
            if dtl is not None:
                for eve in dtl[1]:
                    self.save_files([dtl[0], eve])


if __name__ == '__main__':
    tempt = OfficeTemplate()
    tempt.main()
