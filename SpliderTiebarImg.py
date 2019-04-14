# coding=utf-8
import urllib
import urllib2
from lxml import etree


class Splider:

    def __init__(self):
        self.name = raw_input("input name of tieba:")
        self.start_page = int(raw_input("input index of start page:"))
        self.end_page = int(raw_input("input index of end page:"))

        self.url = 'http://tieba.baidu.com/f'
        self.url_header = {"User-Agent": "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1 Trident/5.0;"}
        self.user_name = 1  # 图片编号

    def tiebar_splider(self):
        for page in range(self.start_page, self.end_page + 1):
            pn = (page - 1) * 50
            word = {"pn": pn, "kw": self.name}
            my_url = self.url + "?" + urllib.urlencode(word)
            self.load_page(my_url)

    def load_page(self, my_url):
        request = urllib2.Request(my_url, headers=self.url_header)
        html = urllib2.urlopen(request).read()
        # 解析html
        selector = etree.HTML(html)
        links = selector.xpath('//div[@class="threadlist_lz clearfix"]/div/a/@href')

        for link in links:
            link = "http://tieba.baidu.com" + link
            self.load_img(link)

    def load_img(self, link):
        req = urllib2.Request(link, headers=self.url_header)
        html = urllib2.urlopen(req).read()
        selector = etree.HTML(html)
        img_links = selector.xpath("//img[@class='BDE_Image']/@src")
        self.write_img_2_local(img_links)

    def write_img_2_local(self, img_links):
        for link in img_links:
            file = open("E:\storys\img\\" + str(self.user_name) + ".png", 'wb')
            image = urllib2.urlopen(link).read()
            file.write(image)
            file.close()
            self.user_name += 1


if __name__ == '__main__':
    Splider().tiebar_splider()
