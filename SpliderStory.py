# coding=utf-8
import urllib2
import re


class Spilder:

    def __init__(self):
        super

    def load_page(self, page):
        url = 'http://www.ruiwen.com/wenxue/gushihui/3146' + str(page) + '.html'  # type: str
        user_agent = 'Mozilla/5.0 (compatible; MSIE 9.0; Windows NT6.1; Trident/5.0'
        headers = {'User-Agent': user_agent}
        request = urllib2.Request(url, headers=headers)
        try:
            response = urllib2.urlopen(request)
            html = response.read().decode('gbk').encode('utf-8')
            # 从html中筛选出故事
            pattern = re.compile(r'<div.*?class="content">(.*?)</div>', re.S)
            data_list = pattern.findall(html)
            for item in data_list:
                # 根据h标签分成多个故事
                pattern_h = re.compile(r'<h2>(.*?)</h2>', re.S)
                storys = re.split(pattern_h, item)
                file = 'E:\storys\story1.txt'
                for item_h in storys:
                    handle = open(file, 'a')
                    pattern_p = re.compile(r'<p>|</p>')
                    pattern_blank = re.compile(r'&rdquo;|&ldquo;|&hellip;')
                    handle.write(pattern_blank.sub('  ', pattern_p.sub("", item_h)))
                    handle.write("---" * 30)
                    print item_h
        except (urllib2.HTTPError, UnicodeDecodeError, AttributeError) as error:
            print error.message
            # self.load_page(page + 1)

    pass

    # def write_to_file(self,story_text):


if __name__ == '__main__':
    splider = Spilder()
    for i in range(1, 100, 1):
        splider.load_page("%2d" % i)
    pass
