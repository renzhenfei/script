# coding=utf-8
import threading
import time
import json
from Queue import Queue
from lxml import etree
import requests

data_queue = Queue()
lock = threading.Lock()
exitFlag_parser = True


class thread_crawl(threading.Thread):

    def __init__(self, thread_id, q):
        threading.Thread.__init__(self)
        self.q = q
        self.thread_id = thread_id

    def run(self):
        self.qiushi_spider()
        pass

    def qiushi_spider(self):
        while True:
            if self.q.empty():
                break
            else:
                page = self.q.get()
                url = 'http://www.qiushibaike.com/8hr/page/' + str(page) + '/'
                headers = {"User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) "
                                         "Chrome/52.0.2743.116 Safari/537.36 "
                                         "Accept-Language", "zh-CN,zh;q=0.8"}
                try:
                    content = requests.get(url, headers=headers)
                    data_queue.put(content.text)
                except Exception, e:
                    print e.message
        pass


class Thread_Parser(threading.Thread):
    def __init__(self, thread_id, queue, lock1, f):
        super(Thread_Parser, self).__init__()
        self.f = f
        self.lock1 = lock1
        self.queue = queue
        self.thread_id = thread_id

    def run(self):
        while not exitFlag_parser:
            try:
                page = self.queue.get(True)
                if not page:
                    pass
                self.parse_data(page)
                self.queue.task_done()
                pass
            except Exception, e:
                print e.message
        pass

    def parse_data(self, page):
        try:
            html = etree.HTML(page)
            result = html.xpath("//div[contains(@id,'qiushi_tag')]")
            for item in result:
                author_info = item.xpath('.//div[@class="author clearfix"]')[0]
                user_avatar_url = author_info.xpath(".//img/@src")[0]
                user_name = author_info.xpath('.//a[@target="_blank"]/h2')[0].text
                content = item.xpath(".//div[@class='content']/span")[0].text
                img_url = ""
                try:
                    img_url = item.xpath(".//div/a/img/@src")[0]
                except Exception, e:
                    print e.message
                    pass
                result = {"user_avatar_url": user_avatar_url,
                          "user_name": user_name,
                          "content": content,
                          "img_url": img_url}
                with self.lock1:
                    self.f.write(json.dumps(result, ensure_ascii=False).encode('utf-8') + '/n')
            pass
        except Exception, e:
            print e.message
        pass


if __name__ == '__main__':
    output = open('joke.json', 'a')
    page_queue = Queue(50)
    for page in range(1, 50):
        page_queue.put(page)
    # 初始化采集线程
    crawl_threads = []
    craw_list = ["craw-1", "craw-2", "craw-3"]
    for thread_id in craw_list:
        thread = thread_crawl(thread_id, page_queue)
        thread.start()
        crawl_threads.append(thread)
    # 初始化解析线程parserList
    parse_threads = []
    parse_list = ["parse-1", "parse-2", "parse-3"]
    for thread_id in parse_list:
        thread = Thread_Parser(thread_id, data_queue, lock, output)
        thread.start()
        parse_threads.append(thread)
    # 等待队列清空
    while not page_queue.empty():
        pass
    # 等待所有线程完成
    for t in crawl_threads:
        t.join()
    while not page_queue.empty():
        pass
    exitFlag_parser = True
    for t in parse_threads:
        t.join()
    with lock:
        output.close()
    pass
