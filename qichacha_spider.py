# -*- coding: gbk -*-

from bs4 import BeautifulSoup
import requests
import xlrd
import xlwt
from xlutils.copy import copy
import time
import winsound

#企查查网站爬虫类
class EnterpriseInfoSpider:
    def __init__(self):

        #文件相关
        self.excelPath = 'enterprise_data.xls'
        self.sheetName = 'details'
        self.workbook = None
        self.table = None
        self.beginRow = None

        # 目录页
        self.catalogUrl = "http://www.qichacha.com/search_index"

        # 详情页（前缀+firmXXXX+后缀）
        self.detailsUrl = "http://www.qichacha.com/company_getinfos"

        self.cookie = raw_input("请输入cookie:").decode("gbk").encode("utf-8")
        self.host = "www.qichacha.com"
        self.userAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.106 Safari/537.36"

        self.headers = {
            "cookie" : self.cookie,
            "host" : self.host,
            "user-agent" : self.userAgent
        }

        #数据字段名17个
        self.fields = ['公司名称','电话号码','邮箱','统一社会信用代码','注册号','组织机构代码','经营状态','公司类型','成立日期','法定代表人','注册资本',
                       '营业期限','登记机关','发照日期','公司规模','所属行业','英文名','曾用名','企业地址','经营范围']

    #爬虫开始前的一些预处理
    def init(self):

        try:
            #试探是否有该excel文件，#获取行数：workbook.sheets()[0].nrows
            readWorkbook = xlrd.open_workbook(self.excelPath)
            self.beginRow = readWorkbook.sheets()[0].nrows #获取行数
            self.workbook = copy(readWorkbook)
            self.table = self.workbook.get_sheet(0)

        except Exception,e:
            print e
            self.workbook = xlwt.Workbook(encoding='utf-8')
            self.table = self.workbook.add_sheet(self.sheetName)

            #创建表头字段
            col = 0
            for field in self.fields:
                self.table.write(0,col,field.decode('gbk').encode('utf-8'))
                col += 1

            self.workbook.save(self.excelPath)
            self.beginRow = 1
            print "已在当前目录下创建enterprise_data.xls数据表"


    #从keyword/1页 得到的html中获得总页码数
    def getTotalPage(self,catalogPageCode):
        soup = BeautifulSoup(catalogPageCode,"html.parser")
        pagebar = soup.select("li #ajaxpage")
        if pagebar == None or pagebar == []:
            return -1
        return int(soup.select("li #ajaxpage")[-1].string.strip(' .'))

    #从keyword/page页 得到html中的所有公司条目
    def getFirmIdDoms(self,catalogPageCode):
        soup = BeautifulSoup(catalogPageCode,"html.parser")
        return soup.select("#searchlist .table-search-list .tp2 a")

    #爬虫开始
    def start(self):
        keyword = raw_input("请输入关键字：").decode("gbk").encode("utf-8")
        while keyword != "end":
            #先获取keyword第一页内容的页码
            totalPage = self.getTotalPage(self.getCatalogPageCode(keyword, 1))
            if totalPage == -1:
                # 请求下一轮查询的关键字
                keyword = raw_input("爬取结束,请输入关键字：").decode("gbk").encode("utf-8")
                continue

            #模拟翻页操作
            for page in range(1,totalPage+1):

                print "正在爬取第",page,"页的数据,请稍等..."

                #获取第page页代码
                catalogPageCode = self.getCatalogPageCode(keyword,page)
                firmIdDoms = self.getFirmIdDoms(catalogPageCode)
                for firmIdDom in firmIdDoms:
                    firmId = firmIdDom['href'][6:-6]
                    companyname = "" #公司名
                    for string in firmIdDom.strings:
                        companyname += string

                    tdDom = firmIdDom.find_parent().find_parent()
                    phoneDom = tdDom.select('.i-phone3')
                    emailDom = tdDom.select('.fa-envelope-o')
                    phone = ""
                    email = ""
                    if phoneDom != None and phoneDom != []:
                        phone = phoneDom[0].next_sibling.strip() #手机
                    if emailDom != None and emailDom != []:
                        email = emailDom[0].next_sibling.strip() #邮箱

                    detailsPageCode = self.getDetailsPageCode(firmId,companyname)
                    self.writeDetailsToExcel(detailsPageCode,companyname,phone,email)
                    time.sleep(0.3) #0.5s后再爬防止反爬虫机制

            #请求下一轮查询的关键字
            keyword = raw_input("爬取结束,请输入关键字：").decode("gbk").encode("utf-8")

        print "爬虫已完全结束！"

    #根据keyword和page构造查询串
    #其中keyword中的空格换成+
    #返回查询字符串构成的字典
    def getCatalogQueryString(self,keyword,page):
        keyword.replace(' ','+')
        return {"key": keyword, "index": "0", "p": page}

    def getDetailQueryString(self,firmId,companyname):
        return {"unique": firmId, "companyname":companyname,"tab": "base"}

    # 根据keyword关键字获取目录页代码
    def getCatalogPageCode(self, keyword, page):
        queryString = self.getCatalogQueryString(keyword, page)
        response = requests.request("GET", self.catalogUrl, headers=self.headers, params=queryString)
        return response.text

    # 根据firmId获取公司的详情页代码
    def getDetailsPageCode(self,firmId,companyname):
        queryString = self.getDetailQueryString(firmId,companyname)
        response = requests.request("GET", self.detailsUrl, headers=self.headers, params=queryString)
        return response.text

    #抓取detailsPageCode页上该企业所有信息，并存入excel
    def writeDetailsToExcel(self,detailsPageCode,companyname,phone,email):
        detailDoms = self.getDetailDoms(detailsPageCode)

        self.table.write(self.beginRow, 0, companyname)
        self.table.write(self.beginRow, 1, phone)
        self.table.write(self.beginRow, 2, email)

        col = 3
        for detailDom in detailDoms:
            detailName = detailDom.label.string.strip()[:-1]
            detailValue = detailDom.label.next_sibling.string.strip()
            while col < len(self.fields):
                # 找到匹配的那列字段
                if detailName == self.fields[col].decode('gbk'):
                    self.table.write(self.beginRow, col, detailValue) #写入excel
                    col += 1
                    break
                else:
                    col += 1
        self.workbook.save(self.excelPath)  # 保存至文件
        self.beginRow += 1

    #根据detailsPageCode获得它的所有detailDoms元素
    def getDetailDoms(self,detailsPageCode):
        soup = BeautifulSoup(detailsPageCode,"html.parser")
        return soup.select(".company-base li")



########
#爬虫入口
########
spider = EnterpriseInfoSpider()
spider.init()
spider.start()
