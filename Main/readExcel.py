#coding=utf-8
'''
Created on 2015-1-18
'''
import os
import sys
import xlrd 
import time
import unittest
import HTMLTestRunner
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
reload(sys)
sys.path.append('..')
sys.setdefaultencoding('utf8')  # @UndefinedVariable

class GetData(object):
    '''对应excel中的tab名称'''
    TestCases = 'Test Cases'
    Base = 'Base'
    TestSteps = 'Test Steps'
    Element = 'Element'
    DATA = 'DATA'
    
    def __init__(self,path = r'..\data\data.xls'):
        self.data = xlrd.open_workbook(path)
        self.table = None
    
    def getCaseNames(self):
        '''获取用例名称
            @return : casenames type:list 返回用例名称的列表
        '''
        self.table = self.data.sheet_by_name(self.TestCases)
        nrows = self.table.nrows
        casenames = []
        for i  in range(1,nrows):
            casename = self.table.row_values(i)
            if casename[2] == 'Y':
                casenames.append(casename[0])
        return casenames
    
    def getCaseStep(self,casename):
        '''获取用例的执行步骤
            @param casename: 测试用例名称
            @return: 以字典形式返回用例的执行步骤
        '''
        self.table = self.data.sheet_by_name(self.TestSteps)
        casenames = self.table.col_values(0)
        count = casenames.count(casename)
        start = casenames.index(casename)
        end = start+count
        return self.getKey_Obj(casename, start, end)
    
    def getKey_Obj(self,casename,start,end):   
        '''根据用例名称，开始位置，结束位置获得用例的执行步骤
            @param kw: 关键字
            @param e_id: 元素对象id
            @param da: 测试数据
            @param ast: 断言
            @return: mp type dict ，以字典形式返回
        '''
        self.table = self.data.sheet_by_name(self.TestSteps)
        kw = self.table.col_values(3)[start:end]
        e_id = self.table.col_values(4)[start:end]
        da = self.table.col_values(5)[start:end]
        ast = self.table.col_values(6)[start:end]
        mp = {}
        mp[casename] = zip(kw,e_id,da,ast)
        return mp
        
    def getElementPath(self,e_id):
        '''获取元素定位方法
            @param e_id: 元素id 
            @return: 以元组方式返回
        '''
        self.table = self.data.sheet_by_name(self.Element)
        nrows = self.table.nrows
        for i in range(0,nrows):
            if e_id in self.table.row_values(i):
                return tuple(self.table.row_values(i)[1:])
    
    def getContent(self,d):
        '''获取测试数据 但是这个方法感觉不怎么好
        @return: 返回测试数据
        '''
        self.table = self.data.sheet_by_name(self.DATA)
        nrows = self.table.nrows
        for i in range(0,nrows):
            if d in self.table.row_values(i):
                return self.table.row_values(i)
    
    def getType_browser_url(self):
        '''仅仅是为了获取浏览器类型和url 方法很烂后期要改 
            @return: browserType: 浏览器对象，url：地址
        '''
        self.table = self.data.sheet_by_name(self.Base)
        for i in range(1,self.table.nrows):
            if i[0] == 'browserType':
                browserType = i[1]
            if i[0] == 'url':
                url = i[1]
        return browserType,url

    
class keywordswitch(object):
    '''根据从excel表格中获取到对应的关键字运行相应的方法
    '''
    def __init__(self,driver):
        self.driver = driver
        self.data = GetData()
        self.operator = {
                    'OpenBrowser':self.OpenBrowser,
                    'InputContent':self.InputContent,
                    'ClickButton':self.ClickButton,
                    'GetTitle':self.GetTitle,
                    'GetText':self.GetText,
                    'AssertTitle':self.AssertTitle,
                    'CloseBrowser':self.CloseBrowser
                    }
    def runKey(self,key,val,content,ast):
        '''运行关键字方法
        @param key: 关键字
        @param val: 元素id
        @param content:输入框需要输入的值
        @param ast: 断言 
        '''
        self.operator.get(key)(val,content,ast)
    
    def OpenBrowser(self,e_id = None,content = None,ast = None):
        '''打开浏览器并输入地址
            @param e_id: excel中元素id
            @param content: 元素对应的输入值 
            @param ast : 断言
        '''
        self.driver.get('http://www.baidu.com/')
        
    def InputContent(self,e_id = None,content = None,ast = None):
        '''向控件输入值
            @param e_id: excel中元素id
            @param content: 元素对应的输入值 
            @param ast : 断言
        '''        
        loc = self.data.getElementPath(e_id)
        self.find_element(*loc).send_keys(content)
    
    def ClickButton(self,e_id = None,content = None,ast = None):
        '''按钮点击
            @param e_id: excel中元素id
            @param content: 元素对应的输入值 
            @param ast : 断言
        '''                
        loc = self.data.getElementPath(e_id)
        self.find_element(*loc).click()
        time.sleep(2)
    
    def GetText(self,e_id = None,content = None ,ast = None):
        '''获取文本
            @param e_id: excel中元素id
            @param content: 元素对应的输入值 
            @param ast : 断言
        '''                
        loc = self.data.getElementPath(e_id)
        return self.find_element(*loc).getText()
    
    def GetTitle(self,e_id = None,content = None,ast = None):
        '''获取title
            @param e_id: excel中元素id
            @param content: 元素对应的输入值 
            @param ast : 断言
        '''                
        return self.driver.title
    
    def AssertTitle(self,e_id = None,content = None,ast = None):
        '''Title断言
            @param e_id: excel中元素id
            @param content: 元素对应的输入值 
            @param ast : 断言
        '''                
        try:
            tl = self.driver.title
            assert tl == ast
        except AssertionError ,e :
            print e 
    
    def CloseBrowser(self,e_id = None,content = None,ast = None):
        '''关闭浏览器
            @param e_id: excel中元素id
            @param content: 元素对应的输入值 
            @param ast : 断言
        '''                
        self.driver.quit()
    
    def find_element(self,*loc):
        '''重写查找元素方法'''
        try:
            WebDriverWait(self.driver,5).until(lambda driver : driver.find_element(*loc).is_displayed())
            return self.driver.find_element(*loc)
        except:
            print u"%s 页面中未能找到 %s 元素" % (self, loc)
            self.save_error_pic(self.driver, 'find_element')            

    def save_ok_pic(self, driver, name):
        '''保存成功截图方法'''
        try:
            time.sleep(1)
            now = time.strftime("%Y-%m-%d-%H_%M_%S", time.localtime(time.time()))
            pic_path = r'..\report\image\%s--%s-ok.png' % (name, now)
            print '截图保存路径为:\n%s' % os.path.abspath(pic_path)
            driver.get_screenshot_as_file(pic_path)
        except:
            print u'保存截图失败'

    def save_error_pic(self, driver, name):
        '''保存失败截图方法'''
        try:
            time.sleep(1)
            now = time.strftime("%Y-%m-%d-%H_%M_%S", time.localtime(time.time()))
            pic_path = r'..\report\image\%s--%s--error.png' % (name, now)
            print '截图保存路径为:\n%s' % os.path.abspath(pic_path)
            driver.get_screenshot_as_file(pic_path)
        except:
            print u'保存截图失败'

class MyAssert(Exception):
    '''自定义异常
    '''
    pass

class MyTest(unittest.TestCase):
    
    def setUp(self):
        self.driver = webdriver.Chrome()
        self.ks = keywordswitch(self.driver)
    
    def action(self,casestep):
        '''测试用例
            @param casestep: 测试用例执行步骤
            @warning: i[0] = kw,i[1] = e_id,i[2] = data i[3] =ast
        '''
        for i in casestep:
            self.ks.runKey(i[0], i[1], i[2], i[3])
    
    @staticmethod
    def getTestFunc(casestep):
        def func(self):
            self.action(casestep)
        return func
    
    def tearDown(self):
        print 'END'


def __AddCase():
    '''动态添加测试方法'''
    data = GetData()
    casenames = data.getCaseNames()
    for casename in casenames:
        casestep = data.getCaseStep(casename).get(casename)
        setattr(MyTest,'test name : %s' %casename, MyTest.getTestFunc(casestep))

__AddCase()    

if __name__ == '__main__':
#     test_support.run_unittest(MyTest)

    testunit = unittest.TestSuite()
    testunit.addTest(MyTest('action'))
    now_time = time.strftime("%Y-%m-%d-%H_%M_%S",time.localtime())
    filename = '..\\report\\'+now_time+'result.html'
    fp = open(filename,'wb')
    runner = HTMLTestRunner.HTMLTestRunner(stream=fp,
                                           title = u'MyAutoTest',
                                           description=u'The Result')       
    runner.run(testunit)
        