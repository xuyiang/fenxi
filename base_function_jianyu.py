# coding:utf-8

import base_function as base_fun
import json

import re
from datetime import  datetime,date
from dateutil.relativedelta import relativedelta
#import time
from time import sleep

from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.support.ui import  WebDriverWait
from selenium.webdriver.support import  expected_conditions as EC
from selenium.common import TimeoutException

from  openai import OpenAI

import config as cfg


class Struct_business_opportunity:
    # 将值按照column_order_list顺序压入list. column_order_list 所有元素为字符串
    def push_data_into_list(self):
        if self.jianyu_link is None or len(self.jianyu_link)<1:  #空的链接
            jianyu_link_string = "剑鱼无链接"
        else:
            decode_hyper_string = base_fun.decode_keyword_in_url(self.jianyu_link, "link")
            jianyu_link_string = f'=hyperlink("{decode_hyper_string}","访问剑鱼链接")'

        if self.original_link is None or len(self.original_link)<1:  #空的链接
            original_link_string = "原始无链接"
        else:
            decode_hyper_string = base_fun.decode_keyword_in_url(self.original_link, "link")
            original_link_string = f'=hyperlink("{decode_hyper_string}","访问原始链接")'
        #联系方式：
        contact_method_list = []
        if len(self.tenderer_contact) > 0:
            contact_method_list.append(f"【采购方联系方式】：\n{self.tenderer_contact}")
        if len(self.agency_contact) > 0:
            contact_method_list.append(f"【采购方联系方式】：\n{self.agency_contact}")
        contact_method = "\n".join(contact_method_list)

        value_list = [self.bidding_type, self.title, self.project_tags,
                      self.relevance ,
                      #summary_text,complete_text, #与信息清洗，同步调整 @lqh 20250410
                      self.summary_info,self.complete_info,
                      self.budget,
                      jianyu_link_string, #网站链接
                      original_link_string, #原始网页链接
                      #hyper_link,
                      self.tenderer, self.bid_agency,
                      contact_method,
                      self.release_date,   #发布时间在爬取时，统一调整为日期格式
                      self.deadline,
                      self.sales_department , self.sales,
                      self.collected_date]
        return value_list


    def __init__(self):
        self._columns_name = ['类型', '标题','项目标签',
                              '业务相关性',
                              '公告摘要','公告正文', '预算金额(万)',
                              '网站链接','原始链接',
                              '采购单位', '招标代理','联系方式',
                              '发布日期','报名截至日期',
                              '销售部门','客户代表',
                              '信息采集时间']
        self.bidding_type = ""
        self.title = ""
        self.project_tags = ""
        self.relevance = ""
        self.summary_info = ""
        self.complete_info = ""
        self.budget = 0.0
        self.jianyu_link = ""
        self.original_link =""
        self.tenderer = ""
        self.tenderer_contact = ""
        self.bid_agency = ""
        self.agency_contact =""
        self.release_date = ""
        self.deadline =""
        self.sales_department = ""
        self.sales = ""
        self.collected_date = ""

    # 限制_columns_name ,不可被修改
    @property
    def columns_name(self):
        return self._columns_name

    def re_init(self):
        self.bidding_type = ""
        self.title = ""
        self.project_tags = ""
        self.relevance = ""
        self.summary_info = ""
        self.complete_info = ""
        self.budget = 0.0
        self.jianyu_link = ""
        self.original_link = ""
        self.tenderer = ""
        self.tenderer_contact = ""
        self.bid_agency = ""
        self.agency_contact = ""
        self.release_date = ""
        self.deadline = ""
        self.sales_department = ""
        self.sales = ""
        self.collected_date = ""

def login(web_driver:webdriver , return_current_url = False):
    current_url = web_driver.current_url
    current_zoom = web_driver.execute_script("return document.body.style.zoom;")

    web_driver.get('https://www.jianyu360.cn/')
    base_fun.set_document_zoom_scale(web_driver,1,0.5)
    #获取登录按钮
    login_element = WebDriverWait(web_driver,20).until(EC.visibility_of_element_located((By.XPATH,'/html/body/section/header/section/section[1]/main/div[2]/ul/li[8]/button')))
    #判断登录状态
    str_style = login_element.get_attribute("style")
    if "display: none" in str_style:  #用户已经登录，返回
        #恢复原网页链接
        web_driver.get(current_url)
        web_driver.execute_script(f"document.body.style.zoom='{current_zoom}'")

        base_fun.wait_document_ready(web_driver, 20)
        print("用户已经登录，从登录函数退出！")
        return
    login_element = WebDriverWait(web_driver, 20).until(EC.element_to_be_clickable(
        (By.XPATH, '/html/body/section/header/section/section[1]/main/div[2]/ul/li[8]/button')))
    login_element.click()
    #点击密码登录
    WebDriverWait(web_driver, 20).until(EC.element_to_be_clickable(
        (By.XPATH, '/html/body/section/header/div[1]/div/div[1]/div[2]/div[3]/span[2]'))).click()

    #输入用户名，密码
    username_x_path = "/html/body/section/header/div[1]/div/div[1]/div[2]/div[5]/div[1]/input"
    passwd_x_path =   "/html/body/section/header/div[1]/div/div[1]/div[2]/div[5]/div[2]/input"
    username_element = WebDriverWait(web_driver, 20).until(EC.presence_of_element_located((By.XPATH,username_x_path)))
    passwd_element = WebDriverWait(web_driver, 20).until(EC.presence_of_element_located((By.XPATH, passwd_x_path)))
    username_element.send_keys(cfg.username)
    passwd_element.send_keys(cfg.passwd)
    #略微休息一下
    base_fun.wait_after_operation(cfg.time_after_input)
    #提交
    submit_button_x_path = "/html/body/section/header/div[1]/div/div[1]/div[2]/div[5]/button"
    WebDriverWait(web_driver, 20).until(EC.element_to_be_clickable((By.XPATH, submit_button_x_path))).click()
    base_fun.wait_document_ready(web_driver,20)
    #回到以前的页面
    if return_current_url:
        web_driver.get(current_url)
        web_driver.execute_script(f"document.body.style.zoom='{current_zoom}'")

#切换至iframe
def switch_to_iframe_jianyu(web_driver, retry = False):
    """
    剑鱼特有的数据存储结构，页面中存在iframe , 主页面和详情页面都可以用。
    :param web_driver:
    :param retry:  默认 False，如果为True,持续循环。直至切换进入 iframe
    :return: 成功切换，返回 true,否则返回  false
    """
    web_driver.switch_to.default_content()
    print("尝试切换至 iframe")
    iframe_x_path= ".//iframe[@class='iframe-container']"
    result = False
    while not result:
        try:
            iframe_element = WebDriverWait(web_driver,10).until(EC.presence_of_element_located((By.XPATH, iframe_x_path)))
            web_driver.switch_to.frame(iframe_element)
            result = True
        except  Exception :
            result = False
        if not retry :
            break
        if not result :
            input("在函数 switch_to_iframe_jianyu 切换至 iframe 错误，请尝试手动解决后，回车继续执行")
    if result:
        print("成功切换至 iframe！")
    else:
        print("切换至 iframe 失败！")
    return result


#计算两个日期之间日期跨度，
def calculate_date_span(date_or_datetime_begin, date_or_datetime_end)->str:
    """
    计算时间开度选择对应的发布时间选线
    :param date_or_datetime_begin: 类型： datetime ,date均可
    :param date_or_datetime_end:
    :return:
    """
    #为了确保线程可用，以后尽量不对参数进行修改。
    date_begin = date_or_datetime_begin
    date_end = date_or_datetime_end
    if date_begin is None or date_end is None :
        return "all"
    if isinstance(date_begin,datetime) :
        date_begin = date_begin.date()
    if isinstance(date_end, datetime) :
        date_end = date_end.date()
    delta = relativedelta(date_end,date_begin)
    years = delta.years
    months = delta.months
    days = delta.days
    if years > 3 or (years ==3 and (months>0 or days>0)):
        date_span = "last5year"
    elif years >1   or (years ==1 and (months>0 or days>0)):
        date_span = "last3year"
    elif months >=1  or days>30:
        date_span = "last1year"
    elif days > 7:
        date_span = "last30day"
    else:
        date_span = "last7days"
    return date_span


#判断dataframe 中是否存在指定的记录：
def is_record_exist(df_records, bid_type, title,tenderer, release_date,budget, jianyu_hyperlink)->bool:
    """
    根据输入参数，判断是否存在相同记录
    :param df_records:
    :param bid_type:
    :param title:
    :param tenderer:
    :param release_date:
    :param budget:
    :param jianyu_hyperlink:
    :return:
    """
    if df_records is None :
        return False
    df_filter = df_records[df_records["采购单位"] == tenderer]
    if len(df_filter) == 0:
        return False

    str_query_1 = f'类型 ==@bid_type and 发布日期==@release_date and 标题 ==@title'
    df_filter = df_filter.query(str_query_1)
    if len(df_filter)==0:
        return False

    df_filter = df_filter[df_filter["预算金额(万)"]==budget]
    # 允许一定误差范围内视为相等（例如误差小于 1e-9）
    #df_filter = df_filter[numpy.isclose(df_filter['预算金额(万)'], budget, atol=1e-9)]
    if len(df_filter)==0:
        return False
    #print(len(df_filter))
    df_filter = df_filter[df_filter["网站链接"] == jianyu_hyperlink]
    if len(df_filter)==0:
        return False
    #print(df_filter2.to_string(max_rows=None, max_cols=None))
    return True

#对剑鱼结构中的complete_info进行分析，若内容成功分析，修改 relevance ， summary_info，
# 若complete_info ,不符合要求，使用 title进行判断， 不修改summary_info
# 使用 base_fun.is_valid_content ,判断是否可进行AI分析
def AI_analysis_content_in_jianyu(jianyu_business_opportunity: Struct_business_opportunity)->bool:
    """
    对剑鱼结构中的complete_info进行分析，若内容成功分析，修改 relevance ， summary_info，
    若complete_info ,不符合要求，使用 title进行判断， 不修改summary_info
    :param jianyu_business_opportunity: 待分析的内容
    :return:分析成功True, 内容无法分析，返回False
    """

    is_title_valid = False
    user_content = jianyu_business_opportunity.complete_info
    is_complete_info_valid = base_fun.is_valid_content(user_content)
    if is_complete_info_valid :
        user_content = jianyu_business_opportunity.complete_info
    else:
        user_content = jianyu_business_opportunity.title
        is_title_valid = base_fun.is_valid_content(user_content,min_len=10)
    if not is_complete_info_valid and not is_title_valid :
        return False

    sys_content = """你是单位的一名商机采集员，你的职责是商机采集及分析。
                    一、你所在的单位业务范围：
                    1)通信运营商服务内容，通讯服务、短信、算力服务、云计算、网络接入等，
                    2)系统集成商服务内容，包括弱电智能化、软硬件信息系统集成、信息系统软件开发、信息系统硬件采购等工作。
                    3)公司业务不包括监理服务。
                    二、项目与公司业务的相关性判断方法：
                    1）判断项目的相关性，基本准则是项目内容与是否符合公司的业务范围。
                    2）业务与公司业务相关性，按照从高到底的顺序为： 很高、一般、轻微、无关。若较难判断，则相关性为：待定。
                    3）对于新建或维修大型建筑物，如：教学楼、宿舍楼、园区、新院区、新校区等工程项目。后续很有可能存在有弱电智能化、安防系统等的项目。相关性为：潜在商机。
                    4）工程项目中的绿化，水务，河道，道路工程，桩基等，无信息化业务需求。
                    5）家具购买，无信息化业务需求。
                    三、项目分析判断输出：
                    从获取的招标信息，采购意向等，概括出《项目概况》、《项目内容》、《投标人资格要求》及《项目相关性》4项。
                     1）项目概况：主要包括项目类型及项目预算。项目类型如：通讯服务、网络接入、工程建设，系统集成，软件系统，软硬件采购，硬件采购，软件采购等等。项目金额以万元为单位。 
                     2）项目内容：主要为实施内容及数量概况。如购买112台设备，39套家具，2套软件系统，建筑面积等。后面罗列3-5项内容。总字数100以内。
                     3）投标人资格要求：主要关注资质方面的要求。
                     4）项目相关性：按照前述的单位业务范围，项目相关性判断方法，判断项目的相关性。
                    以 JSON 的形式输出项目数量。遵循以下JSON 格式：
                      {\"项目相关性\":<项目相关性>,
                       \"项目信息\":{ 
                                    \"项目概况\"：<项目概况。文字描述。不使用JSON 格式。>，
                                    \"项目内容\"：<项目内容。文字描述。不使用JSON 格式。>，
                                    \"投标人资格要求\"：<概括投标人资质要求，50字以内。文字描述。不使用JSON 格式。>
                                   }"""
    client = OpenAI(
        base_url="https://api.deepseek.com/",
        api_key="sk-c549e0e1692844a7aeac77dc5eb92b5c"
    )

    completion = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {
                "role": "system",
                "content": sys_content},
            {
                "role": "user",
                "content": user_content
            }
        ],
        response_format={
            'type': 'json_object'
        }
    )
    try :
        response_dict = json.loads(completion.choices[0].message.content)
        if is_complete_info_valid:
            jianyu_business_opportunity.relevance = response_dict["项目相关性"]
            project_info = response_dict["项目信息"]
            jianyu_business_opportunity.summary_info = '\n'.join(
                [f"【{key}】： {value}" for key, value in project_info.items()])
        else: #使用标题判断的情况
            jianyu_business_opportunity.relevance = f'简判\n{response_dict["项目相关性"]}'
    except Exception :
        print(f"解析项目：{jianyu_business_opportunity.title} 信息出错")
        return False
    return True

#循环进行内容采集。并在config 中设定的AI 分析时间段内，进行内容分析
def analysis_business_opportunity(browse_driver, company_info, df_exist_records, business_opportunity_list, success_record_num, datetime_begin=None, datetime_end = None, analysis_statistic_result = None) -> int:
    """
    返回本次成功增加的记录数量
    :param browse_driver: 网页浏览器驱动。当前正在使用该该驱动器操作当前需要解析的页面，
    :param company_info: 当前被查询采购意向信息的公司名称
    :param df_exist_records: 在本次查询时间范围内，已经存在记录。
    :param business_opportunity_list:记录有效的商机息， 格式 list
    :param success_record_num: 当前已经分析符合要求的记录数
    :param datetime_begin: 时间范围起点
    :param datetime_end: 时间范围终点
    :param analysis_statistic_result:记录本轮分析的记录。 list:  总记录数， 有效记录数， 高相关性记录数：
    :return: 当前操作的结果。累计成功分析的记录数
    """

    # 存在iframe中，需要切换到iframe中进行定位
    operation_map = {
        "total_records": ".//div[@class='search-bidding-list-container']//div[@class='header-text--tip']/div/span[@class='highlight-text']",

        "search-result-container": ".//div[@class='search-result-container']",
        # 有多个，需要逐个解析
        "search-result-div": "./div",
        "total_pages": "./div[contains(@class, 'search-result-footer-container')]/div/ul/li[last()]",
        "next_page": "./div[contains(@class, 'search-result-footer-container')]/div/button[2]",
        "records_per_page_input": "./div[contains(@class, 'search-result-footer-container')]/div/span[@class='el-pagination__sizes']/div/div/input"
    }

    # 均是相对于 iframe
    new_page_info = {
        # 用于需要付费的信息
        "un_present": ".//div[@class = 'big-member-page in-app']//div[@class = 'content-mask-container com-prebuilt purchase']",
        # 用于公告摘要
        "summary": ".//div[@class='common-content-summary']//table[@class='el-table__body']/tbody",
        # 用于公告正文
        "announcement_label": "/html/body/div[4]/div/div/div/div/div[1]/div[1]/div/div[3]/div[1]/div[1]/div[2]",
        "announcement": ".//div[@class='first-content-container']/div[2][@class='content-card watch-tab-content']",
        "original_link": ".//button[@class='el-button origin-detail-action el-button--default']"

    }

    #格式处理相关的正则表达式
    title_pattern = re.compile('^\d+\.\s*(.*)')
    page_number_pattern = re.compile("\d+")

    #相关参数初始化：
    if analysis_statistic_result is not None and isinstance(analysis_statistic_result, list):
        # 总记录数， 有效记录数，高相关性记录数 0
        analysis_statistic_result.clear()
        analysis_statistic_result.append(0)
        analysis_statistic_result.append(0)
        analysis_statistic_result.append(0)
    else:
        analysis_statistic_result = [0,0,0]  #通过赋值产生的list ,不会被回传，但可简化后续处理。

    browse_driver.switch_to.default_content()
    switch_to_iframe_jianyu(browse_driver)
    #获取总记录条数
    # 解析当前的查询是否有符合要求的信息记录
    # 将页面拉到低端，查看页面数，虽然没有必要，但是为了防止意外，拖动一次.同时，记录拖动前的位置
    original_position = browse_driver.execute_script("return window.pageYOffset;")
    browse_driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    base_fun.wait_after_operation(cfg.time_after_scroll)
    # 赋值被查询公司名称
    company = company_info[0].strip()
    try:
        total_records = int(browse_driver.find_element(By.XPATH, operation_map["total_records"]).text.strip())
        search_result_container_element = browse_driver.find_element(By.XPATH, operation_map["search-result-container"])
        total_pages = int(search_result_container_element.find_element(By.XPATH, operation_map["total_pages"]).text.strip())

        #预设records_per_page 初值
        records_per_page = 0
        records_per_page_text = search_result_container_element.find_element(By.XPATH, operation_map[
            "records_per_page_input"]).get_attribute("value")
        match = page_number_pattern.match(records_per_page_text)
        if match:
            records_per_page = int(match.group())

        #如果不能找到 records_per_page的值，存在异常，需要退出本次查询
        if records_per_page <1:
            raise Exception("未能获取每页的记录数")

        #计算实际可显示的查询结果数量
        if total_pages * records_per_page < total_records:
            total_records = total_pages * records_per_page
        last_page_records = total_records % records_per_page
        #如果整除，需要将最后一页记录数调整为每页记录数
        if last_page_records == 0 and total_records > 0 :
            last_page_records = records_per_page
    except Exception as e:
        # 如page解析出错，那么直接向企微发送报错信息
        message = f'{company} 不存在查询结果，或者计算页面数据错误，退出函数def analysis_business_opportunity!'
        base_fun.to_qw_error(message, exception_info=e)
        return success_record_num
    #处理结束后，鼠标拖动回到初始位置
    browse_driver.execute_script(f"window.scrollTo(0, {original_position});")
    base_fun.wait_after_operation(cfg.time_after_scroll)
    page_num = 1
    available_records_amount = 0 #记录本轮符合查询条件的记录总数
    high_relevance_records_amount = 0 #记录本轮分析后符合高相关性的记录总数
    while page_num <= total_pages:
        '''
        if page_num < total_pages:
            div_amount = records_per_page
        else :
            #最后一页的记录数目
            div_amount = last_page_records
        '''
        print(f'当前处理{company}商机信息，待分析记录条数总计{total_records}条。当前分析第{page_num}/{total_pages}页数据')

        # 定位到表格
        try:
            search_result_container_element = browse_driver.find_element(By.XPATH, operation_map["search-result-container"])
        except Exception as e:
            message = f'单位：{company} ，处理第{page_num}/{total_pages}页，无法定位到查询结果容器，退出信息检索！'
            base_fun.to_qw_error(message, exception_info=e)
            # 然后退出while循环
            break

        # 开始处理本页的所有记录
        record_in_page = 0
        one_available_business_opportunity_record = Struct_business_opportunity()
        #每次动态计算本页的记录数. 等号右边-1 ， 是因为页脚div 也包装在 search_result_container_element  中
        div_amount = len(search_result_container_element.find_elements( By.XPATH,operation_map["search-result-div"]))-1
        while record_in_page < div_amount:
            one_available_business_opportunity_record.re_init()
            #根据当前单位名，赋值销售单位，及销售代表
            one_available_business_opportunity_record.sales_department = company_info[3]
            one_available_business_opportunity_record.sales = company_info[4]
            #当前是页内的第几条记录
            record_in_page += 1
            #跟随处理的记录进行同步页面滑动
            base_fun.scroll_window_synchronize(browse_driver,record_in_page,div_amount)
            print(f'当前处理单位：{company} ，第{page_num}/{total_pages}页，第{record_in_page}条数据。累计成功{success_record_num}条记录。')
            div_locator = f'{operation_map["search-result-div"]}[{record_in_page}]'
            #开始定位div
            try:
                div = search_result_container_element.find_element(By.XPATH, div_locator)
                list_item_content_element = div.find_element(By.XPATH, ".//div[@class='list-item-content']")
            except Exception as e:
                message = f'单位：{company} ，处理第{page_num}/{total_pages}页，无法定位第{record_in_page}条记录，跳转下一轮查询！'
                base_fun.to_qw_error(message, exception_info=e)
                continue

            # 解析标题
            title_x_path = "./div/div/div/div"   # 相对于 list-item-user_content 的路径
            try:
                title_hyper_link_element = list_item_content_element.find_element(By.XPATH, title_x_path)
                one_available_business_opportunity_record.title =title_hyper_link_element.text.strip()
                match = title_pattern.match(title_hyper_link_element.text.strip())
                if match:
                    one_available_business_opportunity_record.title = match.group(1).strip()
                # print(f'项目名称：{title}')
            except Exception:
                one_available_business_opportunity_record.title = "获取项目名称失败"
                title_hyper_link_element = None  #标注本项目无法获取超链接，获取项目详情

            #解析发布时间
            try:
                #相对于 list-item-user_content 的路径
                release_date_x_path = "./div/div/div[@class='time-container']/span[@class='time-text']"
                release_date_text  = list_item_content_element.find_element(By.XPATH, release_date_x_path).text.strip()

                # 获取发布时间
                one_available_business_opportunity_record.release_date = base_fun.string_to_datetime(release_date_text)
                if not base_fun.audit_date_time_range(one_available_business_opportunity_record.release_date, datetime_begin, datetime_end):
                    print(f'时间范围不符合要求，忽略本条记录')
                    # 时间不在有效区间，跳过
                    continue
                else: #对于符合要求的发布时间，统一转换为日期格式********************************************************************************
                    one_available_business_opportunity_record.release_date = one_available_business_opportunity_record.release_date.date()

            except Exception as e1:
                message = f'单位：{company} \n项目名称：{one_available_business_opportunity_record.title}\n第{page_num}/{total_pages}页，第{record_in_page}条数据，查询对应的发布时间出错!跳转下一条记录 '
                base_fun.to_qw_error(message, exception_info=e1)
                ##必须获取时间,不能获得，就跳过本记录
                continue


            #解析标签(其中，类型，预算进的， 单独赋值到struct business_opportunity 的其他变量等)
            try:
                #相对于 list-item-user_content 的路径
                tags_x_path = "./div[@class='a-i-right']/div[@class='tags']"
                tags_element = list_item_content_element.find_element(By.XPATH, tags_x_path)
                tag_elements = tags_element.find_elements(By.XPATH, "./span[contains(@class, 'tag')]")
                index = 0
                tags_text_list = []
                for one_element in tag_elements:
                    tag_text = one_element.text.strip()
                    index += 1
                    if index > 2:
                        budge_amount = base_fun.calculate_amount(tag_text)
                        if budge_amount:
                            one_available_business_opportunity_record.budget = budge_amount
                        else:
                            tags_text_list.append(tag_text)
                    elif index == 2:  # 采购类型
                        one_available_business_opportunity_record.bidding_type = tag_text
                    else:  # 区域 index == 1 , 区域用于项目实施位置定位
                        #continue
                        tags_text_list.append(tag_text)
                one_available_business_opportunity_record.project_tags = "|".join(tags_text_list)
            except Exception as e:
                message = f'单位：{company} \n项目名称：{one_available_business_opportunity_record.title}\n第{page_num}/{total_pages}页，第{record_in_page}条数据，查询对应的类型、预算金额、项目属性标签出错! '
                base_fun.to_qw_error(message, exception_info=e)


            # 解析详细列表：提取采购单位，采购单位联系方式， 代理机构联系方式，投标截至日期
            #这里初步设置项目概况信息 ，便于项目信息无法使用时刻，可以进行概况填写。
            one_available_business_opportunity_record.summary_info = []
            try:
                # 相对于 list-item-user_content 的路径
                l_d_items_x_path = "./div[@class='list-detail']/p[@class='l-d-item']"
                # 相对于 l-d-item 的路径
                item_spans_x_path = "./span"
                l_d_items = list_item_content_element.find_elements(By.XPATH, l_d_items_x_path)
                for l_d_item in l_d_items:
                    item_spans = l_d_item.find_elements(By.XPATH, item_spans_x_path)
                    for item_span in item_spans:
                        item_span_text = item_span.text.strip().replace(" ", "").replace(":", "：")

                        if item_span_text.endswith("获取更多"):
                            item_span_text = item_span_text[:-4]

                        one_available_business_opportunity_record.summary_info.append(item_span_text)
                        item_span_value = item_span_text.split("：")
                        if len(item_span_value) < 2:
                            continue
                        match item_span_value[0]:
                            case "采购单位":
                                one_available_business_opportunity_record.tenderer = item_span_value[1]
                            case "采购单位联系方式":
                                one_available_business_opportunity_record.tenderer_contact = item_span_value[1]
                            case "预算金额":
                                budge_value = base_fun.calculate_amount(item_span_value[1])
                                if budge_value:
                                    one_available_business_opportunity_record.budget = budge_value
                            case "代理机构":
                                one_available_business_opportunity_record.bid_agency = item_span_value[1]
                            case "代理机构联系方式":
                                one_available_business_opportunity_record.agency_contact = item_span_value[1]
                            case "投标截止日期":
                                one_available_business_opportunity_record.deadline = f"{base_fun.string_to_datetime(item_span_value[1]).date()}"
                            case _:
                                print(f"{item_span_text}")

            except Exception as e:
                message = f'单位：{company} \n项目名称：{one_available_business_opportunity_record.title}\n第{page_num}/{total_pages}页，第{record_in_page}条数据，解析详细列表出错! '
                base_fun.to_qw_error(message, exception_info=e)
            finally:
                one_available_business_opportunity_record.summary_info = "    ".join(one_available_business_opportunity_record.summary_info)

            #开始处理当前单位是否满足要求：
            if len(one_available_business_opportunity_record.tenderer)>0:
                if not base_fun.audit_company_name(one_available_business_opportunity_record.tenderer, company, company_info[1], company_info[2]):
                    message = f'单位：{company} 项目名称：{one_available_business_opportunity_record.title}\n第{page_num}/{total_pages}页，第{record_in_page}条数据，采购单位：{one_available_business_opportunity_record.tenderer} 不满足要求! '
                    base_fun.to_qw_error(message)
                    continue
            else:  #没有解析到采购单位，放弃当前记录
                message = f'单位：{company} 项目名称：{one_available_business_opportunity_record.title}\n第{page_num}/{total_pages}页，第{record_in_page}条数据，采购单位为空， 不满足要求! '
                base_fun.to_qw_error(message)
                continue


            ############################################################################################################
            #超链接打开代码；按打开顺序分层级
            #信息查询网页为第一层，
            # 剑鱼链接为第二层（点击：is_hyper_link_work_done, 页面打开：is_hyper_link_opened, 页面关闭：is_hyper_link_closed），
            # 原文链接为第三层（点击：is_original_link_work_done, 页面打开：is_original_link_opened, 页面关闭：is_original_link_closed）

            #0： 参数初始化
            is_current_record_already_in_file = False  #当前记录是否存在
            is_hyper_link_work_done = False
            is_hyper_link_opened = False
            is_hyper_link_closed = False
            original_link_element = None
            is_original_link_work_done = False
            is_original_link_opened = False
            is_original_link_closed = False
            #1：打开剑鱼超链接
            current_window_numer = len(browse_driver.window_handles)
            print(f"在第一层信息查询网页，点击超链接前的窗口数量：{current_window_numer}")
            if title_hyper_link_element is not None: #第一层存在超链接的情况， 尝试打开第二层
                # 获取当前的窗口数量，判断是否生成了新窗口
                is_hyper_link_work_done = base_fun.click_or_move_mouse(browse_driver, title_hyper_link_element,
                                                                      name_of_element="title_hyper_link_element")
            #is_hyper_link_work_done 为 false 由俩个可能性
            # 1）： title_hyper_link_element is None ,
            # 2） title_hyper_link_element is not None ，但点击 title_hyper_link_element 失败。
            if not is_hyper_link_work_done:  #没有链接，title_hyper_link_element 直接jianyu_link ,原始链接，详细信息赋值。
                if title_hyper_link_element is None:
                    one_available_business_opportunity_record.complete_info = f'单位：{company} 第{page_num}/{total_pages}页，第{record_in_page}条数据，未获得项目名称! '
                else: #点击剑鱼超链接失败
                    one_available_business_opportunity_record.complete_info = f'单位：{company} ， 点击鲢鱼超链接失败! '
            else: #is_hyper_link_work_done == True:
                is_hyper_link_opened = base_fun.judge_window_opened(browse_driver, current_window_numer + 1,
                                                                    window_name="剑鱼详情链接")
                if not is_hyper_link_opened:
                    one_available_business_opportunity_record.complete_info = f'单位：{company} ， 打开鲢鱼超链接失败! '

            if not is_hyper_link_opened:
                is_current_record_already_in_file = is_record_exist(df_exist_records,
                                                                    one_available_business_opportunity_record.bidding_type,
                                                                    one_available_business_opportunity_record.title,
                                                                    one_available_business_opportunity_record.tenderer,
                                                                    one_available_business_opportunity_record.release_date,
                                                                    one_available_business_opportunity_record.budget,
                                                                    one_available_business_opportunity_record.jianyu_link)
                #在没有打开第二层窗口的条件下，对于记录已经存在的情况，直接跳转到下一个记录
                if is_current_record_already_in_file:
                    print(f'第{page_num}/{total_pages}页，第{record_in_page}条数据,已经存在记录！\n单位：{company}  项目名称：{one_available_business_opportunity_record.title}\n 直接跳转！')
                    continue

            #2：对于剑鱼链链接被打开的情况：
            if is_hyper_link_opened:
                base_fun.wait_after_operation(cfg.time_after_open_new_window)
                print(f"进入第二层详情页面，当前窗口数量：{len(browse_driver.window_handles)}")
                base_fun.wait_document_ready(browse_driver, 50)
                # 切换到当前最新打开的窗口，-1指的是最后一个标签页，也就是最新的标签页
                browse_driver.switch_to.window(browse_driver.window_handles[-1])
                # 确认页面加载完成
                base_fun.wait_document_ready(browse_driver, 50)
                one_available_business_opportunity_record.jianyu_link = browse_driver.current_url
                #检查记录是否存在
                is_current_record_already_in_file = is_record_exist(df_exist_records,
                                                                    one_available_business_opportunity_record.bidding_type,
                                                                    one_available_business_opportunity_record.title,
                                                                    one_available_business_opportunity_record.tenderer,
                                                                    one_available_business_opportunity_record.release_date,
                                                                    one_available_business_opportunity_record.budget,
                                                                    base_fun.decode_keyword_in_url(one_available_business_opportunity_record.jianyu_link,"link")
                                                                    )
                if is_current_record_already_in_file: #若但钱字符已经存在，关闭剑鱼详情链接，进入下一个记录的循环
                    print(f'第{page_num}/{total_pages}页，第{record_in_page}条数据,已经存在记录！\n单位：{company}  项目名称：{one_available_business_opportunity_record.title}\n 关闭第二层窗口后跳转！')
                    browse_driver.switch_to.window(browse_driver.window_handles[-1])
                    base_fun.close_window_to_one(browse_driver,"剑鱼详情链接")
                    # 关闭剑鱼详情链接后，恢复现场：
                    browse_driver.switch_to.window(browse_driver.window_handles[-1])
                    browse_driver.switch_to.default_content()
                    # 再次进入iframe,查询下一个数据
                    switch_to_iframe_jianyu(browse_driver)
                    continue

                # 等待内容是否加载成功：
                # 2-1）切换到iframe
                switch_to_iframe_jianyu(browse_driver, retry=True)
                # 2-2) 检查需要的元素是否显示， 包括：权限不足，可以看到详情 两种情况
                #新增页面因为异常情况不能打开的情况  @lqh  20250407

                try:
                    element_got = WebDriverWait(browse_driver, 50).until(
                        EC.any_of(
                            EC.visibility_of_element_located((By.XPATH, new_page_info["un_present"])),
                            EC.visibility_of_element_located((By.XPATH, new_page_info["summary"]))
                        )
                    )
                except Exception:
                    print("打开第2层异常！")
                    element_got= None
                #给一次用户手工处理的手段：
                if element_got is None:
                    input("异常，等待用户手工处理， 用户处理结束后，请点击回车继续！")
                    try:
                        switch_to_iframe_jianyu(browse_driver, retry=True)
                        element_got = WebDriverWait(browse_driver, 50).until(
                            EC.any_of(
                                EC.visibility_of_element_located((By.XPATH, new_page_info["un_present"])),
                                EC.visibility_of_element_located((By.XPATH, new_page_info["summary"]))
                            )
                        )
                    except Exception:
                        print("用户手工处理后，依然失败！")
                        element_got = None

                # 2-3) 对页面显示内容进行分析， 判断具体属于：权限不足，可以看到详情
                if element_got is None:
                    class_of_element_got = None
                else:
                    class_of_element_got = element_got.get_attribute("class")
                #开始进行第二层内容给的处理
                if class_of_element_got is None:
                    one_available_business_opportunity_record.complete_info = "页面异常，未获得访问权限！"
                    one_available_business_opportunity_record.original_link = ""
                elif "content-mask-container" in class_of_element_got:  # 当前页面是 权限不足：
                    print("权限不足，无法查看")
                    one_available_business_opportunity_record.complete_info = "权限不足，无法查看"
                    one_available_business_opportunity_record.original_link = ""
                else:   #当前页面可以显示公告详情
                    #2-3-1) 解析公告摘要的table， 仅仅用于查找投标日期
                    is_find_deadline = False  #summary 仅用于寻找投标期限
                    tr_elements = element_got.find_elements(By.XPATH,"./tr")
                    for tr in tr_elements:
                        if is_find_deadline:
                            break
                        td_elements = tr.find_elements(By.XPATH, "./td")
                        index =1
                        while index < len(td_elements):
                            category_name = td_elements[index-1].text.replace(" ","")
                            td_content = td_elements[index].text.replace(" ","")
                            '''
                            if td_content.endswith("查看详情") :
                                td_content= td_content[:-4]
                            if td_content.endswith("更多联系人"):
                                td_content = td_content[:-5]
                            '''
                            if category_name == "报名截止日期":
                                one_available_business_opportunity_record.deadline = f"{base_fun.string_to_datetime(td_content).date()}"
                                is_find_deadline = True
                                break
                            index += 2
                    #2-3-2) 解析公告正文
                    try:
                        base_fun.wait_after_operation(cfg.time_after_scroll)
                        announcement_label_element = WebDriverWait(browse_driver, 20).until(
                            EC.visibility_of_element_located((By.XPATH, new_page_info["announcement_label"])))
                        announcement_label_element.click()
                        one_available_business_opportunity_record.complete_info = base_fun.get_content_by_driver_wait(
                            browse_driver, By.XPATH, new_page_info["announcement"], max_wait_seconds=10)
                    except Exception as e:
                        one_available_business_opportunity_record.complete_info="详情解析异常！"
                        message = "解析公告详情异常"
                        base_fun.to_qw_error(message, exception_info=e, is_debug=True)

                    #2-3-3） 解析原文超链接
                    try:
                        original_link_element = WebDriverWait(browse_driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, new_page_info["original_link"])))
                        is_original_link_element_exception = False
                    except (TimeoutException,Exception ) as e:
                        is_original_link_element_exception = True
                        original_link_element = None
                        message = f'单位：{company} 项目名称：{one_available_business_opportunity_record.title}\n第{page_num}/{total_pages}页，第{record_in_page}条数据，获取原始链接出错! '
                        base_fun.to_qw_error(message, exception_info=e, is_debug=True)

                    #未能获得元素original_link_element，尝试使用find_elements
                    if is_original_link_element_exception :
                        try:
                            original_link_elements = browse_driver.find_elements(By.XPATH, new_page_info["original_link"])
                            if len(original_link_elements)>0:
                                original_link_element = original_link_elements[0]
                            else:
                                original_link_element = None
                        except Exception as e:
                            original_link_element = None

            #3:打开原文超链接
            #因出现超链接一次打开两个页面的特殊情况，特设置变量： window_number_before_original_link_work_done
            window_number_before_original_link_work_done = len(browse_driver.window_handles)
            if is_hyper_link_opened and   not original_link_element is None: #原文超链接不为空，尝试打开，
                current_window_numer = len(browse_driver.window_handles)
                is_original_link_work_done = base_fun.click_or_move_mouse(browse_driver,original_link_element,name_of_element="原文超链接")

                if is_original_link_work_done:
                    is_original_link_opened = base_fun.judge_window_opened(browse_driver, current_window_numer + 1,
                                                                           window_name="原文链接")
                    base_fun.wait_after_operation(cfg.time_after_external_link)

                #3-1 处理原文链接打开的情况
                if is_original_link_opened:
                    print(f"已打开第三层原文链接页面，当前窗口数量：{len(browse_driver.window_handles)}")
                    #强制休息1秒， 用于原文页面加载
                    base_fun.wait_after_operation(cfg.time_after_external_link)
                    # 切换到当前最新打开的窗口，-1指的是最后一个标签页，也就是最新的标签页
                    browse_driver.switch_to.default_content()
                    browse_driver.switch_to.window(browse_driver.window_handles[-1])
                    browse_driver.switch_to.default_content()
                    base_fun.wait_document_ready(browse_driver, 50)
                    one_available_business_opportunity_record.original_link = browse_driver.current_url

            #4 关闭原文链接
            if is_original_link_opened:
                current_window_numer = len(browse_driver.window_handles)
                print(f"准备关闭第三层原文链接页面，当前窗口数量：{current_window_numer}")
                window_index = window_number_before_original_link_work_done - current_window_numer
                while window_index <= -1 and not is_original_link_closed:
                    try :
                        original_link = browse_driver.current_url
                        browse_driver.switch_to.window(browse_driver.window_handles[window_index])
                        browse_driver.close()
                        is_original_link_closed =  base_fun.judge_window_closed(browse_driver, current_window_numer - 1, window_name=f"原文链接[{window_index}]")
                        if is_original_link_closed:
                            one_available_business_opportunity_record.original_link = original_link
                            break
                    except Exception as e:
                        print(f"关闭窗口 index= {window_index}，异常")
                    window_index += 1

                base_fun.wait_after_operation(cfg.time_after_close_window)
                if not is_original_link_closed:
                    print(f"程序关闭原文链接失败，请手动关闭窗口至2个，程序继续执行！")
                    while len(browse_driver.window_handles) >= current_window_numer:
                        base_fun.wait_after_operation(cfg.time_wait_user_operation_between_poll)
                    is_original_link_closed = True
                #关闭原文连接后，恢复现场：
                browse_driver.switch_to.window(browse_driver.window_handles[-1])

            #5 关闭剑鱼详情链接
            if is_hyper_link_opened:
                browse_driver.switch_to.window(browse_driver.window_handles[-1])
                base_fun.close_window_to_one(browse_driver,"剑鱼详情链接")

                # 关闭剑鱼详情链接后，恢复现场：
                browse_driver.switch_to.window(browse_driver.window_handles[-1])
                browse_driver.switch_to.default_content()
                # 再次进入iframe,查询下一个数据
                switch_to_iframe_jianyu(browse_driver)

            # 添加当前记录采集时间
            one_available_business_opportunity_record.collected_date = datetime.now()
            #增加时间判断：
            current_time = datetime.now().time()
            if cfg.auto_AI_always  or cfg.auto_AI_begin < current_time < cfg.auto_AI_end:
                #时间范围满足要求,进行判断
                if AI_analysis_content_in_jianyu(one_available_business_opportunity_record):
                    print(f"项目相关性：{one_available_business_opportunity_record.relevance}")
                    print(f"项目概况：\n{one_available_business_opportunity_record.summary_info}")
                    if "很高" in one_available_business_opportunity_record.relevance:
                        high_relevance_records_amount +=1
            # 成功处理的记录数加1
            success_record_num += 1
            available_records_amount +=1
            # 将获取到的信息添加到数组中
            # 类型(采购意向)bidding_type、标题 title、项目标签 project_tags、相关产品 product_list、预算金额budget、招标单位tenderer、发布时间release_date
            business_opportunity_list.append(one_available_business_opportunity_record.push_data_into_list())
        # end for div in divs 循环，处理每个分页内的记录

        # 如果当前分页等于总的分页数， 停止while循环
        if page_num >= total_pages:
            break
            # 点击下一页
        try:
            # 滑动到页面底端，然后点击下一页
            browse_driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            # 避免爬虫运行太快， 每完成一页检索，滑动到页面底端后，休息，然后点击 下一页
            base_fun.wait_after_operation(cfg.time_after_scroll)
            browse_driver.find_element(By.XPATH, ".//button[@class='btn-next']").click()
            #############################################################################################################
            #每次点击下一页后，重新计算最大页数：
            total_pages = int(search_result_container_element.find_element(By.XPATH, operation_map["total_pages"]).text.strip())
            #############################################################################################################
            print(f'在第{page_num}页 点击下一页，等待进入第 {page_num+1} 页')
            # 等待页面被加载完成，最大等待值20秒
            base_fun.wait_document_ready(browse_driver,20)
            #翻页后等待，让内容加载完成
            #sleep(cfg.time_wait_next_page)

            first_search_result_item_x_path = ".//div[@class='search-result-container']/div"
            WebDriverWait(browse_driver, 20).until(
                EC.visibility_of_element_located((By.XPATH, first_search_result_item_x_path)))
        except Exception as e:
            message = f'单位：{company} ，第{page_num}页，点击下一页失败!'
            base_fun.to_qw_error(message, exception_info=e, is_debug= True)
            break
        page_num += 1
        #end  分页循环

    #总记录数， 有效记录数，高相关性记录数 0
    analysis_statistic_result[0] = total_records
    analysis_statistic_result[1] = available_records_amount
    analysis_statistic_result[2] = high_relevance_records_amount
    return success_record_num


######################################################################################################################
#####页面交互函数定义
# 设置日期
def set_search_date_range(web_driver:webdriver, begin_date:date, end_date:date)->bool:
    jianyu_start_date_x_path = "/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[1]/div[2]/div/div/div[2]/div/div/input[1]"
    jianyu_end_date_x_path = "/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[1]/div[2]/div/div/div[2]/div/div/input[2]"

    #获取日期元素
    jianyu_start_date_element = WebDriverWait(web_driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, jianyu_start_date_x_path)))
    jianyu_end_date_element = WebDriverWait(web_driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, jianyu_end_date_x_path)))
    #点击，调出日历
    jianyu_start_date_element.click()
    # 出现日历框后
    el_picker_panel_left_x_path = ".//div[@class='el-picker-panel__body']/div[@class='el-picker-panel__content el-date-range-picker__content is-left']"
    el_picker_panel_right_x_path = ".//div[@class='el-picker-panel__body']/div[@class='el-picker-panel__content el-date-range-picker__content is-right']"
    date_range_header = "./div[@class='el-date-range-picker__header']"
    # 左侧panel
    el_picker_panel_left_element = WebDriverWait(web_driver, 20).until(
        EC.presence_of_element_located((By.XPATH, el_picker_panel_left_x_path)))
    # 左侧表头
    date_range_header_left = WebDriverWait(el_picker_panel_left_element, 20).until(
        EC.presence_of_element_located((By.XPATH, date_range_header)))
    # 左侧表头文字
    date_range_header_left_text = WebDriverWait(date_range_header_left, 20).until(
        EC.presence_of_element_located((By.XPATH, "./div")))
    # 左侧表头button年，路径相对于：date_range_header_left
    left_header_btn_left_year_x_path = "./button[1]"
    left_header_btn_left_year = WebDriverWait(date_range_header_left, 20).until(
        EC.presence_of_element_located((By.XPATH, left_header_btn_left_year_x_path)))
    # 左侧表头button月，路径相对于：date_range_header_left
    left_header_btn_left_month_x_path = "./button[2]"
    left_header_btn_left_month = WebDriverWait(date_range_header_left, 20).until(
        EC.presence_of_element_located((By.XPATH, left_header_btn_left_month_x_path)))
    # 左侧表头button年，路径相对于：date_range_header_left
    left_header_btn_right_year_x_path = "./button[3]"
    left_header_btn_right_year = WebDriverWait(date_range_header_left, 20).until(
        EC.presence_of_element_located((By.XPATH, left_header_btn_right_year_x_path)))
    # 左侧表头button月，路径相对于：date_range_header_left
    left_header_btn_right_month_x_path = "./button[4]"
    left_header_btn_right_month = WebDriverWait(date_range_header_left, 20).until(
        EC.presence_of_element_located((By.XPATH, left_header_btn_right_month_x_path)))
    # 左侧日期table/tbody ，第一行为周几, 路径相对于：el_picker_panel_left_element
    left_table_tbody_x_path = "./table[@class='el-date-table']/tbody"
    left_table_tbody = WebDriverWait(el_picker_panel_left_element, 20).until(
        EC.presence_of_element_located((By.XPATH, left_table_tbody_x_path)))

    # 右侧panel
    po_picker_panel_right_element = WebDriverWait(web_driver, 20).until(
        EC.presence_of_element_located((By.XPATH, el_picker_panel_right_x_path)))
    # 右侧表头
    date_range_header_right = WebDriverWait(po_picker_panel_right_element, 20).until(
        EC.presence_of_element_located((By.XPATH, date_range_header)))
    # 右侧表头文字
    date_range_header_right_text = WebDriverWait(date_range_header_right, 20).until(
        EC.presence_of_element_located((By.XPATH, "./div")))
    # 右侧表头button年，路径相对于：date_range_header_right
    right_header_btn_left_year_x_path = "./button[1]"
    right_header_btn_left_year = WebDriverWait(date_range_header_right, 20).until(
        EC.presence_of_element_located((By.XPATH, right_header_btn_left_year_x_path)))
    # 右侧表头button月，路径相对于：date_range_header_right
    right_header_btn_left_month_x_path = "./button[2]"
    right_header_btn_left_month = WebDriverWait(date_range_header_right, 20).until(
        EC.presence_of_element_located((By.XPATH, right_header_btn_left_month_x_path)))
    # 右侧表头button年，路径相对于：date_range_header_right
    right_header_btn_right_year_x_path = "./button[3]"
    right_header_btn_right_year = WebDriverWait(date_range_header_right, 20).until(
        EC.presence_of_element_located((By.XPATH, right_header_btn_right_year_x_path)))
    # 右侧表头button月，路径相对于：date_range_header_right
    right_header_btn_right_month_x_path = "./button[4]"
    right_header_btn_right_month = WebDriverWait(date_range_header_right, 20).until(
        EC.presence_of_element_located((By.XPATH, right_header_btn_right_month_x_path)))
    # 左侧日期table/tbody ，第一行为周几, 路径相对于：po_picker_panel_right_element
    right_table_tbody_x_path = "./table[@class='el-date-table']/tbody"
    right_table_tbody = WebDriverWait(po_picker_panel_right_element, 20).until(
        EC.presence_of_element_located((By.XPATH, right_table_tbody_x_path)))

    # 获取年月的正则表达式
    pattern = r"(\d{4})\s*年\s*(\d{1,2})\s*月"
    #读取当前左侧日历的时间
    date_list = [begin_date, end_date]
    #为了简便，所有的日期都放在左侧进行选择
    for index in [0,1]:
        target_date =  date_list[index]
        matches_left = re.search(pattern, date_range_header_left_text.text)
        matches_right=  re.search(pattern, date_range_header_right_text.text)
        # 遍历匹配结果
        if not matches_left:
            print(f"无法解析日历的年月: { date_range_header_left_text.text}")
            return False
        if not matches_right:
            print(f"无法解析日历的年月: { date_range_header_right_text.text}")
            return False
        #日期匹配成功
        calendar_left_year = int(matches_left.group(1))
        calendar_left_month = int(matches_left.group(2))
        calendar_right_year = int(matches_right.group(1))
        calendar_right_month = int(matches_right.group(2))
        #开始选择选择左侧还是右侧。
        calendar_right_first_day =date(calendar_right_year,calendar_right_month,1)
        if target_date < calendar_right_first_day:
            #选择左侧 panel
            table_body = left_table_tbody
            year_sub_btn = left_header_btn_left_year
            year_add_btn = left_header_btn_right_year
            month_sub_btn= left_header_btn_left_month
            month_add_btn = left_header_btn_right_month
            calendar_year = calendar_left_year
            calendar_month = calendar_left_month
        else:
            # 选择右侧 panel
            table_body = right_table_tbody
            year_sub_btn = right_header_btn_left_year
            year_add_btn = right_header_btn_right_year
            month_sub_btn = right_header_btn_left_month
            month_add_btn = right_header_btn_right_month
            calendar_year = calendar_right_year
            calendar_month = calendar_right_month


        #比较当前日历日期的年月目标日期的年月

        delta_month =  target_date.month-calendar_month
        delta_year = target_date.year-calendar_year
        if delta_year >0:
            year_element = year_add_btn
        else:
            year_element = year_sub_btn
            delta_year *= -1

        if delta_month>0:
            month_element = month_add_btn
        else:
            month_element = month_sub_btn
            delta_month *= -1

        while delta_year > 0:
            delta_year -= 1
            year_element.click()
            sleep(0.1)

        while delta_month >0:
            delta_month -= 1
            month_element.click()
            sleep(0.1)

        #获取左侧日的第一单元格：
        first_day_element = table_body.find_element(By.XPATH,"./tr[2]/td[1]")
        day_no = int(first_day_element.text.strip())
        month_no = target_date.month
        year_no = target_date.year
        if "prev-month" in first_day_element.get_attribute("class"):
            month_no -= 1
            if month_no == 0 :
                year_no -= 1
                month_no = 12
        try :
            first_space_day = date(year_no, month_no, day_no)
        except Exception:
            print(f"计算日期：{target_date}所在月的第一个格子日期出错： {year_no}年{month_no}月{day_no}日")
            return False

        day_space_no = (target_date - first_space_day).days +1
        row_no = int(day_space_no / 7) + 2
        col_no = day_space_no % 7
        if col_no == 0:
            col_no = 7
            row_no -= 1
        day_element = WebDriverWait(table_body,20).until(EC.element_to_be_clickable((By.XPATH,f"./tr[{row_no}]/td[{col_no}]")))
        day_element.click()
        sleep(0.2)
    return True
