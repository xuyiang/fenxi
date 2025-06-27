# coding:utf-8

from time import sleep
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.support.ui import  WebDriverWait
from selenium.webdriver.support import  expected_conditions as EC
from selenium.common import TimeoutException
from datetime import  datetime
import time


import config as cfg
import base_function_jianyu as base_jianyu
import base_function as base_fun
from base_function_jianyu import  Struct_business_opportunity

def do_jianyu_crawler(main_web_driver:webdriver):
    """
    这个页面会进入iframe, 程序执行结束时，需要退出 iframe
    :param main_web_driver:
    :return:
    """
    # 打开网页
    main_web_driver.get('https://www.jianyu360.cn/')

    # 调整显示比例
    base_fun.set_document_zoom_scale(main_web_driver, 1, 0.5)

    # 登录系统，由于存在弹窗等可能， 需要进行异常防止
    for try_login_counter in [1]: #:
        # 按以下三个步骤处理1、自动登录，成功，退出。 2、 若失败，尝试按预案关闭弹窗 后登录，成功，退出。3、提示人工关闭阻碍登录的元素后，再登录
        try:
            base_jianyu.login(main_web_driver)
            break
        except Exception :
            print("登录页面失败， 尝试按照预定方法关闭弹窗等异常因素")
        try :  #尝试按照预定方法关闭弹窗等异常因素
            WebDriverWait(main_web_driver, 5).until(EC.presence_of_element_located((By.ID, "close2x"))).click()
        except Exception:
            pass
        try:
            base_jianyu.login(main_web_driver)
            break
        except Exception :
            input("请手动关闭阻碍登录的弹窗后，输入 Enter， 再进行登录")
            base_jianyu.login(main_web_driver)
    base_fun.wait_document_ready(main_web_driver, 20)  #确认页面加载完成

    # 调整显示比例
    base_fun.set_document_zoom_scale(main_web_driver, 1, 0.5)
    #登录后。尝试等关闭弹出的消失提示框
    try:
        #定位： shadow_host
        div_shadow_host_id = "__qiankun_microapp_wrapper_for_big_member_sub_app__"
        div_shadow_host =WebDriverWait(main_web_driver,5).until(EC.presence_of_element_located((By.ID, div_shadow_host_id)))
        # 获取 Shadow Root
        shadow_root = driver.execute_script("return arguments[0].shadowRoot", div_shadow_host)
        #找到close_button
        close_button_class_name =  ".el-icon-circle-close"
        WebDriverWait(shadow_root, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, close_button_class_name))).click()
    except Exception:
        input("登录后，请检查是否有需要关闭的弹窗界面。确认后，输入 Enter，程序继续执行")



    # 登录后，页面进入工作台
    workspace_url = "https://www.jianyu360.cn/page_workDesktop/work-bench/app/big/workspace/dashboard"
    main_web_driver.get(workspace_url)
    base_fun.wait_document_ready(main_web_driver, 20)

    # 进入商机页面
    business_opportunity_x_path = "/html/body/section/section/aside/div/div[1]/div[1]/div[3]/span"
    business_opportunity_element = WebDriverWait(main_web_driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, business_opportunity_x_path)))
    business_opportunity_element.click()
    # 等待商机页面加载完成
    base_fun.wait_document_ready(main_web_driver, 20)



    print("当前进程：剑鱼招标网商机查询")

    #进入iframe
    #iframe_x_path = "/html/body/section/section/main/div/iframe"
    #等待 iframe 可见
    #iframe_element = WebDriverWait(main_web_driver,20).until(EC.presence_of_element_located((By.XPATH, iframe_x_path)))
    #main_web_driver.switch_to.frame(iframe_element)

    base_jianyu.switch_to_iframe_jianyu(main_web_driver,retry=True)



    #根据历史查询情况，确定本次查询参数。
    # 如上次查询未全部结束，last_queried_companies  为上次已经查询并记录的公司列表。
    additional_info = {}  # 用于获取本次已经完成查询公司所对应的总记录数、有效记录数、相关记录数的统计之和
    save_record_file_name, date_range_begin, date_range_end, last_queried_companies, is_an_new_query = base_fun.get_statistic_history_information(
        cfg.filename_history_query_log, "business_opportunity_statistic", release_date_type="date",
        additional_info=additional_info)
    if date_range_begin > date_range_end:
        print(f'本次查询时间范围：{date_range_begin}---{date_range_end}')
        print("时间范围无效，退出程序执行！")
        exit(cfg.ErrorCode_invalid_query_time_range)

    # 选择采购意向参数
    # 选择发布时间，是单选框部分，
    release_date_range_operations = {
                "last7days": '/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[1]/div[2]/div/div/div[1]/div[1]',
                "last30day": '/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[1]/div[2]/div/div/div[1]/div[2]',
                "last1year": '/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[1]/div[2]/div/div/div[1]/div[3]',
                "last3year": '/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[1]/div[2]/div/div/div[1]/div[4]',
                "last5year": '/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[1]/div[2]/div/div/div[1]/div[5]',
                "start_date":'/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[1]/div[2]/div/div/div[2]/div/div/input[1]',
                "end_date": '/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[1]/div[2]/div/div/div[2]/div/div/input[2]'
    }

    # 根据本次查询范围，设置界面中的查询日期范围
    try:
        base_jianyu.set_search_date_range(main_web_driver,date_range_begin.date(), date_range_end.date())
    except Exception:
        input("请手动检查页面是否存在异常，修正后，点击 Enter 继续！")


    search_rang_operations = {
                # 至少需要设置一个范围条件，故将采购单位提前处置
                "caigoudanwei": ["/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[2]/div[2]/div/div/div/div[5]/span[2]","/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[2]/div[2]/div/div/div/div[5]/span[1]"],
                "biaoti":["/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[2]/div[2]/div/div/div/div[1]/span[2]","/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[2]/div[2]/div/div/div/div[1]/span[1]"],
                "zhengwen":["/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[2]/div[2]/div/div/div/div[2]/span[2]","/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[2]/div[2]/div/div/div/div[2]/span[1]"],
                "fujian":["/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[2]/div[2]/div/div/div/div[3]/span[2]","/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[2]/div[2]/div/div/div/div[3]/span[1]"],
                "xiangmumingcheng":["/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[2]/div[2]/div/div/div/div[4]/span[2]","/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[2]/div[2]/div/div/div/div[4]/span[1]"],
                "zhongbiaoqiye":["/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[2]/div[2]/div/div/div/div[6]/span[2]","/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[2]/div[2]/div/div/div/div[6]/span[1]"],
                "zhaobiaodailijigou":["/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[2]/div[2]/div/div/div/div[7]/span[2]","/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[2]/div[2]/div/div/div/div[7]/span[1]"]
    }

    for onekey in search_rang_operations.keys():
        one_element_sibling = WebDriverWait(main_web_driver,20).until(EC.element_to_be_clickable((By.XPATH, search_rang_operations[onekey][1])))
        sibling_class = one_element_sibling.get_attribute('class').lower()
        one_element_search_range = WebDriverWait(main_web_driver,20).until(EC.element_to_be_clickable((By.XPATH, search_rang_operations[onekey][0])))
        if onekey == "caigoudanwei":  # 必须选择
            if not "checked" in sibling_class:
                one_element_search_range.click()
        else:
            if "checked" in sibling_class:
                one_element_search_range.click()
        base_fun.wait_after_operation(cfg.time_after_click)

    #信息类型
    bid_information_type_operations = {
        "quanbu":"/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[3]/div[2]/div/div[2]/div/div/div[1]/button",
        "nijianxiangmu":"/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[3]/div[2]/div/div[2]/div/div/div[2]/button",
        "caigouyixiang":"/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[3]/div[2]/div/div[2]/div/div/div[3]/button",
        "zhaobiaoyugao":"/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[3]/div[2]/div/div[2]/div/div/div[4]/div/div[1]/div/span",
        "zhaobiaogonggao":"/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[3]/div[2]/div/div[2]/div/div/div[5]/div/div[1]/div/span",
        "zhaobiaojieguo":"/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[3]/div[2]/div/div[2]/div/div/div[6]/div/div[1]/div/span",
        "zhaobiaoxinyongxinxi":"/html/body/div[4]/div/div/div/div/div[2]/div/div[2]/div[1]/div/div[3]/div[2]/div/div[2]/div/div/div[7]/div/div[1]/div/span"
    }
    quanbu_element =  main_web_driver.find_element(By.XPATH, bid_information_type_operations["quanbu"])
    if  not "active" in quanbu_element.get_attribute("class").lower():
        quanbu_element.click()
        base_fun.wait_after_operation(cfg.time_after_click)

    zhaobiaoyugao_element = main_web_driver.find_element(By.XPATH, bid_information_type_operations["zhaobiaoyugao"])
    while (not "highlight-text" in zhaobiaoyugao_element.get_attribute("class").lower()):
        zhaobiaoyugao_element.click()
        base_fun.wait_after_operation(cfg.time_after_click)

    zhaobiaogonggao_element = main_web_driver.find_element(By.XPATH, bid_information_type_operations["zhaobiaogonggao"])
    while (not "highlight-text" in zhaobiaogonggao_element.get_attribute("class").lower()):
        zhaobiaogonggao_element.click()
        base_fun.wait_after_operation(cfg.time_after_click)

    # 移动鼠标，便于 前面"zhaobiaoyugao"，"zhaobiaogonggao"下拉框消失
    main_web_driver.find_element(By.XPATH, bid_information_type_operations["caigouyixiang"]).click()
    base_fun.wait_after_operation(cfg.time_after_click)
    #再次点击，取消采购意向
    main_web_driver.find_element(By.XPATH, bid_information_type_operations["caigouyixiang"]).click()


    #使用详细列表进行查询
    detail_list_x_path = "/html/body/div[4]/div/div/div/div/div[4]/div[1]/div[1]/div[2]/div[1]/div[2]/span"
    detail_list_element = main_web_driver.find_element(By.XPATH, detail_list_x_path)
    detail_list_element.click()

    input("请确认查询参数后，输入Enter，程序继续执行")
    #用户确认后， 把页面显示比例缩小，便于进行页面操作。
    base_fun.set_document_zoom_scale(main_web_driver, 1, 0.5)

    # 检查用户是否点选设置了时间段，若设置了需要进行检查
    # 提取用户提供的起始时间和结束时间
    date_range_x_path = "/html/body/div[4]/div/div/div/div/div[2]/div/div[1]/div/div/section/div[1]/div/ul/li[1]/span"
    date_range_text = main_web_driver.find_element(By.XPATH, date_range_x_path).text
    result = base_fun.resolve_date_range(date_range_text)
    if result:
        user_select_query_date_range_begin= result[0]
        user_select_query_date_range_end = result[1]
        date_range_begin_date = f'{date_range_begin.date()}'.replace("-", "/")
        date_range_end_date = f'{date_range_end.date()}'.replace("-", "/")
        while date_range_begin_date != user_select_query_date_range_begin or date_range_end_date !=user_select_query_date_range_end :
            input(f"提高查询效率，需选择 开始日期：{date_range_begin_date}， 结束日期：{date_range_end_date}。选择结束后，回车 ")
            date_range_text = main_web_driver.find_element(By.XPATH, date_range_x_path).text
            result = base_fun.resolve_date_range(date_range_text)
            if result:
                user_select_query_date_range_begin = result[0]
                user_select_query_date_range_end = result[1]
                print(f"用户选择 开始日期：{date_range_begin_date}， 结束日期：{date_range_end_date}。")
            else:
                print("用户放弃使用日期范围，继续！")
                break

    # 初始化待查询公司信息
    # 1 、全量查询公司列表
    company_lis = base_fun.get_company_list(cfg.filename_companies_to_be_queried)

    base_fun.initialize_record_file(excel_filename=save_record_file_name, is_new_query_file=is_an_new_query,
                                    column_list=Struct_business_opportunity().columns_name)
    begin_query_operation_no = -3  # 默认为是继续上次未完成的查询
    begin_query_expression = f"继续查询：{save_record_file_name}"
    if is_an_new_query:
        begin_query_operation_no = -2  # 表示本次是新查询，开始新查询状态下， 使用文件名作为 公司名称字段。
        begin_query_expression = save_record_file_name
    if not base_fun.append_record_list_into_file(cfg.filename_history_query_log, [
        [begin_query_expression, "",begin_query_operation_no,"", datetime.now(), date_range_begin, date_range_end]],
                                          sheet_name="business_opportunity_statistic"):
        #添加历史记录出错，需要告警退出。
        print(f"向文件：{cfg.filename_history_query_log} 添加记录失败，退出程序执行！")
        exit(cfg.ErrorCode_read_write_file)

    # 开始按照待查公司名称进行商机信息查询
    # 定义一个变量，记录所有查询的采购意向公司累计 "成功" 处理的记录数
    total_success_record_num = additional_info["有效记录数"]
    # 定一个空数组，存储获取到的采购意向信息，然后这个数组会放到Pandas中进行处理。
    # 页面解析的采购意向相关信息
    available_record_list = []

    # 获取本次检索范围对应的已有记录
    df_exist_records_in_file = base_fun.get_dataframe_from_excel_file_according_date(save_record_file_name,["网站链接","发布日期","类型","标题","采购单位","预算金额(万)"], 1,date_range_begin.date(), date_range_end.date())
    #将预算金额(万)转为float
    if df_exist_records_in_file is not None:
        base_fun.transform_column_to_float(df_exist_records_in_file,"预算金额(万)")
        base_fun.extract_column_by_re(df_exist_records_in_file,"网站链接",r'"(http(s)?://(.*?))"',0)
    # 所有需要查询采购意向信息的公司
    company_index =0
    total_company_amount = len(company_lis)
    for row_current_company in company_lis:
        company_index += 1
        company_name = row_current_company[0].strip()
        # 如果已经执行过的表就不处理了
        if company_name in last_queried_companies:
            # 已经执行过的表不进行处理，跳过本次循环
            continue

        # 使用当前 company_name 进行查询
        print(f'当前查询{company_name}的商机信息。总体进度：{company_index}/{total_company_amount}')
        # 首先将页面上拉到顶端，然后定位输入框
        windows = main_web_driver.window_handles
        # 切换到当前最新打开的窗口，-1指的是最后一个标签页，也就是最新的标签页
        main_web_driver.switch_to.window(windows[-1])

        # 如页面未完全加载， 循环运行代码，需要一定的时间让网页加载。故在此等待，然后再进行页面滑动
        main_web_driver.execute_script("window.scrollTo(0, 0);")
        sleep(cfg.time_after_scroll)

        #main_web_driver.switch_to.frame(iframe_element)
        base_jianyu.switch_to_iframe_jianyu(main_web_driver)
        # 记录可用的输入源
        query_input_x_path = "/html/body/div[4]/div/div/div/div/div[1]/div/div/div[2]/div[2]/div/div[1]/div[1]/input"

        #query_input_element  = main_web_driver.find_element(By.XPATH,query_input_x_path)
        query_input_element = WebDriverWait(main_web_driver, 50).until(EC.visibility_of_element_located((By.XPATH,query_input_x_path)))
        # 先清空输入框，保证输入公司名之前输入框中是空的
        query_input_element.clear()
        # 将公司名称写入到输入框中
        query_input_element.send_keys(company_name)
        base_fun.wait_after_operation(cfg.time_after_input)
        # 点击提交,
        # 等待 元素出现后,并且可点击，后再提交
        query_submit_x_path = "/html/body/div[4]/div/div/div/div/div[1]/div/div/div[2]/div[2]/div/div[1]/div[1]/div"
        try:
            query_submit_element = WebDriverWait(main_web_driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, query_submit_x_path)))
            query_submit_element.click()
        except TimeoutException:
            print("等待输入可提交超时，退出程序")
            exit(cfg.ErrorCode_timeout_to_wait_submit)
        #提交后，等待一小会时间。
        base_fun.wait_after_operation(cfg.time_after_submit)

        # 等待页面被加载完成，最大等待值20秒
        base_fun.wait_document_ready(main_web_driver, 20)

        #判断查询结果容器是否加载完成
        start_time = time.time()
        search_result_container_x_path = "/html/body/div[4]/div/div/div/div/div[4]/div[1]/div[2]"
        try :
            search_result_container_element = WebDriverWait(main_web_driver,20).until(EC.visibility_of_element_located((By.XPATH, search_result_container_x_path)))
        except TimeoutException :
            print("已等待： {:.2f}秒，等待输出结果超时，退出程序".format(time.time()-start_time))
            exit(cfg.ErrorCode_timeout_to_wait_submit)

        #开始等待详细加载，
        in_loading_data_x_path = './div[@class="empty-container mtb60"]'
        try_index = 0
        while True:
            try :
                in_loading_data_elements = search_result_container_element.find_elements(By.XPATH, in_loading_data_x_path)
                if len(in_loading_data_elements) == 0:
                    break
                try_index+= 1
                if try_index == 1:
                    print("数据加载中.", end='')
                elif (try_index % 60) ==1:
                    print("\n数据加载中.", end='')
                else:
                    print(".", end='')
                sleep(0.5)
            except Exception as e:
                print("等待加载结束异常！")
                break

        #等待空元素和有查询结果脚标元素出现
        empty_container_x_path = './/div[@class="empty-container"]'
        search_result_foot_container_x_path = './/div[contains(@class,"search-result-footer-container")]'
        try:
            search_result_empty_or_foot =  WebDriverWait(search_result_container_element,20).until(
                EC.any_of(
                    EC.visibility_of_element_located((By.XPATH,empty_container_x_path)),
                    EC.visibility_of_element_located((By.XPATH, search_result_foot_container_x_path))
                )
            )
        except TimeoutException:
            mess = f'\n{company_name} 加载查询结果超时，退出程序！'
            base_fun.to_qw_error(mess)
            exit(cfg.ErrorCode_timeout_to_wait_submit)
        print("\n输出显示，总计用时：{:.2f}秒。".format(time.time()-start_time))

        #如果结果是空元素出现
        if search_result_empty_or_foot.get_attribute("class") == "empty-container" :
            base_fun.append_record_list_into_file(cfg.filename_history_query_log,[[company_name,0, 0, 0, datetime.now()]], sheet_name="business_opportunity_statistic")
            print(f'{company_name} 查询结果为空，查询下一家公司！')
            continue



        # 一个被查询采购意向公司的全部商机信息
        previous_success_record_num = total_success_record_num
        analysis_statistic_result = []
        total_success_record_num = base_jianyu.analysis_business_opportunity(main_web_driver, row_current_company,df_exist_records_in_file, available_record_list,
                                                                 total_success_record_num,
                                                                 datetime_begin=date_range_begin,
                                                                 datetime_end=date_range_end, analysis_statistic_result= analysis_statistic_result)
        additional_info["有效记录数"]= total_success_record_num  #记录本轮查询总的有效记录数之和。
        additional_info["总记录数"] += analysis_statistic_result[0]
        additional_info["相关记录数"] += analysis_statistic_result[2]
        print("*" * 50)
        print(f'本轮查询，累计查询记录：{additional_info["总记录数"]} 条； 累计有效记录：{additional_info["有效记录数"]}条；累计高相关性记录：{additional_info["相关记录数"]}条。')
        print("*" * 50)
        # 完成一家公司，将记录写入record文件中,记录个数信息写入日志文件
        available_records_num = total_success_record_num - previous_success_record_num
        if available_records_num > 0:
            base_fun.append_record_list_into_file(save_record_file_name, available_record_list,column_list=Struct_business_opportunity().columns_name)
            available_record_list.clear()
        base_fun.append_record_list_into_file(cfg.filename_history_query_log,[[company_name, analysis_statistic_result[0],available_records_num,analysis_statistic_result[2], datetime.now()]],  sheet_name="business_opportunity_statistic")

    # 所有需要查询招标信息的公司循环结束
    base_fun.append_record_list_into_file(cfg.filename_history_query_log, [
        [save_record_file_name, additional_info["总记录数"],-1, f'{additional_info["有效记录数"]}/{additional_info["相关记录数"]}',datetime.now(), date_range_begin, date_range_end]], sheet_name="business_opportunity_statistic")
    if additional_info["总记录数"] >0: #更新统计
        base_fun.update_statistic_data(cfg.filename_history_query_log, [additional_info["总记录数"], additional_info["有效记录数"], additional_info["相关记录数"]],sheet_name="business_opportunity_statistic")
    if total_success_record_num > 0:
        print(
            f'本次共查询{additional_info["总记录数"]}条记录，筛选出{total_success_record_num}条名单客户的商机信息，其中高相关性记录{additional_info["相关记录数"]}条。\n查询结果见文件：{save_record_file_name}')
    else:
        print(f'本次共查询{additional_info["总记录数"]}条记录，未筛选出名单客户商机信息。')
    #退出函数前，驱动器恢复到主文档
    main_web_driver.switch_to.default_content()

### 程序运行执行部分
if __name__ =="__main__":

    user_select = input("选择您的浏览器，直接 【回车】： Chrome, 【其它键】： Edge ")
    if user_select=="" :
        driver = base_fun.initial_web_driver("chrome")
    else:
        driver = base_fun.initial_web_driver("edge")

    #为了显示
    # 全屏化窗口
    #driver.maximize_window()


    do_jianyu_crawler(driver)


    driver.close()
    # 结束任务
    driver.quit()
