# 存放一些公用函数
from time import sleep

import requests

from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By


# 定义一个函数，通过test账号对API进行调用
def get_response_data(url, query):
    username = 'C5338206'
    password = 'aTns7IzJiI5}Q'
    headers = {"Accept": "application/json"}
    params = {"sysparm_limit": "100", "sysparm_query": query}
    response = requests.get(url, params=params, headers=headers, auth=(username, password))

    if response.status_code == 200:
        # print("Data is successfully got！")
        return response.json()
    else:
        print("Request failed：", response.status_code)
        return


# 定义一个函数，根据current_update_set生成下一个编号
def generate_next_number(current_number):
    prefix = current_number[:-2]  # 提取前缀部分
    suffix = current_number[-2:]  # 提取后缀部分
    # 将后缀部分转换为整数并递增
    next_suffix = int(suffix) + 1
    # 根据需要进行格式化，确保后缀部分有两位数
    next_suffix_formatted = str(next_suffix).zfill(2)
    # 组合前缀和递增后的后缀生成下一个编号
    next_number = prefix + next_suffix_formatted
    return next_number


# 定义一个函数，在dev中进行搜索
def get_search_result_in_dev(keyword, driver):
    # 通过aria-label属性定位input元素
    search_box = driver.find_element(By.CSS_SELECTOR, '[aria-label="Search"]')
    # 找到搜索输入框并输入关键字
    search_box.send_keys(keyword)
    # 模拟按下回车键进行搜索
    search_box.send_keys(Keys.RETURN)
    # 停留一段时间等待搜索结果加载
    driver.implicitly_wait(1)  # 等待1秒钟
    # 找到具有指定 ID 的表格
    table = driver.find_element(By.ID, 'sys_update_set_table')
    # 提取数据行
    result_rows = []
    result = table.find_elements(By.TAG_NAME, 'tr')[2:]  # 跳过第一、二行，因为它是表头
    for row in result:
        row_cells = row.find_elements(By.TAG_NAME, 'td')
        row_data = []
        for cell in row_cells:
            row_data.append(cell.text.strip())
        result_rows.append(row_data)
    return result_rows


# 定义一个函数，用于打开新的edge页面并定位至iframe
def open_webpage(url_choice):
    # 创建 EdgeOptions 对象
    edge_options = webdriver.EdgeOptions()
    # 添加选项，以保持浏览器常开
    edge_options.add_experimental_option("detach", True)
    # 使用Edge浏览器打开Dev
    driver = webdriver.Edge(options=edge_options)
    driver.maximize_window()  # 最大化窗口

    url = ""
    if url_choice == "ppm":
        url = ("https://sapppm.service-now.com/now/nav/ui/classic/params/target/rm_story_list.do"
               "%3Fsysparm_clear_stack%3Dtrue")
    elif url_choice == "nonprod_ppm_release":
        url = ("https://sapnonprod1.service-now.com/now/nav/ui/classic/params/target/rm_release_scrum_list.do"
               "%3Fsysparm_query%3D%26sysparm_first_row%3D1%26sysparm_view%3D")
    elif url_choice == "nonprod_ppm_story":
        url = ("https://sapnonprod1.service-now.com/now/nav/ui/classic/params/target/sn_safe_story_list.do"
               "%3Fsysparm_clear_stack%3Dtrue")
    elif url_choice == "dev":
        url = ("https://sapsandbox.service-now.com/now/nav/ui/classic/params/target/sys_update_set_list.do"
               "%3Fsysparm_userpref_module%3D50047c06c0a8016c0135a14cebc8191b")
    elif url_choice == "test":
        url = ("https://sapsandbox.service-now.com/now/nav/ui/classic/params/target/sys_remote_update_set_list.do"
               "%3Fsysparm_userpref_module%3Dbf1184a10a0a0b5000d8f781992a9b5e%26sysparm_fixed_query%3Dsys_class_name"
               "%3Dsys_remote_update_set")
    driver.get(url)

    if url_choice == "nonprod_ppm_release" or url_choice == "nonprod_ppm_story":
        # 设置最大等待时间
        sleep(5)
        username_input = driver.find_element(By.ID, 'user_name')
        passwd_input = driver.find_element(By.ID, 'user_password')
        username_input.send_keys("I746326")
        passwd_input.send_keys("20030929shyDUCK@")
        # 模拟按下回车键进行搜索
        passwd_input.send_keys(Keys.RETURN)

    # 设置最大等待时间
    sleep(9)
    # 获取第一层 Shadow Root
    shadow_root_1 = driver.execute_script(
        "return document.querySelector('macroponent-f51912f4c700201072b211d4d8c26010').shadowRoot")
    # 查找div元素
    div_element = shadow_root_1.find_element(By.CSS_SELECTOR, "div")
    # 查找sn-canvas-appshell-root元素
    sn_canvas_appshell_root = div_element.find_element(By.TAG_NAME, "sn-canvas-appshell-root")
    # 查找sn-canvas-appshell-layout元素
    sn_canvas_appshell_layout = sn_canvas_appshell_root.find_element(By.TAG_NAME, "sn-canvas-appshell-layout")
    sn_polaris_layout = sn_canvas_appshell_layout.find_element(By.TAG_NAME, "sn-polaris-layout")
    # 在sn_polaris_layout中查找iframe元素
    iframe = sn_polaris_layout.find_element(By.TAG_NAME, "iframe")
    driver.switch_to.frame(iframe)

    return driver
