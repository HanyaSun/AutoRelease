import re
from time import sleep

import pandas as pd
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

from public_func import get_response_data, generate_next_number, get_search_result_in_dev, open_webpage


class AutoRelease:
    current_release_number = None
    C_TEST_Update_Set = {
        "name": "",
        "href": "",
    }
    ppm_story_number = None
    dev_story_number = None
    ppm_driver = None
    dev_driver = None
    test_driver = None

    def __init__(self, current_release_number):
        self.current_release_number = current_release_number  # 保存该release的number，例如'RLSE0011779'
        self.C_TEST_Update_Set["name"] = self.copy_update_set()  # 获取Current Update Set
        # self.ppm_driver = None  # open_webpage("nonprod_ppm")
        # self.dev_driver = None  # open_webpage("dev")
        # self.test_driver = None  # open_webpage("test")

    def main_procedure(self, ppm_query_story):
        if self.ppm_story_get_record(ppm_query_story):
            update_set_state = self.search_in_Dev()
            if update_set_state:
                self.check_state(update_set_state)
                self.open_current_update_set()

                self.compare()
                # self.filter()
                # self.check_stories_deployed()
                # self.set_complete_and_export()
                # self.open_test()

        # else:
        return

    def ppm_story_get_record(self, ppm_query_story):
        # ppm_query_story = "release=b0fe0d66dbe33910bfa850d3f3961900^state=13"

        # Step 1. Open Release Mgmt. Dashboard
        # 获取story列表
        ppm_url_story = "https://sapppm.service-now.com/api/now/table/sn_safe_story"
        data_story = get_response_data(ppm_url_story, ppm_query_story)

        # Step 2. Check if current release has story to be deployed
        if data_story['result']:
            # 有需要被deploy的story
            df_story_return = pd.DataFrame(data_story['result'])
            # df_story.to_csv('data_story.csv', index=False)  # 保存为CSV文件
            # df_story.to_excel('data_story.xlsx', index=False)  # 保存为Excel文件
            count = len(df_story_return)

            # Step 3. Note down record number of the stories in current release
            # 存储task_effective_number列所有的数据,以便在Step12进行对比
            self.ppm_story_number = df_story_return['task_effective_number'].tolist()

            # 需修改 为了测试造的数据
            self.ppm_story_number = ['SFSTRY0053510',
                                     'SFSTRY0053512',
                                     'SFSTRY0053513'
                                     ]
            count = 3
            print("Current release has", count, "stories to be deployed:", self.ppm_story_number)
            return True
        else:
            print("Current release has no story to be deployed!")
            return False

    def copy_update_set(self):
        # Step 4. Copy string field: Current Release Update Set from PPM
        ppm_url_release = "https://sapppm.service-now.com/api/now/table/rm_release_scrum"
        ppm_query_release = "u_test_deployment=true"
        data_release = get_response_data(ppm_url_release, ppm_query_release)

        if data_release is not None:
            # 请求成功，进行下一步操作
            df_release_return = pd.DataFrame(data_release['result'])

            # 存储number为当前release的u_current_parent_update_set列数据
            UCPUS = (df_release_return[df_release_return['number'] == self.current_release_number]
                     ['u_current_parent_update_set'].tolist())
            # 获取url以便访问链接得到u_name
            url = UCPUS[0]['link']
            params = {"sysparm_limit": "100"}

            data = get_response_data(url, params)
            if data is not None:
                # 提取name标签
                self.C_TEST_Update_Set["name"] = data['result']['u_name']

                # 需修改 下面这一句是为了测试，直接赋值常量
                self.C_TEST_Update_Set["name"] = "RLSE0011779_HCSM Release 24_2_TEST_V180"

                print("Current Update Set:", self.C_TEST_Update_Set["name"])
                return self.C_TEST_Update_Set["name"]
            else:
                print("Failed to get Current Update Set's name.")
                return
        else:
            print("Failed to get releases.")
            return

    def search_in_Dev(self):
        # Step 5. Open System_Update_Sets
        self.dev_driver = open_webpage("dev")

        # Step 6. Input Current Release Update Set in search box & enter down
        # Step 7. Display results
        result = get_search_result_in_dev(self.C_TEST_Update_Set["name"], self.dev_driver)
        if result:
            # 提取第一行数据列并保存第五列的元素到变量
            state = result[0][4]  # 获取结果列表的第五个元素
            print("Current Update Set's state: ", state)
            return state
        else:
            # 一般不会出现的情况
            print("Dev has no record of Update Set:", self.C_TEST_Update_Set["name"])
            return False

    def check_state(self, state):
        # Step 8.Check if the state is complete or in progress
        if state == "Complete":
            # Step 8.1. Display Alert Message to notify release manager that current release has been completed.
            print("Update set", self.C_TEST_Update_Set["name"], "has been completed.")
            # 形成下一版本的Update Set
            next_C_Test_Update_Set = generate_next_number(self.C_TEST_Update_Set["name"])
            # Step 8.2. Input next Update Set in search box & enter down
            result = get_search_result_in_dev(next_C_Test_Update_Set, self.dev_driver)

            if len(result) == 0:
                # 如果下一版本update set不存在
                # Step 8.2.1. Click New button
                new_btn = self.dev_driver.find_element(By.ID, 'sysverb_new')
                new_btn.click()
                name_box = self.dev_driver.find_element(By.ID, 'sys_update_set.name')
                name_box.clear()  # 清除输入框内容

                # Step 8.2.2. Input next version of Update Set in Name box
                name_box.send_keys(next_C_Test_Update_Set)

                # Step 8.2.3.Click Submit
                submit_btn = self.dev_driver.find_element(By.ID, 'sysverb_insert_bottom')
                submit_btn.click()
                print("New version of Update Set", next_C_Test_Update_Set, "has been created successfully!")

            # 更新当前update_set
            self.C_TEST_Update_Set["name"] = next_C_Test_Update_Set
            # 更新一次result
            result = get_search_result_in_dev(next_C_Test_Update_Set, self.dev_driver)
            # 提取第一行数据列并保存第五列的元素到变量
            next_state = result[0][4]  # 获取结果列表的第五个元素
            print("New Update Set", self.C_TEST_Update_Set["name"], "state: ", next_state)

            # 在PPM里面更新
            # Step 8.3.1 Open rm_release_scrum.list
            self.ppm_driver = open_webpage("nonprod_ppm_release")
            # Step 8.3.2 Input Release number in search box
            search_box = self.ppm_driver.find_element(By.CSS_SELECTOR, '[aria-label="Search"]')
            search_box.send_keys(self.current_release_number)
            search_box.send_keys(Keys.RETURN)
            # 找到这条记录
            sleep(1)
            release_element = self.ppm_driver.find_element(By.CSS_SELECTOR,
                                                           '[aria-label^="' + self.current_release_number + '"]')
            release_href = release_element.get_attribute('href')

            # 点击进入
            # Step 8.3.3 Click to open Release
            self.ppm_driver.get(release_href)
            current_release_change_input = self.ppm_driver.find_element(By.ID, 'sys_display.rm_release_scrum'
                                                                               '.u_current_parent_update_set')

            # Step 8.3.4 Change Current Update Set & enter down
            current_release_change_input.clear()
            print(self.C_TEST_Update_Set["name"])
            current_release_change_input.send_keys(self.C_TEST_Update_Set["name"])
            sleep(1)
            current_release_change_input.send_keys(Keys.RETURN)

            # Step 8.3.5 Click Update Button
            update_btn = self.ppm_driver.find_element(By.ID, 'sysverb_update')
            update_btn.click()
            sleep(1)
            self.ppm_driver.close()

        # 需要deploy的正常流程
        # print("Ready to open current Update set")
        element = self.dev_driver.find_element(By.CSS_SELECTOR,
                                               '[aria-label="Open record: ' + self.C_TEST_Update_Set["name"] + '"]')
        self.C_TEST_Update_Set["href"] = element.get_attribute('href')

    def open_current_update_set(self):
        # Step 9.Open Current Update Set
        self.dev_driver.get(self.C_TEST_Update_Set["href"])
        get_number_btn = self.dev_driver.find_element(By.ID, '31364893dbe4cc503da8366af4961939_bottom')

        # Step 10. Click Get Numbers to get story number in current update set
        get_number_btn.click()
        sleep(2)
        output_div = self.dev_driver.find_element(By.CSS_SELECTOR, '[class="outputmsg_text"]')
        # 获取元素的文本内容
        div_content = output_div.get_attribute("innerHTML")
        # 使用字符串分割提取目标值
        values = div_content.split("<br>")

        # 11. Copy Numbers
        # 保存dev中safe story的名称
        self.dev_story_number = [value for value in values if value.startswith("SFSTRY")]

    def compare(self):

        # 12. Compare story number between dev and ppm in local code
        ppm_different_numbers = list(set(self.ppm_story_number) - set(self.dev_story_number))
        dev_different_numbers = list(set(self.dev_story_number) - set(self.ppm_story_number))
        count_ppm = len(ppm_different_numbers)
        count_dev = len(dev_different_numbers)

        if count_ppm != 0:
            missing_numbers_str = ", ".join(ppm_different_numbers)
            # Step 12.0 Display Message to inform release manager which stories are different
            print(count_ppm, "stories are found in PPM but not in Dev: ", missing_numbers_str)

            self.ppm_driver = open_webpage("nonprod_ppm_story")

            # Step 12.2.1 open the different story link
            for number in ppm_different_numbers:
                search_box = self.ppm_driver.find_element(By.CSS_SELECTOR, '[aria-label="Search"]')
                # 找到下拉选择框元素
                select_element = self.ppm_driver.find_element(By.CSS_SELECTOR, '[aria-label="Search a specific field '
                                                                               'of the SAFe stories list, 19 items"]')
                # 创建 Select 对象
                select = Select(select_element)
                # 通过可见文本选择选项
                select.select_by_visible_text('Number')

                search_box.send_keys(number)
                search_box.send_keys(Keys.RETURN)

                number_element = self.ppm_driver.find_element(By.XPATH, "//table[@id='sn_safe_story_table']//tr["
                                                                        "1]/td[3]")
                number_element.click()

                # Step 12.2.2 Change story state to "Work In Progress"
                state = self.ppm_driver.find_element(By.ID, 'sn_safe_story.state')
                select = Select(state)
                # 选择 value 为 "IN" 的选项
                select.select_by_value("2")

                # Step 12.2.3 Copy developer's information
                developer_name_element = self.ppm_driver.find_element(By.ID, 'sys_display.sn_safe_story.u_developer'
                                                                          )
                developer = re.sub(r'\(.*\)', '', developer_name_element.get_attribute("value"))

                # Step 12.2.4 Change "Assigned to" field to developer's name
                assigned_to_input = self.ppm_driver.find_element(By.NAME, 'assigned_to')
                assigned_to_input.click()
                if assigned_to_input.get_attribute("value"):
                    assigned_to_input.clear()
                assigned_to_input.send_keys(developer)
                sleep(0.5)
                assigned_to_input.send_keys(Keys.RETURN)

                # Step 12.2.5 Choose tab "History"
                history_btn = self.ppm_driver.find_element(By.XPATH, "//span[contains(text(), 'History')]")
                history_btn.click()

                text_area = self.ppm_driver.find_element(By.ID, 'activity-stream-textarea')
                text_area.send_keys("Hi @" + developer)
                sleep(0.5)
                text_area.send_keys(Keys.RETURN)
                text_area.send_keys(", please pay attention that your story state is not ready for test deploy but "
                                    "your update set was added in the current update set and set as completed. I "
                                    "set your story back to WIP. Please double check. Thanks.")

                # Step 12.2.7 Click Post button & Save button
                post_btn = self.ppm_driver.find_element(By.XPATH, "//button[contains(text(), 'Post')]")
                post_btn.click()  # 需修改 先不用发了……发的太多了

                update_btn = self.ppm_driver.find_element(By.ID, 'sysverb_update')
                update_btn.click()

            return False

        elif count_dev != 0:
            missing_numbers_str = ", ".join(dev_different_numbers)
            print(count_dev, "stories are found in Dev but not in PPM: ", missing_numbers_str)
            sleep(1)
            # Step 12.1.1 Click "Child Update Sets" in Dev
            get_update_btn = self.dev_driver.find_element(By.XPATH, "//span[contains(text(), 'Child Update Sets')]")
            get_update_btn.click()
            sleep(1)

            # 定位表格元素
            table_element = self.dev_driver.find_element(By.XPATH,
                                                         "//table[@id='sys_update_set.sys_update_set.parent_table"
                                                         "']/tbody")

            # 获取tbody中的所有行元素
            rows = table_element.find_elements(By.XPATH, ".//tr")
            href_list = []
            # 遍历每一行，提取第三个td中a标签的href属性
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, 'td')
                if len(cells) >= 3:
                    third_cell = cells[2]
                    link = third_cell.find_element(By.XPATH, ".//a")
                    if any(number in link.text for number in dev_different_numbers):
                        href = link.get_attribute("href")
                        href_list.append({
                            "story": link.text,
                            "href": href
                        })  # 将 href 添加到 href_list 中

            # Step 12.1.2 Click every update set related to the different story to open
            # 遍历 href_list
            for href in href_list:
                self.dev_driver.execute_script("window.open('about:blank', '_blank');")
                # 切换到新打开的标签页
                self.dev_driver.switch_to.window(self.dev_driver.window_handles[-1])
                # 导航到 href
                self.dev_driver.get(href["href"])
                # Step 12.1.5 Clear the string field "Parent"
                parent_input = self.dev_driver.find_element(By.ID, "sys_display.sys_update_set.parent")
                # parent_input.clear()  # clear无法删除
                # 全选输入字段中的文本
                parent_input.send_keys(Keys.CONTROL + "a")
                # 删除选定的文本
                parent_input.send_keys(Keys.DELETE)
                parent_input.send_keys(Keys.RETURN)

                # Step 12.1.6 Click Save button
                save_btn = self.dev_driver.find_element(By.ID, 'sysverb_update_and_stay')
                save_btn.click()

                # 关闭当前标签页
                self.dev_driver.close()
                # 切换回原始标签页
                self.dev_driver.switch_to.window(self.dev_driver.window_handles[0])

                # Step 12.1.7 Open the story in PPM
                self.ppm_driver = open_webpage("nonprod_ppm_story")
                search_box = self.ppm_driver.find_element(By.CSS_SELECTOR, '[aria-label="Search"]')
                search_box.send_keys(href["story"])
                search_box.send_keys(Keys.RETURN)
                sleep(0.5)
                result = self.ppm_driver.find_element(By.CSS_SELECTOR,
                                                      '[aria-label="Open record: ' + href["story"] + '"]')
                result.click()
                sleep(1)

                history_btn = self.ppm_driver.find_element(By.XPATH, "//span[contains(text(), 'History')]")
                history_btn.click()

                developer_name_element = self.ppm_driver.find_element(By.ID, 'sys_display.sn_safe_story.u_developer'
                                                                      )
                developer = re.sub(r'\(.*\)', '', developer_name_element.get_attribute("value"))

                text_area = self.ppm_driver.find_element(By.ID, 'activity-stream-textarea')
                text_area.send_keys("Hi @" + developer)
                sleep(0.5)
                text_area.send_keys(Keys.RETURN)
                text_area.send_keys(", please make sure your update set state matches your story state. When your "
                                    "story is ready for test deploy, your update set should be completed. Thanks.")

                post_btn = self.ppm_driver.find_element(By.XPATH, "//button[contains(text(), 'Post')]")
                # post_btn.click()  # 需修改 先不用发了……发的太多了
            # 刷新当前标签页
            self.dev_driver.refresh()

            return False

        else:
            print("Stories in PPM are the same as those in Dev!")
            return True

    def filter(self):
        # Step 13. Filter the stories according to Numbers in ppm (filter by Number is one of them)
        self.ppm_driver = open_webpage("ppm")
        filter_btn = self.ppm_driver.find_element(By.ID,
                                                  'rm_story_filter_toggle_image'  # PPM里是这个，nonprod里是下面那个
                                                  # 'sn_safe_story_filter_toggle_image')
                                                  )
        filter_btn.click()
        sleep(2)
        select_number = self.ppm_driver.find_element(By.ID, 's2id_autogen1')
        select_number.click()
        input_number = self.ppm_driver.find_element(By.ID, 's2id_autogen2_search')
        input_number.send_keys('Number')
        # 定位 ul 元素
        ul_element = self.ppm_driver.find_element(By.ID, 'select2-results-2')
        # 定位所有 li 元素
        li_elements = ul_element.find_elements(By.TAG_NAME, 'li')
        # 获取第二个 li 元素并点击
        second_li_element = li_elements[1]
        second_li_element.click()

        select_in = self.ppm_driver.find_element(By.CSS_SELECTOR,
                                                 '[aria-label="Operator For Condition 1: Number starts with "')
        select = Select(select_in)
        # 选择 value 为 "IN" 的选项
        select.select_by_value("IN")

        text_number = self.ppm_driver.find_element(By.CSS_SELECTOR, '[aria-label="Input value for field: Number"')
        input_text = '\n'.join(self.dev_story_number)
        text_number.send_keys(input_text)

        # 按下RUN按钮
        run_btn = self.ppm_driver.find_element(By.ID, 'test_filter_action_toolbar_run')
        run_btn.click()

        sleep(1)
        # Step 14. Select all filtered stories
        select_all_btn = self.ppm_driver.find_element(By.CSS_SELECTOR, '[class="input-group-checkbox"]')
        select_all_btn.click()

        # Step 15. Select actions: check stories to be deployed
        check_stories_select = self.ppm_driver.find_element(By.ID, 'rm_story_choice_actions')
        select_check = Select(check_stories_select.find_element(By.TAG_NAME, 'select'))
        select_check.select_by_value("b476cb1cdbc5489020627d78f496195f")

    def check_stories_deployed(self):

        # 切换到新标签页
        window_handles = self.ppm_driver.window_handles
        self.ppm_driver.switch_to.window(window_handles[-1])

        # 在新标签页中完成操作
        new_tab_title = self.ppm_driver.title
        print("New tab title:", new_tab_title)

        # 查找表格元素
        table = self.ppm_driver.find_element(By.ID, 'rm_task_table')
        # 检查是否有内容
        if table.text.strip():
            # Step 15.1 Display Message to inform release manager of the stories that require manual check (show links)
            print("There're stories that require manual check!")
            # 有的话还需要把link给manager，之后添加
        else:
            print("No stories require manual check.")

    def set_complete_and_export(self):
        # Step 16. Set the update set Complete.
        state_select = self.dev_driver.find_element(By.ID, 'sys_update_set.state')
        select = Select(state_select)
        # 选择 value 为 "complete" 的选项
        select.select_by_value("complete")
        # 点击保存按钮
        save_btn = self.dev_driver.find_element(By.ID, 'sysverb_update_and_stay')
        save_btn.click()

        # Step 17. Click related link: Export Update Set Batch to XML
        export_btn = self.dev_driver.find_element(By.ID, 'd2339c9347231200c17e19fbac9a7173')
        export_btn.click()

    def open_test(self):
        # Step 18. Copy Current Update Set
        # Step 19. Open Retrieved Update Sets
        test_driver = open_webpage("test")
        search_box = test_driver.find_element(By.CSS_SELECTOR, '[aria-label="Search"]')

        # Step 20. Input Current Update Set in search box & enter down
        search_box.send_keys(self.C_TEST_Update_Set["name"])
        search_box.send_keys(Keys.RETURN)

        sleep(1)
        # 查找包含特定文本的元素
        no_records_element = test_driver.find_element(By.XPATH, "//*[contains(text(), 'No records to display')]")
        # Step 21. Check if there's result
        if no_records_element:
            # 如果不存在结果，则继续进行
            # Step 22. Click related link: Import Update Set from XML
            import_div = test_driver.find_element(By.ID, '0583c6760a0a0b8000d06ad9224a81a2')
            import_div.click()
            sleep(1.5)
            # Step 23. Click Choose File button
            load_file_btn = test_driver.find_element(By.ID, 'attachFile')
            # load_file_btn.click()
        else:
            print("There's already an imported record.")

        return test_driver
