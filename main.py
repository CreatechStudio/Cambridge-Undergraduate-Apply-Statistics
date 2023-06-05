import pytest
import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support import expected_conditions as EC

from bs4 import BeautifulSoup
import requests
import json

import sys
import os

import re
import openpyxl


def filePreProcess():
    # 检查文件是否存在
    if os.path.exists("output.txt"):
        # 如果存在，则删除文件
        os.remove("output.txt")
    if os.path.exists("output.xlsx"):
        os.remove("output.xlsx")

class Main:
    def setup_method(self, method=None):
        self.driver = webdriver.Chrome()
        self.vars = {}

    def teardown_method(self, method=None):
        self.driver.quit()

    def autoFunction(self):
        self.driver.get("https://www.undergraduate.study.cam.ac.uk/apply/statistics")
        self.driver.find_element(By.ID, "edit-year").click()
        self.driver.find_element(By.ID, "edit-open-open").click()
        self.driver.find_element(By.ID, "edit-winter-winter").click()
        self.driver.find_element(By.ID, "edit-summer-summer").click()
        self.driver.find_element(By.ID, "edit-group-course").click()

        for year in range(2013, 2023):
            dropdown = self.driver.find_element(By.ID, "edit-year")
            # dropdown.find_element(By.XPATH, "//option[. = '2021']").click()
            replaced_xpath = "//option[. = '{}']".format(year)
            dropdown.find_element(By.XPATH, replaced_xpath).click()
            for i in range(1, 32):
                if i == 18 or i == 19:
                    continue
                self.driver.find_element(
                    By.CSS_SELECTOR, "#edit_colleges_chosen > .chosen-choices"
                ).click()
                css_selector = f".chosen-results > li:nth-child({i})"
                self.driver.find_element(By.CSS_SELECTOR, css_selector).click()
                time.sleep(2)
                self.driver.find_element(By.ID, "edit-submit").click()
                try:
                    wait = WebDriverWait(self.driver, 10)
                    wait.until(
                        EC.visibility_of_element_located(
                            (By.CLASS_NAME, "highcharts-background")
                        )
                    )
                except Exception as message:
                    print("元素定位报错%s" % message)
                finally:
                    pass

                # 假设html为包含完整HTML代码的字符串
                soup = BeautifulSoup(self.driver.page_source, "html.parser")
                # 查找包含"data-chart"的<div>标签
                div_tag = soup.find(
                    "div",
                    class_="charts-highchart chart charts-highchart-processed",
                    attrs={"data-chart": True},
                )
                # 获取"data-chart"属性的值
                data_chart = div_tag["data-chart"]
                # 解析"data-chart"属性值中的JSON数据
                json_data = json.loads(data_chart)
                university_name = soup.select_one(".search-choice span").text

                # 打开文件
                file = open("output.txt", "a")  # "w" 表示写入模式
                # 将输出重定向到文件
                sys.stdout = file

                subject_data = {}
                for item in json_data["series"]:
                    subject = item["name"]
                    data = item["data"]
                    subject_data[subject] = data

                # 输出数据
                print("University:", university_name)
                print("Year:", year)
                for subject, data in subject_data.items():
                    print("Name:", subject)
                    print("Categories:", json_data["xAxis"][0]["categories"])
                    print("Data:", data)
                    print()
                # 恢复标准输出
                sys.stdout = sys.__stdout__
                # 关闭文件
                file.close()

                self.driver.find_element(
                    By.CSS_SELECTOR, ".search-choice-close"
                ).click()

def dataProcess():
    def parse_data(filename):
        with open(filename, 'r') as file:
            lines = file.readlines()

        data = []
        university = ''
        year = ''
        current_name = ''
        current_categories = []
        current_data = []

        for line in lines:
            line = line.strip()
            if line.startswith('University:'):
                university = line.split(':')[1].strip()
            elif line.startswith('Year:'):
                year = line.split(':')[1].strip()
            elif line.startswith('Name:'):
                current_name = line.split(':')[1].strip()
            elif line.startswith('Data:'):
                current_data = re.findall(r'\d+', line)
                current_data = list(map(int, current_data))

                if len(current_data) < len(current_categories):
                    current_data += [0] * (len(current_categories) - len(current_data))

                data.append({
                    'University': university,
                    'Year': year,
                    'Name': current_name,
                    'Categories': current_categories,
                    'Data': current_data
                })

            elif line.startswith('Categories:'):
                current_categories = re.findall(r"'(.*?)'", line)

        return data

    def fill_missing_data(data, categories):
        for item in data:
            missing_categories = set(categories) - set(item['Categories'])
            missing_categories_data = [0] * len(missing_categories)
            for category in missing_categories:
                index = categories.index(category)
                item['Categories'].insert(index, category)
                item['Data'].insert(index, 0)

        return data

    def write_to_excel(data, output_file):
        workbook = openpyxl.Workbook()
        sheet = workbook.active

        # Write header row with complete category list
        header = ['University', 'Year', 'Name'] + categories
        sheet.append(header)

        # Write data rows
        for item in data:
            row = [item['University'], item['Year'], item['Name']]
            row_data = item['Data']

            if len(row_data) < len(categories):
                row_data += [0] * (len(categories) - len(row_data))

            row += row_data
            sheet.append(row)

        workbook.save(output_file)

    filename = 'output.txt'
    output_file = 'output.xlsx'

    parsed_data = parse_data(filename)

    # Extract all categories from the data
    categories = [
        "Anglo-Saxon, Norse, and Celtic",
        "Archaeology",
        "Architecture",
        "Asian and Middle Eastern Studies",
        "Chemical Engineering via Engineering",
        "Chemical Engineering via Natural Sciences",
        "Classics",
        "Classics (4 years)",
        "Computer Science",
        "Economics",
        "Education",
        "Engineering",
        "English",
        "Foundation Year in Arts, Humanities and Social Sciences",
        "Geography",
        "History",
        "History and Modern Languages",
        "History and Politics",
        "History of Art",
        "Human, Social, and Political Sciences",
        "Land Economy",
        "Law",
        "Linguistics",
        "Mathematics",
        "Medicine",
        "Medicine (Graduate course)",
        "Modern and Medieval Languages",
        "Music",
        "Natural Sciences",
        "Philosophy",
        "Psychological and Behavioural Sciences",
        "Theology, Religion and Philosophy of Religion",
        "Veterinary Medicine"
    ]

    filled_data = fill_missing_data(parsed_data, categories)
    write_to_excel(filled_data, output_file)

if __name__ == "__main__":
    test = Main()
    test.setup_method()
    print("自动化开始")
    filePreProcess()
    test.autoFunction()
    time.sleep(2)
    test.teardown_method()
    print("网页处理结束，开始生成Excel")
    dataProcess()
    print("生成结束，程序结束")
