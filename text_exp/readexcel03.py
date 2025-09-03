import os
import subprocess
import sys
import pandas as pd
import time
import pytest
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import allure

@pytest.fixture(scope="module")
def setup():
    excel_path = "text_cases.xlsx"
    driver_path = "D:\\Users\\29430\\PycharmProjects\\PythonProject1\\msedgedriver.exe"
    service = Service(driver_path)
    df = pd.read_excel(excel_path, sheet_name="Sheet2")
    df = df.ffill()
    return service, df

@pytest.mark.parametrize("case_id", ["Register-001", "Register-002", "Register-003", "Register-004"])
def test_register_cases(setup, case_id):
    service, df = setup
    register_case = df[df["用例 ID"] == case_id].iloc[0]

    driver = webdriver.Edge(service=service)
    driver.maximize_window()

    try:
        # 提取测试 URL
        test_steps = register_case["测试步骤"].split("\n")
        test_url = None
        for step in test_steps:
            if "访问注册页面" in step:
                parts = step.split("：")
                if len(parts) > 1:
                    test_url = parts[-1].strip()
                    break
        test_url = test_url if test_url else "http://120.24.56.229:8082"
        if not (test_url.startswith(("http://", "https://")) and "." in test_url):
            raise ValueError(f"URL 格式无效，当前值：{test_url}")
        driver.get(test_url)
        time.sleep(3)

        # 进入注册页面
        register_link = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/div/ul[1]/div/div/a[2]"))
        )
        register_link.click()
        time.sleep(2)

        # 输入用户名
        username = None
        for step in test_steps:
            if "输入符合规则的用户名" in step:
                username = step.split(":")[-1].strip()
                break
        username = username if username else "testuser1234"
        username_input = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located(
                (By.XPATH, "/html/body/div[1]/div[1]/div[3]/div/div/div[2]/div/div[1]/div[1]/form/div[1]/input"))
        )
        username_input.send_keys(username)
        time.sleep(1)

        # 输入密码
        password = None
        for step in test_steps:
            if "输入符合规则的密码" in step:
                password = step.split(":")[-1].strip()
                break
        password = password if password else "Test123456!"
        password_input = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located(
                (By.XPATH, "/html/body/div[1]/div[1]/div[3]/div/div/div[2]/div/div[1]/div[1]/form/div[2]/input"))
        )
        password_input.send_keys(password)
        time.sleep(1)

        # 勾选协议
        agreement_checkbox = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[3]/div/div/div[2]/div/div[1]/div[1]/form/div[3]/label"))
        )
        if not agreement_checkbox.is_selected():
            agreement_checkbox.click()
            allure.attach("已勾选协议", "协议状态")
            time.sleep(1)
        else:
            allure.attach("协议已默认勾选，无需操作", "协议状态")

        # 提交注册
        register_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[1]/div[1]/div[3]/div/div/div[2]/div/div[1]/div[1]/form/div[4]/button"))
        )
        register_btn.click()
        time.sleep(2)

        # 验证结果
        expected_url = "http://120.24.56.229:8082"
        current_url = driver.current_url.rstrip("/")
        expected_url_processed = expected_url.rstrip("/")
        actual_result = "注册成功" if current_url == expected_url_processed else "注册失败"
        if actual_result == "注册失败":
            allure.attach(f"URL 不匹配：当前 URL={current_url}，预期 URL={expected_url_processed}", "错误信息")
        expected_result = register_case["预期结果"].strip()
        assert actual_result == expected_result, f"用例 {case_id} 测试失败"
    finally:
        driver.quit()

if __name__ == "__main__":
    allure_results_dir = "allure-results"
    if not os.path.exists(allure_results_dir):
        os.makedirs(allure_results_dir)
        print(f"创建目录：{allure_results_dir}")
    else:
        print(f"目录已存在：{allure_results_dir}")

    pytest.main(["-s", "--alluredir", allure_results_dir])

    if os.path.exists(allure_results_dir) and os.listdir(allure_results_dir):
        print(f"Allure 结果目录非空：{allure_results_dir}")
        try:
            subprocess.run(f'allure serve {allure_results_dir}', shell=True, check=True)
        except subprocess.CalledProcessError as e:
            print(f"生成 Allure 报告出错: {e}")
    else:
        print(f"Allure 结果目录为空或不存在：{allure_results_dir}")