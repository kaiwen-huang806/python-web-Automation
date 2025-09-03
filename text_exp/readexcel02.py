import re
import os
import subprocess
import pandas as pd
import time
import pytest
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import allure


# 读取Excel测试用例（确保必要列存在，处理空值）
def get_test_cases():
    excel_path = "text_cases.xlsx"
    required_cols = ["用例 ID", "测试步骤", "预期结果"]  # 确保Excel包含这三列
    try:
        df = pd.read_excel(excel_path, sheet_name="Sheet1")
        df = df.ffill()  # 向上填充空值，避免缺失数据
        # 检查必要列是否存在
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise KeyError(f"Excel缺少必要列：{', '.join(missing_cols)}")
        return df
    except FileNotFoundError:
        allure.attach(f"文件路径：{excel_path}", "Excel文件未找到")
        raise
    except Exception as e:
        allure.attach(str(e), "读取用例失败详情")
        raise


# 登录测试用例（参数化执行4个用例）
@pytest.mark.parametrize("case_id", ["Login-001", "Login-002", "Login-003", "Login-004"])
def test_login(case_id):
    # 1. 初始化用例信息，同步到Allure报告
    df = get_test_cases()
    login_case = df[df["用例 ID"] == case_id].iloc[0]
    allure.dynamic.feature("用户登录功能")
    allure.dynamic.story(f"登录用例 {case_id}")
    allure.dynamic.title(login_case.get("用例标题", f"登录验证_{case_id}"))
    allure.dynamic.description(f"核心判定标准：登录后URI必须完全等于 http://120.24.56.229:8082")
    allure.dynamic.severity(allure.severity_level.CRITICAL)

    # 2. 配置核心参数（登录成功的唯一URI）
    TARGET_SUCCESS_URI = "http://120.24.56.229:8082"  # 你指定的目标URI
    driver_path = "D:\\Users\\29430\\PycharmProjects\\PythonProject1\\msedgedriver.exe"

    # 3. 初始化浏览器（Edge驱动，添加全局等待提升稳定性）
    service = Service(driver_path)
    driver = webdriver.Edge(service=service)
    driver.maximize_window()
    driver.implicitly_wait(10)  # 等待元素加载的全局超时时间

    try:
        # 步骤1：提取并访问初始URL（从用例步骤或用默认值）
        with allure.step("1. 提取并访问初始测试地址"):
            test_steps = login_case["测试步骤"].split("\n")
            initial_url = None
            # 从测试步骤中提取“访问登录页面”对应的URL
            for step in test_steps:
                if "访问登录页面" in step:
                    url_match = re.search(r'(https?://\S+)', step)
                    if url_match:
                        initial_url = url_match.group(1).strip()
                        break
            # 提取失败时，默认访问登录入口所在的根地址
            if not initial_url:
                initial_url = "http://120.24.56.229:8082"  # 与目标URI一致（若登录入口在根页面）
                allure.attach(f"未从步骤提取URL，使用默认初始地址：{initial_url}", "URL来源")
            else:
                allure.attach(f"从步骤提取初始URL：{initial_url}", "URL来源")

            # 访问初始地址并截图
            driver.get(initial_url)
            time.sleep(2)  # 等待页面完全加载
            allure.attach(driver.get_screenshot_as_png(), "访问初始页面后", allure.attachment_type.PNG)

        # 步骤2：点击登录入口，进入登录表单页
        with allure.step("2. 点击登录链接，进入登录页"):
            # 定位登录链接（按实际页面XPath，确保可点击）
            login_link = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/div/ul[1]/div/div/a[1]"))
            )
            login_link.click()
            time.sleep(2)
            # 记录当前URL（确认进入登录页，避免后续操作错位）
            current_page_url = driver.current_url
            allure.attach(f"进入登录页后的URL：{current_page_url}", "登录页状态")
            allure.attach(driver.get_screenshot_as_png(), "登录页面截图", allure.attachment_type.PNG)

        # 步骤3：输入用户名（从用例提取或用默认值）
        with allure.step("3. 输入用户名"):
            username = None
            for step in test_steps:
                if "输入用户名" in step:
                    username = step.split(":")[-1].strip()
                    # 处理“空值”场景（若用例需要测试空用户名）
                    if username.lower() == "空":
                        username = ""
                    break
            # 提取失败时使用默认测试账号
            if username is None:
                username = "testuser1234"
                allure.attach(f"未提取到用户名，使用默认值：{username}", "用户名信息")
            else:
                allure.attach(f"提取到用户名：{username}", "用户名信息")

            # 定位用户名输入框并输入（清空后输入，避免残留）
            username_input = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH,
                                                "/html/body/div[1]/div[1]/div[3]/div/div[2]/div[2]/div[2]/div/div[1]/div[1]/form/div[1]/input"))
            )
            username_input.clear()
            username_input.send_keys(username)
            time.sleep(1)

        # 步骤4：输入密码（从用例提取或用默认值）
        with allure.step("4. 输入密码"):
            password = None
            for step in test_steps:
                if "输入密码" in step:
                    password = step.split(":")[-1].strip()
                    # 处理“空值”场景
                    if password.lower() == "空":
                        password = ""
                    break
            # 提取失败时使用默认测试密码
            if not password:
                password = "Test123456!"
                allure.attach(f"未提取到密码，使用默认值：{password}", "密码信息")
            else:
                allure.attach(f"提取到密码：{password}", "密码信息")

            # 定位密码输入框并输入（清空后输入）
            password_input = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH,
                                                "/html/body/div[1]/div[1]/div[3]/div/div[2]/div[2]/div[2]/div/div[1]/div[1]/form/div[2]/input"))
            )
            password_input.clear()
            password_input.send_keys(password)
            time.sleep(1)

        # 步骤5：点击登录按钮，提交登录请求
        with allure.step("5. 点击登录按钮，提交请求"):
            login_btn = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH,
                                            "/html/body/div[1]/div[1]/div[3]/div/div[2]/div[2]/div[2]/div/div[1]/div[1]/form/div[3]/button"))
            )
            login_btn.click()
            time.sleep(3)  # 延长等待，确保URI跳转完成（关键！避免未跳转就判断）
            allure.attach(driver.get_screenshot_as_png(), "点击登录后页面", allure.attachment_type.PNG)

        # 步骤6：核心验证——URI是否完全等于目标地址
        with allure.step("6. 验证登录结果（URI完全匹配判定）"):
            # 获取当前URI（去除末尾可能的“/”，确保格式统一）
            current_uri = driver.current_url.rstrip("/")
            # 目标URI也统一格式（去除末尾“/”）
            target_uri_processed = TARGET_SUCCESS_URI.rstrip("/")

            # 记录关键对比信息到Allure报告
            allure.attach(f"当前URI（处理后）：{current_uri}", "URI对比详情")
            allure.attach(f"目标URI（处理后）：{target_uri_processed}", "URI对比详情")

            # 严格判定：只有当前URI与目标URI完全一致，才是登录成功
            if current_uri == target_uri_processed:
                actual_result = "登录成功"
                allure.attach("✅ 当前URI与目标URI完全一致，判定登录成功", "结果判定依据")
            else:
                actual_result = "登录失败"
                allure.attach(
                    f"❌ 当前URI与目标URI不一致（当前：{current_uri} | 目标：{target_uri_processed}），判定登录失败",
                    "结果判定依据")

            # 从Excel获取预期结果，进行断言
            expected_result = login_case["预期结果"].strip()
            allure.attach(f"预期结果：{expected_result}\n实际结果：{actual_result}", "最终结果对比")

            # 断言失败时，明确提示用例ID和URI差异
            assert actual_result == expected_result, \
                f"用例 {case_id} 失败！预期：{expected_result}，实际：{actual_result}，URI差异：当前={current_uri} vs 目标={target_uri_processed}"

    # 捕获所有异常，记录截图和详情到报告
    except Exception as e:
        error_msg = f"用例 {case_id} 执行异常：{str(e)}"
        allure.attach(error_msg, "异常详情", allure.attachment_type.TEXT)
        allure.attach(driver.get_screenshot_as_png(), "异常时页面截图", allure.attachment_type.PNG)
        raise  # 重新抛出异常，标记用例失败

    # 无论成功/失败，确保浏览器关闭
    finally:
        time.sleep(2)
        driver.quit()
        allure.attach("浏览器已关闭", "资源释放状态")


# 主函数：执行测试并生成Allure报告（Windows适配）
if __name__ == "__main__":
    allure_results_dir = "allure-results"
    # 创建Allure结果目录（若不存在）
    if not os.path.exists(allure_results_dir):
        os.makedirs(allure_results_dir)
        print(f"✅ 创建Allure结果目录：{os.path.abspath(allure_results_dir)}")
    else:
        print(f"ℹ️ Allure结果目录已存在：{os.path.abspath(allure_results_dir)}")

    # 运行Pytest，生成Allure结果文件
    print("\n🚀 开始执行登录测试...")
    pytest_args = [
        "-s",  # 显示控制台输出
        "--alluredir", allure_results_dir,  # 指定结果存储目录
        __file__  # 当前测试文件
    ]
    pytest_exit_code = pytest.main(pytest_args)

    # 生成并打开Allure报告（仅当结果目录非空）
    if os.path.exists(allure_results_dir) and os.listdir(allure_results_dir):
        print(f"\n📊 正在生成Allure报告...")
        try:
            # Windows系统执行allure serve，自动打开浏览器展示报告
            subprocess.run(
                f"allure serve {allure_results_dir}",
                shell=True,
                check=True,
                cwd=os.getcwd()
            )
        except subprocess.CalledProcessError as e:
            print(f"❌ 报告生成失败：{e}")
            print("ℹ️ 请确认：1. Allure已安装 2. 环境变量已配置 3. 结果目录有.json文件")
        except FileNotFoundError:
            print("❌ 未找到Allure命令，请先安装Allure并配置环境变量")
    else:
        print(f"❌ Allure结果目录为空，无法生成报告（测试可能未执行或执行失败）")
    print(f"\n📋 测试退出码：{pytest_exit_code}（0=全部通过，非0=存在失败）")