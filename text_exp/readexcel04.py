from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import allure  # 引入Allure库
import pytest  # 引入pytest测试框架


# --------------------------
# 1. 修复：用例加载改为普通函数（而非Fixture）
# --------------------------
def load_test_cases():
    """
    普通函数：读取Excel测试用例数据（直接调用，不通过Fixture）
    返回：用例字典列表（适配pytest参数化）
    """
    EXCEL_FILE_PATH = "text_cases.xlsx"  # 你的Excel路径
    EXCEL_SHEET_NAME = "Sheet3"          # 你的工作表名称
    try:
        # 读取并清理用例数据（去重、填充空值）
        excel_file = pd.ExcelFile(EXCEL_FILE_PATH)
        df = excel_file.parse(EXCEL_SHEET_NAME)
        fill_cols = ["用例 ID", "模块", "功能点", "测试标题", "前置条件", "测试步骤", "预期结果"]
        df[fill_cols] = df[fill_cols].ffill()  # 填充空值
        # 去重：按用例ID保留唯一用例，避免重复执行
        df_valid = df.dropna(subset=["用例 ID"]).drop_duplicates(subset=["用例 ID"], keep="first")
        print(f"✅ 加载用例成功：共 {len(df_valid)} 个有效用例")
        return df_valid.to_dict("records")  # 转换为字典列表，供参数化使用
    except Exception as e:
        print(f"❌ 加载用例失败：{str(e)}")
        raise  # 加载失败时终止程序


# --------------------------
# 2. 浏览器初始化Fixture（保留，用于前置/后置操作）
# --------------------------
@pytest.fixture(scope="session", autouse=True)
def init_browser():
    """
    pytest Fixture：会话级浏览器初始化（所有用例共用一个浏览器）
    scope="session"：整个测试会话只执行1次（启动/关闭浏览器）
    autouse=True：自动应用到所有测试用例
    """
    EDGE_DRIVER_PATH = "D:\\Users\\29430\\PycharmProjects\\PythonProject1\\msedgedriver.exe"  # 你的驱动路径
    driver = None
    try:
        # 启动浏览器
        service = Service(executable_path=EDGE_DRIVER_PATH)
        driver = webdriver.Edge(service=service)
        driver.maximize_window()  # 最大化窗口，避免元素定位受分辨率影响
        driver.implicitly_wait(5)  # 全局隐式等待，提升稳定性
        print("✅ 浏览器初始化成功")
        yield driver  # yield前：前置操作（启动浏览器）；yield后：后置操作（关闭浏览器）
    except Exception as e:
        print(f"❌ 浏览器初始化失败：{str(e)}")
        raise
    finally:
        # 后置操作：无论用例是否成功，都关闭浏览器
        if driver:
            driver.quit()
            print("✅ 浏览器已关闭")


# --------------------------
# 3. Allure步骤封装（复用+报告记录）
# --------------------------
@allure.step("步骤1：访问首页 -> {homepage_url}")
def open_homepage(driver, homepage_url="http://120.24.56.229:8082"):
    """访问首页并等待加载完成，步骤记录到Allure报告"""
    driver.get(homepage_url)
    # 显式等待：确保搜索框可交互（验证首页加载完成）
    WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="search-input"]'))
    )
    # 附加截图到报告（便于问题排查）
    allure.attach(driver.get_screenshot_as_png(), name="首页截图", attachment_type=allure.attachment_type.PNG)
    allure.attach(driver.current_url, name="首页URL", attachment_type=allure.attachment_type.TEXT)


@allure.step("步骤2：在搜索框输入 -> {search_text}")
def input_search_text(driver, search_text):
    """在搜索框输入内容，步骤记录到Allure报告"""
    # 显式等待：确保搜索框可点击
    search_box = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="search-input"]'))
    )
    search_box.clear()  # 清空输入框（避免残留内容）
    search_box.send_keys(search_text)
    # 附加输入日志到报告
    allure.attach(f"搜索框输入内容：{repr(search_text)}", name="输入日志", attachment_type=allure.attachment_type.TEXT)


@allure.step("步骤3：点击搜索按钮")
def click_search_button(driver):
    """点击搜索按钮并等待结果加载，步骤记录到Allure报告"""
    # 显式等待：确保搜索按钮可点击
    search_btn = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="ai-topsearch"]'))
    )
    search_btn.click()
    # 显式等待：确保结果区域加载完成（无论成功/失败）
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[1]/div[5]/div/div[4]'))
    )
    # 附加搜索后截图和URL到报告
    allure.attach(driver.get_screenshot_as_png(), name="搜索后页面截图", attachment_type=allure.attachment_type.PNG)
    allure.attach(driver.current_url, name="搜索后URL", attachment_type=allure.attachment_type.TEXT)


# --------------------------
# 4. 核心测试用例（适配Allure+Pytest参数化）
# --------------------------
@allure.feature("商品搜索模块")  # 标记测试模块（大功能点，如“商品搜索”）
@allure.story("搜索功能验证")     # 标记测试场景（子功能点，如“正常搜索/空搜索”）
# 修复：参数化调用普通函数load_test_cases()，而非Fixture
@pytest.mark.parametrize("test_case", load_test_cases())
def test_search_function(init_browser, test_case):
    """
    商品搜索测试用例：每个Excel用例会生成1个独立测试用例
    :param init_browser: pytest Fixture注入的浏览器驱动
    :param test_case: 参数化传入的单条用例数据（字典格式）
    """
    driver = init_browser  # 从Fixture获取浏览器驱动
    # 提取用例核心信息（用于Allure报告和日志）
    case_id = str(test_case.get("用例 ID", "未知用例"))
    case_title = str(test_case.get("测试标题", "无标题"))
    expected_result = str(test_case.get("预期结果", "")).strip()
    steps = [s.strip() for s in str(test_case.get("测试步骤", "")).split("\n") if s.strip()]

    # 1. 标记用例信息到Allure报告（提升报告可读性）
    allure.dynamic.title(f"{case_id} - {case_title}")  # 报告中显示的用例标题
    allure.dynamic.description(f"""
    **用例ID**：{case_id}  
    **测试标题**：{case_title}  
    **预期结果**：{expected_result}  
    **测试步骤**：{chr(10).join(steps)}  # chr(10)是换行符，让步骤显示更清晰
    """)  # 用例详情（Markdown格式）

    try:
        # 2. 执行测试步骤（调用封装的Allure步骤）
        for step in steps:
            if "进入首页" in step:
                open_homepage(driver)
            elif "在搜索框商品名称" in step:
                # 提取搜索文本（处理“空输入”场景，如步骤描述为“在搜索框商品名称：空”）
                search_text = step.split("：")[-1].strip() if "：" in step else ""
                search_text = "" if search_text.lower() == "空" else search_text
                input_search_text(driver, search_text)
            elif "点击搜索按钮" in step:
                click_search_button(driver)

        # 3. 结果验证（按预期结果分场景判断）
        with allure.step(f"步骤4：验证结果（预期：{expected_result}）"):
            if expected_result == "搜索成功":
                # 验证“搜索成功”：有结果 + 包含目标关键词
                result_list = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[1]/div[5]/div/div[4]/ul'))
                )
                products = result_list.find_elements(By.XPATH, './/li')
                product_names = [p.find_element(By.XPATH, './/p[@class="am-text-truncate-2-md goods-title"]').text
                                for p in products]

                # 断言1：结果数量>0（确保有商品）
                assert len(products) > 0, "搜索成功场景下，结果列表为空"
                # 断言2：根据用例ID验证关键词（适配Search-001/Search-002）
                if case_id == "Search-001":
                    assert any("华为笔记本电脑" in name for name in product_names), \
                        f"无包含'华为笔记本电脑'的商品，实际结果：{product_names}"
                elif case_id == "Search-002":
                    assert any("华为" in name for name in product_names), \
                        f"无包含'华为'的商品，实际结果：{product_names}"

                # 附加成功结果到报告
                allure.attach(f"搜索结果：{product_names}", name="成功结果详情", attachment_type=allure.attachment_type.TEXT)
                print(f"✅ 用例 {case_id} 执行通过")

            elif expected_result == "搜索失败":
                # 验证“搜索失败”：分场景（Search-004空搜索跳转 / 其他无结果）
                if case_id == "Search-004":
                    # 空搜索跳转验证（目标地址：你提供的URL）
                    target_url = "http://120.24.56.229:8082/?s=search/index.html"
                    WebDriverWait(driver, 15).until(EC.url_to_be(target_url))  # 等待URL跳转
                    assert driver.current_url == target_url, \
                        f"空搜索跳转失败！预期：{target_url}，实际：{driver.current_url}"
                    allure.attach(f"空搜索成功跳转至：{target_url}", name="失败结果详情", attachment_type=allure.attachment_type.TEXT)
                else:
                    # 其他失败场景：验证“没有相关数据”提示或结果为空
                    try:
                        no_result_msg = WebDriverWait(driver, 15).until(
                            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[1]/div[5]/div/div[4]/p'))
                        )
                        assert no_result_msg.text == "没有相关数据", \
                            f"预期提示：没有相关数据，实际：{no_result_msg.text}"
                        allure.attach(f"无结果提示：{no_result_msg.text}", name="失败结果详情", attachment_type=allure.attachment_type.TEXT)
                    except:
                        # 备用验证：结果列表为空（避免提示文本定位失败）
                        result_list = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[5]/div/div[4]/ul')
                        products = result_list.find_elements(By.XPATH, './/li')
                        assert len(products) == 0, f"预期结果为空，实际数量：{len(products)}"
                        allure.attach(f"结果列表为空（数量：0）", name="失败结果详情", attachment_type=allure.attachment_type.TEXT)

                print(f"✅ 用例 {case_id} 执行通过")

    except AssertionError as e:
        # 捕获断言失败：附加截图和错误信息到报告（便于排查）
        allure.attach(driver.get_screenshot_as_png(), name=f"{case_id} 失败截图", attachment_type=allure.attachment_type.PNG)
        allure.attach(str(e), name=f"{case_id} 错误详情", attachment_type=allure.attachment_type.TEXT)
        print(f"❌ 用例 {case_id} 执行失败：{str(e)}")
        raise  # 抛出异常，让pytest标记用例为失败
    except Exception as e:
        # 捕获其他异常（如元素定位失败、页面加载超时）
        allure.attach(driver.get_screenshot_as_png(), name=f"{case_id} 异常截图", attachment_type=allure.attachment_type.PNG)
        allure.attach(str(e), name=f"{case_id} 异常详情", attachment_type=allure.attachment_type.TEXT)
        print(f"❌ 用例 {case_id} 执行异常：{str(e)}")
        raise


# --------------------------
# 程序入口（直接运行时执行pytest）
# --------------------------
if __name__ == "__main__":
    # 执行测试并生成Allure原始结果（存储在./allure-results）
    pytest.main([
        __file__,  # 当前脚本文件
        "--alluredir=./allure-results04",  # Allure结果目录
        "-v",  # 显示详细日志
        "-s"   # 显示print输出（可选）
    ])