import pandas as pd
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException
import pyperclip
import time
import os
import traceback
import random
import requests
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import tkinter as tk
from tkinter import messagebox
import ddddocr
import winsound
import multiprocessing




root = tk.Tk()
root.withdraw()  
root.attributes('-topmost', True)
current_dir = os.path.dirname(os.path.abspath(__file__))
def write_result(username, success):
    with open("result.log", "a", encoding="utf-8") as f:
        if success:
            f.write(f"{username} 成功\n")
        else:
            f.write(f"{username} 失败\n")
# 读取 Excel 前两列
def read_all_credentials_from_file(excel_path):
    try:
        df = pd.read_excel(excel_path, header=None)
        accounts = []
        for _, row in df.iterrows():
            username = str(row[0]).strip()
            password = str(row[1]).strip()
            if username and password:
                accounts.append((username, password))
        return accounts
    except Exception as e:
        print(f"[错误] 读取Excel文件 {excel_path} 出错: {e}")
        return []
def wait_for_user(browser_id=None, username=None, message="请点击确定继续操作"):
    """
    弹出Windows提示框，带提示音，并置于所有窗口最顶层。
    :param browser_id: 浏览器编号
    :param username: 当前账号
    :param message: 额外提示信息
    """
    # 拼接提示信息
    title_info = []
    if browser_id is not None:
        title_info.append(f"浏览器 {browser_id}")
    if username is not None:
        title_info.append(f"账号 {username}")
    prefix = " - ".join(title_info)

    full_message = f"{prefix}\n\n{message}" if prefix else message

    # 播放提示音
    winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)

    # 弹出消息框（阻塞直到点击“确定”）
    messagebox.showinfo("提示", full_message, parent=root)
# 每个浏览器进程循环多个账号
def browser_worker(account_list, browser_id):
    print(f"[启动] 浏览器 {browser_id} 启动中...")

    options = webdriver.ChromeOptions()
    options.add_argument("--incognito")
    options.add_argument("--start-maximized")  # 全屏
      # 无痕模式

    # 浏览器只开一次，无需为每个账号创建 profile
    driver = webdriver.Chrome(
        service=Service("chromedriver-win64/chromedriver.exe"),
        options=options
    )
    driver.maximize_window() 
    driver.get("https://www.autoengine.com")

    try:
        
       

        def click_button(button_locator):
                        try:
                            print("尝试点击第一个关闭按钮...")
                            close_btn1 = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "span.arco-modal-close-icon"))
                )   
                            close_btn1.click()
                            print("[成功] 已点击第一个关闭按钮")
                            button_locator.click()
                        except Exception:
                            print("[提示] 第一个按钮未找到，尝试第二个...")

                # 再    尝试点击第二个按钮
                        try:
                            print("尝试点击第二个关闭按钮...")
                            close_btn2 = WebDriverWait(driver, 3).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, "div.close-button"))
                        )
                            close_btn2.click()
                            print("[成功] 已点击第二个关闭按钮") 
                            button_locator.click()
                        except Exception as e:
                            print("[失败] 未找到关闭按钮，可能没有弹窗:", e)

        driver.get("https://www.autoengine.com")

        
        for username, password in account_list:
                              
                    try:
                                # 等待页面加载
                               
                                print("\n[步骤4] 等待页面加载完成...")
                                wait = WebDriverWait(driver, 20)
                                print("页面已加载完成")
                    
                                # 找到并填写用户名输入框
                                print("\n[步骤5] 正在填写手机号...")
                                print("寻找手机号输入框...")
                                username_input = wait.until(EC.presence_of_element_located((By.NAME, "mobile")))
                                print("清除输入框...")
                                username_input.clear()
                                print(f"输入手机号: {username}")
                                username_input.send_keys(username)
                                print("\n[成功] 手机号已填写完成")
                    
                                # 步骤6: 切换到“密码登录”
                                print("\n[步骤6] 正在切换到密码登录...")
                                try:
                                    switch_button = wait.until(
                                    EC.element_to_be_clickable((By.XPATH, "//div[contains(@class,'account-center-switch-button') and contains(text(),'密码登录')]"))
                                )   
                                    switch_button.click()
                                    print("\n[成功] 已切换到密码登录模式")
                                except Exception as e:
                                    print(f"[警告] 切换密码登录失败: {e}")
                    
                    
                    # 找    到并填写密码输入框
                                time.sleep(1) 
                                print("\n[步骤6] 正在填写密码...")
                                print("寻找密码输入框...")
                                password_input = wait.until(EC.presence_of_element_located((By.NAME, "password")))
                                print("清除输入框...")
                                password_input.clear()
                                print("输入密码: ******")
                                password_input.send_keys(password)
                                print("\n[成功] 密码已填写完成")
                    
                    
                    
                                print("\n[步骤7] 正在勾选同意协议...")
                    
                                try:
                            # 定位到外层 div（而不是 g/svg/use）
                                    agreement_checkbox = wait.until(
                                        EC.element_to_be_clickable((By.XPATH, "//div[contains(@class,'account-center-agreement-check')]"))
                                )
                            # 如果没勾选则点击
                                    if "checked" not in agreement_checkbox.get_attribute("class"):
                                        print("点击复选框...")
                                        agreement_checkbox.click()
                                        print("\n[成功] 已勾选同意协议")
                                    else:
                                        print("\n[提示] 同意协议已经被勾选")
                                except Exception as e:
                                    print(f"\n[警告] 勾选同意协议时出错: {e}")
                    
                    
                    
                    
                                print("\n[步骤8] 正在点击登录按钮...")
                    
                                try:
                                    print("寻找登录按钮...")
                            # 等待带有 active 类且文字为“登录”的按钮可点击
                                    login_button = wait.until(
                                        EC.element_to_be_clickable((
                                            By.XPATH, "//button[contains(@class,'account-center-action-button') and contains(@class,'active') and contains(text(),'登录')]"
                                ))
                            )
                                    print("点击登录按钮...")
                                    login_button.click()
                                    print("\n[成功] 已点击登录按钮")
                                except Exception as e:
                                    print(f"\n[警告] 点击登录按钮时出错: {e}")
                    
                                
                    
                                
                    
                    
                                # 等待登录成功
                                
                                
                                # 登录后的操作可以在这里添加
                                print("\n[步骤9] 正在点击【懂车帝营销】按钮...")
                                try:
                                    print("等待主页跳转...")
                                    WebDriverWait(driver, 999).until(
                                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'top-nav-item')]/span[contains(text(),'懂车帝营销')]"))
                        )
                                    print("主页加载完成，准备点击【懂车帝营销】按钮")
                                    marketing_button = wait.until(
                                        EC.element_to_be_clickable((By.XPATH, "//div[contains(@class,'top-nav-item')]/span[contains(text(),'懂车帝营销')]"))
                                    )
                                    marketing_button.click()
                                    print("\n[成功] 已点击【懂车帝营销】按钮")
                                except Exception as e:
                                
                                    try:click_button(marketing_button)
                                    except:
                                        wait_for_user(browser_id=browser_id, username=username, message="请点击【懂车帝营销】按钮")
                                        print("用户已点击确定，继续执行下一步操作")
                    
                                
                    
                    
                    
                                # 步骤10: 点击文章发布
                    
                                print("\n[步骤10] 正在点击【文章发布】按钮...")
                                try:
                                    article_button = wait.until(
                                        EC.element_to_be_clickable((By.XPATH, "//div[@role='link' and .//span[contains(text(),'文章发布')]]"))
                                    )
                                    article_button.click()
                                    print("\n[成功] 已点击【文章发布】按钮")
                                except Exception as e:
                                    try:click_button(article_button)
                                    except:
                                        wait_for_user(browser_id=browser_id, username=username, message="请点击【文章发布】按钮")
                                        print("用户已点击确定，继续执行下一步操作")
                                
                    
                    
                               
                    
                    
                    
                    # 步骤11    : 点击立即发布
                                time.sleep(2)
                                print("\n[步骤11] 正在点击【立即发布】按钮...")
                                try:
                                        publish_button = wait.until(
                                            EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(),'立即发布')]]"))
                                        )
                                        publish_button.click()
                                        print("\n[成功] 已点击【立即发布】按钮")
                                except Exception as e:
                                        try:click_button(publish_button)
                                        except:
                                            wait_for_user(browser_id=browser_id, username=username, message="请点击【立即发布】按钮")
                                            print("用户已点击确定，继续执行下一步操作")
                    
                                print("等待跳转到文章编辑页面...")
                                try:
                                    WebDriverWait(driver, 30).until(
                                        EC.presence_of_element_located((By.CSS_SELECTOR, "textarea[placeholder='请输入文章标题（5-30个汉字）']"))
                                    )
                                    print("已进入文章编辑页面")
                                except Exception as e:
                                    try:
                                        publish_button = wait.until(
                                            EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(),'立即发布')]]"))
                                        )
                                        publish_button.click()
                                    except:
                                        wait_for_user(browser_id=browser_id, username=username, message="请确认已进入文章编辑页面后点击确定")
                                try:
                                        title_box = WebDriverWait(driver, 5).until(
                                        EC.presence_of_element_located((By.CSS_SELECTOR, "textarea[placeholder='请输入文章标题（5-30个汉字）']"))
                            )
                                        title_box.clear()
                                        title_box.send_keys("岚图助力2025世界斯诺克公开赛成功举办")
                                        print("[成功] 已输入标题")
                                except Exception as e:
                                        print("[失败] 找不到标题输入框:", e)
                                        wait_for_user(browser_id=browser_id, username=username, message="请完成输入标题，点击确定后继续操作")
                    
                    
                                # 找到正文编辑器
                                body_box = WebDriverWait(driver, 5).until(
                                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.ProseMirror[contenteditable='true']"))
                    )   
                                body_box.click()  # 激活编辑器
                                time.sleep(0.5)
                    
                    # 模    拟粘贴（Ctrl+V）
                                body_box.send_keys(Keys.CONTROL, 'v')
                                time.sleep(1)  # 等待内容渲染完成
                                print("[成功] 已从剪贴板粘贴文章内容")
                                
                                try:
                            # 等待元素出现并可点击
                                    cover_add_button = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((By.CLASS_NAME, "article-cover-add"))
                            )
                    
                            # 滚动到可见位置
                                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", cover_add_button)
                                    time.sleep(0.5)  # 等待滚动完成
                    
                            # 如果有遮挡浮层，可以隐藏
                                    try:
                                        footer_tip = driver.find_element(By.CLASS_NAME, "footer-tip")
                                        driver.execute_script("arguments[0].style.display='none';", footer_tip)
                                    except:
                                        pass    
                                    
                            # 点击按钮
                                    cover_add_button.click()
                                    print("[成功] 已点击【添加封面】按钮")
                    
                                except Exception as e:
                                    print("[失败] 点击【添加封面】按钮出错:", e)
                    
                    
                                # 图片路径（本地绝对路径）
                                image_path = "C:/Users/18771/Desktop/k/dream.jpg"
                    
                                try:
                            # 找到隐藏的 input[type=file]
                                    file_input = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
                            )
                            # 直接发送图片路径
                                    file_input.send_keys(image_path)
                                    print("[成功] 已上传图片:", image_path)
                    
                            # 可选：等待上传完成，页面可能显示上传缩略图
                                    time.sleep(1)
                                except Exception as e:
                                    print("[失败] 上传图片出错:", e)
                    
                    
                                print("\n[步骤13] 正在点击【确定】按钮...")
                                try:
                                        confirm_button = WebDriverWait(driver, 5).until(
                                        EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='确定']]"))
                            )
                                        time.sleep(3)  # 等待按钮完全加载
                                        confirm_button.click()
                                        print("[成功] 已点击【确定】按钮")
                                        
                    
                                except Exception as e:
                                        wait_for_user(browser_id=browser_id, username=username, message="请上传图片并点击确定按钮，点击确定后继续操作")
                                        print(f"[警告] 点击【确定】按钮时出错: {e}")
                    
                                
                                try:
                                    publish_button = wait.until(EC.element_to_be_clickable(
                                        (By.XPATH, "//button[span[text()='发布']]")
                    ))
                                    publish_button.click()
                    
                                    print("[成功] 已点击【发布】按钮")
                    
                                except Exception as e:
                                    wait_for_user(browser_id=browser_id, username=username, message="请完成点击发布按钮，点击确定后继续操作")
                                    print(f"[失败] 点击【发布】按钮出错: {e}")
                              
                    # 滚    动到可见位置
                                try:
                                   
                                    print("[步骤] 准备退出登录...")
                       # 等待页面跳转完成（URL或加载状态改变即可）
                                    print("[步骤] 等待页面跳转完成...")
                                    WebDriverWait(driver, 20).until(
                                        lambda d: d.execute_script("return document.readyState") == "complete"
                                    )
                                    time.sleep(1)  # 可加短暂延时确保动画等完成
                                    print("[步骤] 页面已跳转完成，准备退出登录...")
                                    write_result(username, True)
                                    
                                    # 1. 定位账号栏（悬停用）
                                    time.sleep(2)
                                    account_area = wait.until(
                                        EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'avatar-wrap')]"))
                                    )
                    
                                    # 2. 鼠标悬停在账号栏
                                    actions = ActionChains(driver)
                                    actions.move_to_element(account_area).perform()
                                    print("[成功] 已悬停在账号栏，等待菜单出现...")
                    
                                    # 3. 等待并点击退出登录按钮
                                    logout_button = wait.until(
                                        EC.element_to_be_clickable((By.XPATH, "//button[span[text()='退出登录']]"))
                                    )
                                    logout_button.click()
                                    print("[成功] 已点击退出登录按钮")
                    
                                except Exception as e:
                                    print(f"[错误] 退出登录失败: {e}")
                                    wait_for_user(browser_id=browser_id, username=username, message="请点击退出登录按钮，点击确定后继续操作")
                                    
                             
                                
                    
                    
                    except Exception as e:
                                    write_result(username, False)
                                    print(f"[错误] 账号 {username} 执行过程中出错: {e}")
                                            # 尝试恢复浏览器
                                    try:
                                        print("[提示] 尝试重新载入网页...")
                                        driver.get("https://www.autoengine.com")
                                        wait = WebDriverWait(driver, 20)
                                        print("[成功] 网页已重新载入，准备跳过该账号。")
                                    except Exception as reload_err:
                                        print(f"[错误] 网页重新载入失败: {reload_err}")
                                        # 如果重载失败，可以选择退出再重开 driver（可选）
                            
                                    # 跳过当前账号，继续下一个
                                    continue
    except Exception as e:
                        print(f"[{username}] 浏览器启动失败: {e}")
    finally:
                    if driver:
                        print(f"[{username}] 浏览器关闭")
                        driver.quit()      

if __name__ == "__main__":
    # 定义三个账号表格，每个浏览器对应一个表格
    excel_files = [
        os.path.join(current_dir, "K1.xlsx"),
        os.path.join(current_dir, "K2.xlsx"),
        os.path.join(current_dir, "K3.xlsx")
    ]

    processes = []

    for i, excel_path in enumerate(excel_files):
        accounts = read_all_credentials_from_file(excel_path)  # 读取单个表格的函数
        if not accounts:
            print(f"[浏览器{i+1}] 表格 {excel_path} 没有账号，跳过")
            continue

        p = multiprocessing.Process(target=browser_worker, args=(accounts, i+1))
        p.start()
        processes.append(p)

    # 等待所有浏览器进程结束
    for p in processes:
        p.join()
