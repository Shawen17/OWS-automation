from ows_function import ows_login



option= webdriver.ChromeOptions()
#option.add_argument('headless')
browser = webdriver.Chrome(executable_path='C:\\Users\\user\\Anaconda3\\pkgs\\notebook-6.0.0-py37_0\\Lib\\site-packages\\notebook\\tests\\selenium\\chromedriver.exe', options=option)
browser.implicitly_wait(30)
ows_window=browser.current_window_handle
ows_login(browser,ows_window)
sleep(5)
browser.quit()