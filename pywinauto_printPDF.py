from pywinauto.application import Application
from pywinauto import Desktop
import configparser
import time

class SetUP():
    def __init__(self) -> None:
        pass

    def startUP_browser(self):
        self.config = configparser.ConfigParser()
        self.config.read('config.ini', encoding='utf-8')
        app_path = self.config['app_path']['chrome_path']
        Application(backend="uia").start(app_path)

    def get_account(self):
        user_id = self.config['user']['user_id']
        password = self.config['user']['password']
        return user_id, password
    
    def get_top_url(self):
        top_url = self.config['url_list']['top_url'].strip("'")
        return top_url

class WindousAPP():
    def __init__(self) :
        app = Desktop(backend="uia").windows()
        self.desktop_hwnd = app[0].parent()

    def searchElement(self, main_hwnd, search_text):
        found_element = None
        def searchChildElements(hwnd):
            nonlocal found_element
            if found_element is not None:
                return
            try:
                children = hwnd.children()
            except Exception:
                raise 
            for child in children:
                if search_text in str(child):
                    found_element = child
                    return
                # 子要素を再帰的に探索
                searchChildElements(child)
        # main_hwndから探索を開始
        searchChildElements(main_hwnd)
        return found_element

    def get_query(self, search_text, command=None, input_data=None):
        hwnd = self.searchElement(self.desktop_hwnd, search_text)
        if input_data:
            hwnd.set_text(input_data)
        if command:
            hwnd.type_keys(command)
        return hwnd
    
    def login(self, user_id, password):
        search_text = '聖教新聞（創価学会の機関紙）の公式サイト - Google Chrome'
        hwnd = self.get_query(search_text)
        search_text = "'ログイン', Hyperlink"
        hwnd = self.get_query(search_text)
        hwnd.click_input()

        time.sleep(3)
        search_text = 'ログイン - Google Chrome'
        hwnd = self.searchElement(self.desktop_hwnd, search_text)
        hwnd.type_keys("{TAB}")
        hwnd.type_keys(user_id)
        hwnd.type_keys("{TAB}")
        hwnd.type_keys(password)
        hwnd.type_keys("{TAB}")
        hwnd.type_keys("{TAB}")
        hwnd.type_keys("{ENTER}")

    def printing(self):
        search_text = '聖教新聞（創価学会の機関紙）の公式サイト - Google Chrome'
        hwnd = self.get_query(search_text)
        for _ in range(1):
            self.desktop_hwnd.type_keys("{PGDN}")
        time.sleep(1)
        search_text = "'紙面を見る', Hyperlink"
        hwnd = self.searchElement(hwnd, search_text)
        hwnd.click_input()
        time.sleep(3)

        search_text = '聖教新聞：紙面ビューア - Google Chrome'
        hwnd = self.get_query(search_text)
        search_text = "'印刷', Hyperlink"
        hwnd = self.searchElement(hwnd, search_text)
        hwnd.click_input()
        time.sleep(3)

        search_text = '聖教新聞：紙面ビューア - Google Chrome'
        hwnd = self.get_query(search_text)
        search_text = "'印刷', Hyperlink"
        hwnd = self.searchElement(hwnd, search_text)
        hwnd.click_input()
        time.sleep(5)

        self.desktop_hwnd.type_keys("{ENTER}")
        time.sleep(6)
        self.desktop_hwnd.type_keys("{ENTER}")


if __name__ == "__main__":
    con = SetUP()
    con.startUP_browser()
    user_id, password = con.get_account()

    search_text = "新しいタブ - Google Chrome"
    hwnd = WindousAPP().get_query(search_text)
    time.sleep(3)

    # サイト検索
    search_text = "'アドレス検索バー', Edit"
    hwnd = WindousAPP().get_query(search_text, command="{ENTER}", input_data=con.get_top_url())
    time.sleep(3)

    # ログイン
    try:
        WindousAPP().login(user_id, password)
    except Exception:
        print("ログイン済み")
    
    time.sleep(5)

    # 新聞印刷
    WindousAPP().printing()