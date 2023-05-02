# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from ttkbootstrap.toast import ToastNotification
from ttkbootstrap.tableview import Tableview
from ttkbootstrap.validation import add_regex_validation
import sys, os, csv, configparser, pandas as pd, random, datetime, json, ctypes
from logging import getLogger, config
import threading, time, pywinauto
import pyautogui as pg, pyperclip
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import requests
from bs4 import BeautifulSoup
import schedule

INI_FILE = './config/setting.ini'
LOG_JSON_FILE = './config/logformart.json'
LOG_DIR = './log'
INSTAGRAM_LOGIN_URL = 'https://www.instagram.com/accounts/login/'

FONT_ENTRY_TYPE = ("游ゴシック", 10)
FONT_LABEL_TYPE = ("游ゴシック", 10)
FONT_RADIO_TYPE = ("游ゴシック", 10)


class Application(ttkb.Frame):
  def __init__(self, master_window):
    super().__init__(master_window, padding=(20,10,20,40))
    self.pack(fill=BOTH, expand=YES)
    self.window = master_window

    # 設定ファイル読込
    self.read_ini()

    # logger作成
    self.create_logger()

    # ログイン画面作成
    self.create_login_frame()

  def create_login_frame(self):
    self.login_frame = ttkb.Frame(self)
    self.login_frame.pack(fill=BOTH, expand=YES)

    self.login_name = ttkb.StringVar(value=self.login_ini['login_username'])
    self.login_pass = ttkb.StringVar(value=self.login_ini['login_password'])

    instruction = ttkb.Label(self.login_frame, text='ログイン画面', width=60, font=FONT_LABEL_TYPE)
    instruction.pack(fill=X, pady=10)

    self.create_entry(self.login_frame, 'ユーザ名','login_username', self.login_name)
    self.create_entry(self.login_frame, 'パスワード','login_password', self.login_pass, is_hidden=True)

    login_btn = ttkb.Button(master=self.login_frame,text='LOGIN', command=self.on_login,bootstyle=(INFO,OUTLINE),width=10 )
    login_btn.pack(side=RIGHT, padx=5, pady=(30,0))

    self.window.protocol("WM_DELETE_WINDOW", self.close_login_frame)

  def create_frame(self, master, label=None):
    form_container = ttkb.Frame(master)
    form_container.pack(fill=X, expand=YES, pady=10)
    if label:
      form_label = ttkb.Label(master=form_container, text=label, width=18, font=FONT_LABEL_TYPE)
      form_label.pack(side=LEFT, padx=12)
    return form_container

  def create_entry(self, master, label, entry_name,variable, is_hidden=False):
    form_field_container = self.create_frame(master, label)    
    if (is_hidden):
      form_input = ttkb.Entry(master=form_field_container, textvariable=variable, show='*',font=FONT_ENTRY_TYPE, name=entry_name)
    else:
      form_input = ttkb.Entry(master=form_field_container, textvariable=variable, font=FONT_ENTRY_TYPE, name=entry_name)      
    form_input.pack(side=LEFT, padx=5, fill=X, expand=YES)
    add_regex_validation(form_input, r'^[a-zA-Z0-9_]*$')
    return form_input

  def create_csv_form(self, master, label, entry_name, variable):
    form_field_container = self.create_frame(master, label)

    form_entry = ttkb.Entry(master=form_field_container, textvariable=variable, width=35, name=entry_name)
    form_entry.pack(side=LEFT, padx=5, fill=X, expand=YES)
    form_btn = ttkb.Button(master=form_field_container,bootstyle=(INFO,OUTLINE),text='参照', command=self.open_csv_dialog)
    form_btn.pack(side=LEFT, fill=X, expand=YES)
    # add_regex_validation(form_entry, r'^[a-zA-Z0-9_]*$')
    return form_entry

  def create_short_entry(self, master, label,unit_label, entry_name,variable):
    form_field_container = self.create_frame(master, label)    
    form_input = ttkb.Entry(master=form_field_container, textvariable=variable,font=FONT_ENTRY_TYPE, name=entry_name, width=8)
    form_input.pack(side=LEFT, padx=5, fill=X, expand=YES)
    form_unit_label = ttkb.Label(master=form_field_container, text=unit_label, width=45, font=FONT_LABEL_TYPE)
    form_unit_label.pack(side=LEFT, padx=6)
    # add_regex_validation(form_input, r'^[a-zA-Z0-9_]*$')
    return form_input

  def create_interval_radio(self, master, label, variable):
    form_field_container = self.create_frame(master, label)
    self.orderRadio01 = ttkb.Radiobutton(master = form_field_container, text='1分', value=1, variable=variable,bootstyle="info")
    self.orderRadio02 = ttkb.Radiobutton(master = form_field_container, text='10分', value=10, variable=variable,bootstyle="info")
    self.orderRadio03 = ttkb.Radiobutton(master = form_field_container,text='30分', value=30,variable=variable, bootstyle="info")
    self.orderRadio04 = ttkb.Radiobutton(master = form_field_container,text='60分', value=60,variable=variable, bootstyle="info")
    
    self.orderRadio01.pack(side=LEFT, padx=10)
    self.orderRadio02.pack(side=LEFT, padx=10)
    self.orderRadio03.pack(side=LEFT, padx=10)
    self.orderRadio04.pack(side=LEFT, padx=10)

  def create_next_arcticle_form(self, master, label, entry_name, variable):
    form_field_container = self.create_frame(master, label)    
    form_entry = ttkb.Entry(master=form_field_container, textvariable=variable, width=60, name=entry_name)
    form_entry.pack(side=LEFT, padx=5, fill=X, expand=YES)
    form_btn = ttkb.Button(master=form_field_container,bootstyle=(INFO,OUTLINE),text='更新',command=self.set_post_info)
    form_btn.pack(side=LEFT, fill=X, expand=YES)
    # add_regex_validation(form_entry, r'^[a-zA-Z0-9_]*$')
    return form_entry

  def create_main_btn(self, master):
    form_field_container = self.create_frame(master)
    self.stop_btn = ttkb.Button(master=form_field_container, text='停止', command=self.stop_post,  bootstyle=(DANGER,OUTLINE),width=10)
    self.immediate_btn = ttkb.Button(master=form_field_container, text='即時実行', command=self.immediate_exe, bootstyle=(INFO,OUTLINE),width=10)
    self.interval_btn = ttkb.Button(master=form_field_container, text='インターバル実行', command=self.interval_exe, bootstyle=(INFO,OUTLINE))

    self.stop_btn.pack(side=RIGHT, padx=5, pady=(30,0))
    self.immediate_btn.pack(side=RIGHT, padx=5, pady=(30,0))
    self.interval_btn.pack(side=RIGHT, padx=5, pady=(30,0))

  # def create_interval_merter(self, master, label):
  #   form_field_container = self.create_frame(master)    
  #   form_meter = ttkb.Meter(master=form_field_container, metersize=150,amounttotal=10,
  #                           amountused=0, metertype="full", subtext=label, )
  #   form_meter.pack(side=RIGHT, fill=X)

  #   # add_regex_validation(form_entry, r'^[a-zA-Z0-9_]*$')
  #   return form_meter

  def create_interval_gauge(self, master, label, variable):
    # form_field_container = self.create_frame(master) 

    form_field_container = ttkb.Frame(master, bootstyle='dark')
    form_field_container.propagate(False)
    form_field_container.pack(fill=X, expand=YES, pady=10)

    form_gauge = ttkb.Floodgauge(master=form_field_container, maximum=100, bootstyle='INFO', font=(None, 12, 'bold'), mask='Memory Used {}%', variable=variable)
    form_gauge = ttkb.Floodgauge(master=form_field_container, maximum=100, bootstyle='INFO', font=(None, 12, 'bold'), text='', variable=variable)
    form_gauge.pack()
    return form_gauge
  
  def create_main_frame(self):

    self.main_frame = ttkb.Frame(self)
    self.main_frame.pack(fill=BOTH, expand=YES)

    self.insta_user = ttkb.StringVar(value=self.post_ini['username'])
    self.insta_pass = ttkb.StringVar(value=self.post_ini['password'])
    self.csv_path = ttkb.StringVar(value=self.post_ini['csv_path'])
    self.post_interval = ttkb.IntVar(value=self.post_ini['post_interval'])
    self.next_article = ttkb.StringVar(value='')
    self.post_row_num = int(self.post_ini['post_row_number'])
    self.interval_var = ttkb.IntVar()

    instruction = ttkb.Label(self.main_frame, text='', width=90)
    instruction.pack(fill=X, pady=10)

    self.create_entry(self.main_frame, 'インスタユーザ名','username', self.insta_user)
    self.create_entry(self.main_frame, 'インスタパスワード','password', self.insta_pass, is_hidden=True)
    self.create_csv_form(self.main_frame, 'CSVファイル','csv_path', self.csv_path)
    self.create_interval_radio(self.main_frame, 'インターバル', self.post_interval)
    self.nex_article_entry = self.create_next_arcticle_form(self.main_frame, '記事タイトル','next_article', self.next_article)
    self.create_main_btn(self.main_frame)
    # self.interval_merter = self.create_interval_merter(self.main_frame,'test')
    self.interval_gauge =  self.create_interval_gauge(self.main_frame,'test', self.interval_var)

    self.window.protocol("WM_DELETE_WINDOW", self.close_main_frame)
    
  def close_login_frame(self):
    self.write_login_ini()
    self.login_frame.destroy()

  def on_login(self):
    self.close_login_frame()
    self.create_main_frame()

  def open_csv_dialog(self):
    fileTypes = (('テキストファイル', '*.csv'), ('すべてのファイル', '*.*'))
    if filename := filedialog.askopenfilename(filetypes=fileTypes):
      self.csv_path.set(filename)
  
  def read_ini(self):
    # iniファイル読込
    self.config = configparser.ConfigParser()
    self.config.read(INI_FILE)
    self.login_ini = self.config['login']
    self.post_ini = self.config['instagram']

  def write_login_ini(self):
    for wdg in self.get_all_widget(self.login_frame):
      if isinstance(wdg, ttkb.Entry) and wdg._name in self.login_ini:
        self.config.set('login', wdg._name, wdg.get())
    with open(INI_FILE, 'w') as f: self.config.write(f)

  def write_post_ini(self):
    for wdg in self.get_all_widget(self.main_frame):
      if isinstance(wdg, ttkb.Entry) and wdg._name in self.post_ini:
        self.config.set('instagram', wdg._name, wdg.get())
      self.config.set('instagram', 'post_row_number', str(self.post_row_num) if hasattr(self, 'post_row_num') else "0")
      self.config.set('instagram', 'post_interval', str(self.post_interval.get()) if hasattr(self, 'post_interval') else "1")
    with open(INI_FILE, 'w') as f: self.config.write(f)

  def get_all_widget(self, p_wid, finList=None) -> list:
    if finList is None: finList = []
    _list = p_wid.winfo_children()        
    for item in _list :
      finList.append(item)
      self.get_all_widget(item, finList)   
    return finList
  
  def create_logger(self):
    with open(LOG_JSON_FILE, 'r') as f: log_config = json.load(f)
    os.makedirs(LOG_DIR, exist_ok=True) 

    log_config["handlers"]["fileHandler"]["filename"] = os.path.join(LOG_DIR, f'{datetime.datetime.now().strftime("%Y%m%d")}.log')
    config.dictConfig(log_config)
    self.logger = getLogger(__name__)

  def get_browser(self):
      chromeOption = webdriver.ChromeOptions()
      # ポップアップ通知を有効にする。じゃなければ余計な画面が表示される。
      chromeOption.add_experimental_option("prefs", {"profile.default_content_setting_values.notifications": 1})
      # ブラウザ制御コメントを非表示化にする。
      chromeOption.add_experimental_option('excludeSwitches', ['enable-logging'])
      # WebDriverのテスト動作をTrueにする。
      chromeOption.use_chromium = True

      # ChromeDriverManager().install()のプログレスバーを非表示にする
      # exeだとプログレスバーを非表示にしないとエラーになる
      os.environ["WDM_PROGRESS_BAR"] = "0"

      # Chromeドライバの自動更新
      # ホームディレクトリに.wdmフォルダを作成しその中にドライバ配置する
      return webdriver.Chrome(ChromeDriverManager().install(), options=chromeOption)

  def change_all_widget_state(self, p_state):
    for wdg in self.get_all_widget(self.main_frame):
      if isinstance(wdg, ttkb.Button) or isinstance(wdg, ttkb.Entry) or isinstance(wdg, ttkb.Radiobutton): 
        wdg['state'] = p_state

  def change_widget_state(self, p_wdg, p_state):
    p_wdg['state'] = p_state

  def reset_all_widget(self):
    """全ウィジェットを初期状態に戻す(値は戻さない)
    """
    self.change_all_widget_state(tk.NORMAL)
    self.change_widget_state(self.stop_btn, tk.DISABLED)

  def before_exe(self):
    if len(self.next_article.get()) == 0:
      self.set_post_info()      
    self.change_all_widget_state(tk.DISABLED)
    self.change_widget_state(self.stop_btn, tk.NORMAL)
    self.is_post_stop, self.is_post_error = False, False

  def immediate_exe(self):
    self.before_exe()
    self.post_thread = threading.Thread(target=self.posting, name='post_thread', args=(self.post_info['ImageUrl'], self.post_info['Hashtag']))
    self.post_thread.start()

  def interval_exe(self):
    print(datetime.datetime.now().strftime('%H:%M:%S') + '【インターバル実行】')
    # print(self.interval_gauge['maximum'])
    self.interval_gauge['maximum'] = int((self.post_interval.get() * 60) / 0.05)
    # print(self.interval_gauge['maximum'])


    self.before_exe()
    # self.interval_gauge.start(self.post_interval.get() * 600)
    self.interval_gauge.start()
    schedule.every(self.post_interval.get()).minutes.do(self.interval_posting, self.post_info)
    self.monitor = threading.Thread(target=self.run_monitor, name='monitor_thread')
    self.monitor.start()

  def run_monitor(self):
    while not self.is_post_stop or not self.is_post_stop:
      schedule.run_pending()
      time.sleep(0.5)

  def interval_posting(self, p_post_info):
      print(datetime.datetime.now().strftime('%H:%M:%S') + '【開始】')
      self.interval_gauge.stop()
      self.interval_var.set(0)
      self.posting(p_post_info['ImageUrl'], p_post_info['Hashtag'], False)
      self.interval_gauge.start()
      # self.interval_gauge.start(self.post_interval.get() * 600)
      print(datetime.datetime.now().strftime('%H:%M:%S') + '【終了】')

  # def kill_thread(self):
  #   for thread in threading.enumerate():
  #       if thread != threading.main_thread():
  #           print(f"kill: {thread.name}")
  #           ctypes.pythonapi.PyThreadState_SetAsyncExc(thread.native_id, ctypes.py_object(SystemExit))


  def close_main_frame(self):
    self.write_post_ini()
    self.window.destroy()

  def stop_post(self):
    # self.kill_thread()
    self.is_post_stop = True

  def set_post_info(self, event=None):
    df = pd.read_csv(self.csv_path.get(), encoding='shift_jis')
    if self.post_row_num <= len(df):
      self.post_row_num = 1
    else:
      self.post_row_num += 1
    self.post_info = df.iloc[self.post_row_num]
    self.change_widget_state(self.nex_article_entry, tk.NORMAL)
    self.next_article.set(f'{self.post_row_num} {self.post_info["ItemNo"]} {self.post_info["ItemTitle"]}')
    self.change_widget_state(self.nex_article_entry,'readonly')

  def posting(self, p_img_url, p_img_caption, is_reset=True):
    try:      
      self.xpath_section = self.config['xpath']
      self.browser = self.get_browser()
      self.browser.get(INSTAGRAM_LOGIN_URL)
      self.auto_input(self.xpath_section['x_username'],'ログインユーザ入力', self.insta_user.get())
      self.auto_input(self.xpath_section['x_password'],'ログインパスワード入力', self.insta_pass.get())
      self.auto_click(self.xpath_section['x_login_enter'], 'ログインボタンクリック')
      self.auto_click(self.xpath_section['x_post_new'], '投稿ページ作成ボタンクリック')
      self.auto_click(self.xpath_section['x_open_select'], 'コンピュータから選択ボタンクリック')
      self.auto_select_img(p_img_url)
      self.auto_click(self.xpath_section['x_post_next'], '次へボタンクリック①')
      self.auto_click(self.xpath_section['x_post_next'], '次へボタンクリック②')
      self.auto_input(self.xpath_section['x_post_cap'], 'キャプション入力', p_img_caption, 3)
      self.auto_click(self.xpath_section['x_post_share'], 'シェアボタンクリック')
      self.auto_close_browser(5)
      self.write_post_ini()

      if is_reset:
        self.reset_all_widget()
      
    except Exception as ex:
      self.logger.error(ex)
    else:
      pass
      # self.logger.info('OK')

  def auto_input(self, p_xpath, p_info, p_value, p_wait_time=1):
    try:
      if self.is_post_stop or self.is_post_error: return
      lp_cnt = 30
      time.sleep(p_wait_time)
      for i in range(lp_cnt):
        if i == lp_cnt - 1:
          self.logger.error(f'処理名：{p_info}')
          self.logger.error(f'エラー内容：試行回数が上限になったので終了します')
          self.is_post_error = True
          return
        else:
          if self.is_post_stop or self.is_post_error: return
        
          elems =  self.browser.find_elements_by_xpath(p_xpath)
          if elems: elems[0].send_keys(p_value); return
          else: time.sleep(1)            
    except Exception as ex:
      self.is_post_error = True
      self.logger.error(f'処理名：{p_info}')
      self.logger.error(f'エラー内容：{ex}')
      return

  def auto_click(self, p_xpath, p_info, p_wait_time=1):
    try:
      if self.is_post_stop or self.is_post_error: return
      lp_cnt = 30
      time.sleep(p_wait_time)
      for i in range(lp_cnt):
        if i == lp_cnt - 1:
          self.logger.error(f'処理名：{p_info}')
          self.logger.error(f'エラー内容：試行回数が上限になったので終了します')
          self.is_post_error = True
          return
        else:
          if self.is_post_stop or self.is_post_error: return
          elems =  self.browser.find_elements_by_xpath(p_xpath)
          if elems: elems[0].click(); return
          else: time.sleep(1)     
    except Exception as ex:
      self.is_post_error = True
      self.logger.error(f'処理名：{p_info}')
      self.logger.error(f'エラー内容：{ex}')
      return

  def auto_select_img(self, filePath):
    try:
      if self.is_post_stop or self.is_post_error: return
      findWindow = lambda: pywinauto.findwindows.find_windows(title='開く')[0]
      dialog = pywinauto.timings.wait_until_passes(5, 1, findWindow)
      
      pwa_app = pywinauto.Application()
      pwa_app.connect(handle=dialog)
      window = pwa_app['開く']
      
      window.wait('ready')
      pywinauto.keyboard.send_keys("%N")
      edit = window.Edit4
      edit.set_focus()
      edit.set_text(filePath)
      button = window['開く(&O):']
      button.click()
      
    except Exception as ex:
      self.is_post_error = True
      self.logger.error(f'処理名：画像選択')
      self.logger.error(f'エラー内容：{ex}')
      return

  def auto_close_browser(self, p_wait_time=1):
    try:
      if (not self.is_post_stop) and (not self.is_post_error): time.sleep(p_wait_time)
      self.browser.quit()

    except Exception as ex:
        self.is_post_error = True
        print(f'{ex}')

def main():
  app = ttkb.Window(title='インスタグラム自動投稿', themename='superhero', resizable=(False,False))
  Application(app)
  app.mainloop()


if __name__ == "__main__":
    main()
