#@title Python爬取HKEX Announcement (with GUI) 
#HKEX annoucement list by victorsuen
import requests
import json
import os
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
from tkcalendar import DateEntry
from bs4 import BeautifulSoup
import configparser

class HKEXDownloader:
    def __init__(self):
        self.headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.67 Safari/537.36"
        }
        
    def get_stockid(self, stockcode):
        try:
            url = f'https://www1.hkexnews.hk/search/prefix.do?&callback=callback&lang=ZH&type=A&name={stockcode}&market=SEHK&_=1653821865437'
            response = requests.get(url, headers=self.headers)
            
            # 處理 JSONP 響應
            data = response.text[9:-4]
            print(data)
            data2 = json.loads(data)
            stockid = data2["stockInfo"][0]['stockId']
            return stockid
            
        except Exception as e:
            raise Exception(f"獲取股票資訊時發生錯誤: {str(e)}")

    def get_announcement_list(self, stockcode, start_date, end_date, row_number, language, search_keywords):
        try:
            # 獲取 stockId
            stockid = self.get_stockid(stockcode)
            print("StockID = " + str(stockid))
            
            # 使用自定義關鍵字
            search_keyword = search_keywords[-1] if search_keywords else ""
            
            # 準備日期格式
            start_date_str = start_date.strftime("%Y%m%d")
            end_date_str = end_date.strftime("%Y%m%d")
            
            # 構建完整URL
            url = f'https://www1.hkexnews.hk/search/titleSearchServlet.do?sortDir=0&sortByOptions=DateTime&category=0&market=SEHK&stockId={stockid}&documentType=-1&fromDate={start_date_str}&toDate={end_date_str}&title={search_keyword}&searchType=0&t1code=-2&t2Gcode=-2&t2code=-2&rowRange={row_number}&lang={language}'
            print(f"搜尋URL: {url}")
            
            # 發送請求
            response = requests.get(url, headers=self.headers)
            
            if response.status_code != 200:
                raise Exception(f"無法連接到港交所網站 (狀態碼: {response.status_code})")

            # 處理回應數據
            try:
                data = response.text
                # 清理回應文本，完全按照原始代碼的方式
                data = data.replace('"[{','[{').replace('}]"','}]').replace('\\',"").replace('u2013',"-").replace('u0026',"-")
                data5 = json.loads(data)
                
                if not data5 or 'result' not in data5 or not data5['result']:
                    raise Exception("未找到符合條件的公告")
                
                announcements = []
                for item in data5['result']:
                    try:
                        # 獲取標題
                        title = item['TITLE']
                        title = title.replace('/', "-")
                        
                        # 獲取 PDF 連結
                        pdflink = item['FILE_LINK']
                        pdf_link = "https://www1.hkexnews.hk" + pdflink
                        
                        # 處理日期 - 完全按照原始代碼的方式
                        anndate = item['DATE_TIME']
                        anndate = anndate[:10].replace('/', "-")
                        date_object = datetime.strptime(anndate, "%d-%m-%Y")
                        formatted_date = date_object.strftime("%Y-%m-%d")
                        print(formatted_date)
                        
                        announcements.append({
                            'date': formatted_date,
                            'title': title,
                            'link': pdf_link
                        })
                        
                    except Exception as e:
                        print(f"處理公告項目時發生錯誤: {str(e)}")
                        continue

                if not announcements:
                    raise Exception("未找到符合條件的公告")

                return announcements
                
            except json.JSONDecodeError as e:
                print(f"JSON 解析錯誤: {str(e)}")
                print(f"回應內容: {response.text[:200]}")
                raise Exception("伺服器響應格式無效")
            
        except Exception as e:
            raise Exception(f"獲取公告列表時發生錯誤: {str(e)}")

    def download_announcements(self, stockcode, start_date, end_date, row_number, language, search_keywords, save_path, filename_length):
        try:
            announcements = self.get_announcement_list(
                stockcode, start_date, end_date, row_number,
                language, search_keywords
            )
            
            if not announcements:
                raise Exception("未找到符合條件的公告")
            
            savepath = os.path.join(save_path, 'HKEX', stockcode)
            if not os.path.exists(savepath):
                os.makedirs(savepath)
            
            download_count = 0
            for ann in announcements:
                try:
                    response = requests.get(ann['link'], headers=self.headers)
                    if response.status_code == 200:
                        filepath = os.path.join(
                            savepath, 
                            f"{ann['date']}_{stockcode}-{ann['title'][:filename_length]}.pdf"
                        )
                        
                        with open(filepath, 'wb') as f:
                            f.write(response.content)
                            print(f"已下載: {filepath}")
                            download_count += 1
                            
                except Exception as e:
                    print(f"下載文件時發生錯誤: {str(e)}")
                    continue
            
            return savepath, download_count
            
        except Exception as e:
            raise Exception(f"下載過程中出現錯誤: {str(e)}")

class HKEXDownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("港交所公告下載器")
        self.root.geometry("720x750")  # 調整窗口高度
        self.downloader = HKEXDownloader()
        self.default_row_number = 500
        
        # 初始化配置相關的變量
        self.config = configparser.ConfigParser()
        self.config_file = 'hkex_downloader.ini'
        self.remember_path_var = tk.BooleanVar()
        
        # 設置樣式
        style = ttk.Style()
        style.configure('TLabel', font=('Microsoft YaHei UI', 10))
        style.configure('TButton', font=('Microsoft YaHei UI', 10))
        style.configure('TCheckbutton', font=('Microsoft YaHei UI', 10))
        style.configure('TEntry', font=('Microsoft YaHei UI', 10))
        style.configure('Header.TLabel', font=('Microsoft YaHei UI', 12, 'bold'))
        style.configure('Author.TLabel', font=('Microsoft YaHei UI', 8))
        
        # 創建主框架並設置權重
        main_frame = ttk.Frame(root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        root.grid_rowconfigure(0, weight=1)
        root.grid_columnconfigure(0, weight=1)
        
        # 標題和作者信息
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, columnspan=3, pady=(0, 10))  # 減少上下間距
        
        header = ttk.Label(header_frame, text="港交所公告下載工具", style='Header.TLabel')
        header.grid(row=0, column=0, columnspan=3)
        
        author = ttk.Label(header_frame, text="作者：Victor Suen    版本：1.0", style='Author.TLabel')
        author.grid(row=1, column=0, columnspan=3, pady=(2,0))  # 減少上下間距
        
        # 股票代碼
        stock_frame = ttk.Frame(main_frame)
        stock_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=3)  # 減少上下間距
        
        ttk.Label(stock_frame, text="股票代碼:").grid(row=0, column=0, sticky=tk.W)
        self.stockcode = ttk.Entry(stock_frame, width=20)
        self.stockcode.grid(row=0, column=1, sticky=tk.W, padx=5)
        self.stockcode.insert(0, "00081")
        
        ttk.Label(stock_frame, text="(請輸入五位數字，如：00001)", foreground='gray').grid(row=0, column=2, sticky=tk.W, padx=5)
        
        # 日期範圍框架
        date_frame = ttk.LabelFrame(main_frame, text="日期範圍", padding="5")  # 減少內邊距
        date_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)  # 減少上下間距
        
        # 開始日期
        ttk.Label(date_frame, text="開始日期:").grid(row=0, column=0, sticky=tk.W)
        self.start_date = DateEntry(date_frame, width=15, 
                                  date_pattern='yyyy/mm/dd',
                                  year=2024, month=1, day=1,
                                  background='darkblue',
                                  foreground='white',
                                  borderwidth=2)
        self.start_date.grid(row=0, column=1, padx=5)
        
        # 結束日期
        ttk.Label(date_frame, text="結束日期:").grid(row=0, column=2, sticky=tk.W, padx=(20,0))
        self.end_date = DateEntry(date_frame, width=15,
                                date_pattern='yyyy/mm/dd',
                                background='darkblue',
                                foreground='white',
                                borderwidth=2)
        self.end_date.grid(row=0, column=3, padx=5)
        
        # 搜尋關鍵字框架
        keyword_frame = ttk.LabelFrame(main_frame, text="搜尋關鍵字", padding="5")  # 減少內邊距
        keyword_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)  # 減少上下間距
        
        # 自定義關鍵字
        ttk.Label(keyword_frame, text="自定義:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.custom_keyword = ttk.Entry(keyword_frame, width=40)
        self.custom_keyword.grid(row=0, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # 設置框架
        settings_frame = ttk.LabelFrame(main_frame, text="設置", padding="5")  # 減少內邊距
        settings_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)  # 減少上下間距
        
        # 第一行：語言和文件名長度
        lang_frame = ttk.Frame(settings_frame)
        lang_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        
        ttk.Label(lang_frame, text="語言:").grid(row=0, column=0, sticky=tk.W)
        self.language = ttk.Combobox(lang_frame, values=["中文", "English"], width=10, state='readonly')
        self.language.grid(row=0, column=1, sticky=tk.W, padx=(5, 20))
        self.language.set("中文")
        
        ttk.Label(lang_frame, text="文件名長度:").grid(row=0, column=2, sticky=tk.W)
        self.filename_length = ttk.Entry(lang_frame, width=8)
        self.filename_length.grid(row=0, column=3, sticky=tk.W, padx=5)
        self.filename_length.insert(0, "220")
        
        # 第二行：保存路徑
        path_frame = ttk.Frame(settings_frame)
        path_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        ttk.Label(path_frame, text="保存路徑:").grid(row=0, column=0, sticky=tk.W)
        self.save_path = ttk.Entry(path_frame, width=45)
        self.save_path.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        
        # 載入配置
        self.load_config()
        
        # 從配置文件讀取保存路徑
        saved_path = self.config.get('Settings', 'save_path', fallback='')
        if saved_path and os.path.exists(saved_path):
            self.save_path.insert(0, saved_path)
        else:
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            self.save_path.insert(0, desktop_path)
        
        # 選擇路徑按鈕
        ttk.Button(path_frame, text="選擇路徑", command=self.choose_directory).grid(
            row=0, column=2, padx=(5, 0), sticky=tk.E
        )
        
        # 第三行：預設路徑複選框
        self.remember_path_var = tk.BooleanVar()
        self.remember_path_var.set(self.config.getboolean('Settings', 'remember_path', fallback=True))
        self.remember_path_cb = ttk.Checkbutton(settings_frame, text="預設路徑", 
                                               variable=self.remember_path_var,
                                               command=self.save_config)
        self.remember_path_cb.grid(row=2, column=0, sticky=tk.W, pady=(5, 0))
        
        # 按鈕框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=5)  # 減少間距
        
        # 下載按鈕
        self.download_btn = ttk.Button(button_frame, text="開始下載", command=self.start_download, width=20)
        self.download_btn.grid(row=0, column=0, padx=10)
        
        # 打開資料夾按鈕
        self.open_folder_btn = ttk.Button(button_frame, text="打開下載資料夾", command=self.open_folder, width=20)
        self.open_folder_btn.grid(row=0, column=1, padx=10)
        self.open_folder_btn.state(['disabled'])
        
        # 狀態標籤
        self.status_label = ttk.Label(main_frame, text="", font=('Microsoft YaHei UI', 9))
        self.status_label.grid(row=6, column=0, columnspan=3, pady=2)  # 減少間距
        
        # 新增進度條框架
        progress_frame = ttk.LabelFrame(main_frame, text="下載進度", padding="5")
        progress_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # 進度條
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            length=600,  # 調整進度條長度與下方文本框一致
            mode='determinate'
        )
        self.progress_bar.grid(row=0, column=0, columnspan=2, pady=5, sticky=(tk.W, tk.E))
        
        # 進度標籤
        self.progress_label = ttk.Label(progress_frame, text="準備下載...")
        self.progress_label.grid(row=1, column=0, columnspan=2, pady=5)
        
        # 下載信息文本框
        info_frame = ttk.LabelFrame(main_frame, text="下載信息", padding="5")
        info_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # 配置info_frame的行列權重
        info_frame.grid_rowconfigure(0, weight=1)
        info_frame.grid_columnconfigure(0, weight=1)
        
        # 使用Text widget來顯示下載信息
        self.info_text = tk.Text(info_frame, height=8, width=70, font=('Microsoft YaHei UI', 9))
        self.info_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 添加滾動條
        scrollbar = ttk.Scrollbar(info_frame, orient="vertical", command=self.info_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.info_text.configure(yscrollcommand=scrollbar.set)
        
        # 配置main_frame的列權重
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(8, weight=1)  # 讓下載信息區域可以擴展
        
        # 保存最後一次下載的路徑
        self.last_download_path = None
        
    def load_config(self):
        """載入配置文件"""
        if os.path.exists(self.config_file):
            self.config.read(self.config_file)
        if 'Settings' not in self.config:
            self.config['Settings'] = {}
            self.config['Settings']['remember_path'] = 'true'
            self.config['Settings']['save_path'] = os.path.join(os.path.expanduser("~"), "Desktop")
            with open(self.config_file, 'w') as configfile:
                self.config.write(configfile)

    def save_config(self):
        """保存配置到文件"""
        if self.remember_path_var.get():
            self.config['Settings']['save_path'] = self.save_path.get()
        self.config['Settings']['remember_path'] = str(self.remember_path_var.get()).lower()
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)

    def choose_directory(self):
        directory = filedialog.askdirectory(initialdir=self.save_path.get())
        if directory:
            self.save_path.delete(0, tk.END)
            self.save_path.insert(0, directory)
            if self.remember_path_var.get():
                self.save_config()

    def open_folder(self):
        if self.last_download_path and os.path.exists(self.last_download_path):
            if os.name == 'nt':  # Windows
                os.startfile(self.last_download_path)
            else:  # macOS and Linux
                subprocess.run(['xdg-open' if os.name == 'posix' else 'open', self.last_download_path])
        else:
            messagebox.showwarning("警告", "尚未下載文件或路徑不存在")
            
    def update_progress(self, current, total, message=""):
        """更新進度條和進度信息"""
        progress = (current / total) * 100
        self.progress_var.set(progress)
        self.progress_label.config(text=f"{message} ({current}/{total})")
        self.root.update()

    def add_info(self, message):
        """添加信息到文本框"""
        self.info_text.insert(tk.END, message + "\n")
        self.info_text.see(tk.END)  # 自動滾動到最新信息
        self.root.update()

    def start_download(self):
        try:
            # 清空信息顯示
            self.info_text.delete(1.0, tk.END)
            self.progress_var.set(0)
            self.progress_label.config(text="準備下載...")
            
            # 禁用下載按鈕
            self.download_btn.state(['disabled'])
            self.status_label.config(text="正在搜尋公告...", foreground='blue')
            self.root.update()
            
            # 獲取輸入值
            stockcode = self.stockcode.get().strip()
            start_date = self.start_date.get_date()
            end_date = self.end_date.get_date()
            custom_keyword = self.custom_keyword.get().strip()
            language = "zh" if self.language.get() == "中文" else "en"
            save_path = self.save_path.get().strip()
            filename_length = self.filename_length.get().strip()
            
            # 驗證輸入
            if not stockcode:
                raise ValueError("請輸入股票代碼")
            
            if not stockcode.isdigit() or len(stockcode) != 5:
                raise ValueError("股票代碼必須為5位數字")
            
            try:
                filename_length = int(filename_length)
                if filename_length <= 0:
                    raise ValueError
            except ValueError:
                raise ValueError("文件名長度必須為正整數")
            
            if start_date > end_date:
                raise ValueError("開始日期不能晚於結束日期")
            
            # 使用自定義關鍵字
            search_keywords = [custom_keyword] if custom_keyword else []
            
            # 獲取公告列表
            self.add_info("正在搜尋符合條件的公告...")
            announcements = self.downloader.get_announcement_list(
                stockcode, start_date, end_date, self.default_row_number,
                language, search_keywords
            )
            
            self.add_info(f"找到 {len(announcements)} 個符合條件的公告")
            
            # 創建保存目錄
            savepath = os.path.join(save_path, 'HKEX', stockcode)
            if not os.path.exists(savepath):
                os.makedirs(savepath)
            
            # 下載公告
            download_count = 0
            total_announcements = len(announcements)
            
            for i, ann in enumerate(announcements, 1):
                try:
                    self.update_progress(i, total_announcements, f"正在下載: {ann['title'][:50]}...")
                    self.add_info(f"正在下載 ({i}/{total_announcements}): {ann['title']}")
                    
                    response = requests.get(ann['link'], headers=self.downloader.headers)
                    if response.status_code == 200:
                        filepath = os.path.join(
                            savepath,
                            f"{ann['date']}_{stockcode}-{ann['title'][:filename_length]}.pdf"
                        )
                        
                        with open(filepath, 'wb') as f:
                            f.write(response.content)
                            download_count += 1
                            self.add_info(f"✓ 成功下載: {os.path.basename(filepath)}")
                            
                except Exception as e:
                    self.add_info(f"✗ 下載失敗: {ann['title']} - {str(e)}")
                    continue
            
            self.last_download_path = savepath
            self.open_folder_btn.state(['!disabled'])
            
            # 更新最終狀態
            self.progress_var.set(100)
            if download_count > 0:
                success_message = f"下載完成！成功下載 {download_count}/{total_announcements} 個文件到 {savepath}"
                self.status_label.config(text=success_message, foreground='green')
                self.progress_label.config(text=success_message)
                self.add_info("\n" + success_message)
            else:
                fail_message = "未能成功下載任何文件"
                self.status_label.config(text=fail_message, foreground='red')
                self.progress_label.config(text=fail_message)
                self.add_info("\n" + fail_message)
            
            # 下載完成後，如果選擇了記住路徑，則保存配置
            if self.remember_path_var.get():
                self.save_config()
            
        except Exception as e:
            error_message = f"錯誤：{str(e)}"
            self.status_label.config(text=error_message, foreground='red')
            self.progress_label.config(text="下載失敗")
            self.add_info("\n" + error_message)
            messagebox.showerror("錯誤", str(e))
            
        finally:
            # 重新啟用下載按鈕
            self.download_btn.state(['!disabled'])
            self.root.update()

def main():
    root = tk.Tk()
    app = HKEXDownloaderGUI(root)
    root.mainloop()

if __name__ == '__main__':
    main()