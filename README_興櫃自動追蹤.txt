興櫃自動追蹤 - 全自動版

你拿到的檔案：
1. 興櫃自動追蹤_全自動版.xlsx
2. esb_auto_monitor.py
3. .env.example

安裝套件
pip install requests pandas openpyxl python-dotenv lxml html5lib

設定方式
1. 把 .env.example 複製成 .env
2. 填入：
   LINE_CHANNEL_ACCESS_TOKEN=你的token
   LINE_USER_ID=你的user id
   WATCHLIST_CODES=7822,6879,4172
3. 把 Excel 與 Python 放同一資料夾
4. 執行：
   python esb_auto_monitor.py

程式會做什麼
- 讀取櫃買中心興櫃股票當日行情表
- 自動判斷強勢股（預設漲幅 >= 5%）
- 自動判斷爆量股（預設量比 >= 2x）
- 自動把結果寫回 Excel
- 自動累積歷史資料 CSV
- 自動發送 LINE 通知

Windows 自動排程
1. 開啟「工作排程器」
2. 建立基本工作
3. 觸發條件可設：
   09:10
   10:30
   13:35
4. 動作選「啟動程式」
5. 程式填入 python.exe
6. 引數填入 esb_auto_monitor.py
7. 起始位置填入你的資料夾路徑

Mac / Linux cron 範例
10 9 * * 1-5 /usr/bin/python3 /你的路徑/esb_auto_monitor.py
30 10 * * 1-5 /usr/bin/python3 /你的路徑/esb_auto_monitor.py
35 13 * * 1-5 /usr/bin/python3 /你的路徑/esb_auto_monitor.py

資料來源
- 櫃買中心 興櫃股票當日行情表
  https://www.tpex.org.tw/zh-tw/esb/trading/info/pricing.html
- 櫃買中心 興櫃股票成交統計
  https://www.tpex.org.tw/zh-tw/esb/psb/trading/statistics/day.html

提醒
- 若櫃買中心未來調整欄位名稱，程式中的 normalize_quote_df 可能要微調一次。
- 若你想加入「AI、生技、自選名單」分組推播，我可以再幫你擴充第二版。
