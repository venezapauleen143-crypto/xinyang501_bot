"""
Claude Code 工具腳本 - 整合小牛馬所有技能
用法：python claude_tools.py <tool> [參數...]

tools:
  weather <城市>                查詢天氣
  stock <代號> [期間]           查詢股票（期間: 1d 5d 1mo 3mo 6mo 1y，預設 1mo）
  image <prompt> [文字]         生成圖片（prompt 用英文，文字為疊加中文）
  screenshot                    截圖存到桌面
  click <x> <y>                 點擊
  double_click <x> <y>          雙擊
  right_click <x> <y>           右鍵點擊
  move <x> <y>                  移動滑鼠
  type <文字>                   輸入文字
  press <按鍵>                  按下按鍵
  open <程式>                   開啟程式
  scroll <up|down> [格數]       滾動
  pos                           取得目前滑鼠座標
  bot_status                    查看 bot 是否在執行
  bot_restart                   重啟 bot
  schedule_list                 列出所有排程任務
  schedule_add <名稱> <時間HH:MM> <腳本路徑>   新增每日排程（並設定可喚醒）
  schedule_del <名稱>           刪除排程任務
  memory_save <chat_id> <內容>  儲存長期記憶
  memory_list <chat_id>         列出長期記憶
  memory_del <chat_id> <id>     刪除指定記憶
  vision [問題]                 截圖並讓 Claude 分析畫面內容
  find_image <圖片路徑>         在螢幕上找到指定圖片的位置
  browser <動作> [參數]         瀏覽器自動化（open/click/type/get_text/screenshot/goto/close）
  window_list                   列出所有視窗
  window_focus <標題關鍵字>     切換到指定視窗
  window_close <標題關鍵字>     關閉指定視窗
  window_min <標題關鍵字>       最小化視窗
  window_max <標題關鍵字>       最大化視窗
  hotkey <按鍵組合>             執行組合鍵（如 ctrl+c, alt+tab）
  clipboard_get                 讀取剪貼簿內容
  clipboard_set <內容>          寫入剪貼簿
  file_list <路徑>              列出資料夾內容
  file_read <路徑>              讀取文字檔案
  file_write <路徑> <內容>      寫入文字檔案
  file_delete <路徑>            刪除檔案或資料夾
  file_copy <來源> <目標>       複製檔案
  file_move <來源> <目標>       移動/重新命名檔案
  file_search <資料夾> <關鍵字> 搜尋檔案
  sysinfo                       查看 CPU/記憶體/磁碟使用率
  process_list                  列出所有執行中程序
  process_kill <名稱或PID>      結束程序
  notify <標題> <訊息>          發送 Windows 桌面通知
  tts <文字>                    文字轉語音朗讀
  record_start                  開始錄製滑鼠鍵盤動作
  record_stop                   停止錄製並儲存
  record_play <檔案>            重播錄製的動作
  email <收件人> <主旨> <內容>  發送 Email
  stt                           語音辨識（麥克風錄音轉文字）
  ocr [圖片路徑]                OCR 截圖或指定圖片，辨識文字
  workflow_run <json檔>         執行自動化工作流程
  workflow_save <名稱> <json>   儲存工作流程
  screen_watch <圖片> <指令>    監控螢幕出現指定圖片時執行指令
  monitors                      列出所有螢幕資訊
  zip <來源> <目標zip>          壓縮檔案或資料夾
  unzip <zip檔> <目標資料夾>    解壓縮
  download <網址> [儲存路徑]    下載檔案
  print_file <檔案路徑>         列印文件
  wifi_list                     列出可用 WiFi 網路
  wifi_connect <SSID> <密碼>    連線 WiFi
  screen_stream <秒數>          截圖串流（每秒截圖存桌面，持續N秒）
  wake_listen <關鍵字>          等待語音說出關鍵字後執行
  drag <x1> <y1> <x2> <y2>     拖曳從(x1,y1)到(x2,y2)
  right_menu <x> <y> <項目>     右鍵選單並選擇項目
  ai_plan <目標>                AI 自動規劃並執行多步驟任務
  clipboard_history             顯示剪貼簿歷史紀錄
  vdesktop <left|right|new>     切換虛擬桌面
  power <sleep|restart|shutdown> 電源管理
  bt_scan                       掃描藍牙裝置
  bt_connect <MAC位址>          連線藍牙裝置
  run_python <程式碼>           直接執行 Python 程式碼
  run_shell <指令>              執行 PowerShell 指令
  word_read <路徑>              讀取 Word 文件
  word_write <路徑> <內容>      寫入 Word 文件
  excel_read <路徑> [工作表]    讀取 Excel
  excel_write <路徑> <工作表> <資料JSON> 寫入 Excel
  pdf_read <路徑>               讀取 PDF 文字
  screen_diff <間隔秒> <區域>   偵測螢幕區域變化
  scrape <網址> <CSS選擇器>     爬取網頁指定內容
  img_edit <動作> <路徑> [參數] 圖片編輯（crop/resize/text/merge）
  gdrive_upload <檔案> <資料夾ID> 上傳到 Google Drive
  gdrive_download <檔案ID> <路徑> 從 Google Drive 下載
  db_query <資料庫路徑> <SQL>   執行 SQLite 查詢
  db_mysql <host> <db> <SQL>    執行 MySQL 查詢
  encrypt <檔案路徑> <密碼>     加密檔案
  decrypt <檔案路徑> <密碼>     解密檔案
  clipboard_watch <秒數>        監控剪貼簿變化
  qr_gen <內容> [路徑]          生成 QR Code
  qr_scan [圖片路徑]            掃描 QR Code（截圖或指定圖片）
  screen_record [秒數] [輸出路徑]  螢幕錄影存 mp4
  webcam [輸出路徑]             攝影機拍照
  translate <文字> [目標語言] [來源語言]  翻譯（目標預設 zh-TW）
  chart <類型> <資料JSON> [標題] [輸出路徑]  生成圖表（line/bar/pie）
  pptx_read <路徑>              讀取 PowerPoint
  pptx_create <路徑> <slides_json>  建立 PowerPoint
  api_call <方法> <URL> [headers_json] [body_json]  呼叫 REST API
  watchdog <程序名> <腳本路徑> [秒數]  守護程序，崩潰自動重啟
  ssh_run <host> <user> <password> <指令>  SSH 執行遠端指令
  sftp_upload <host> <user> <pass> <local> <remote>  SFTP 上傳
  sftp_download <host> <user> <pass> <remote> <local>  SFTP 下載
  net_ping <host> [次數]        Ping 網路主機
  net_traceroute <host>         路由追蹤
  net_portscan <host> [ports]   Port 掃描
  win_service <list|start|stop> [服務名]  Windows 服務管理
  pdf_merge <paths_json> <輸出路徑>  合併 PDF
  pdf_split <路徑> <輸出資料夾>  分割 PDF 每頁
  pdf_watermark <路徑> <文字> [輸出路徑]  PDF 加浮水印
  audio_convert <輸入> <輸出>   音訊格式轉換
  audio_trim <輸入> <起始ms> <結束ms> [輸出]  音訊剪輯
  discord_notify <webhook_url> <訊息>  Discord 推播
  line_notify <token> <訊息>    LINE Notify 推播
  disk_clean <list|clean>       磁碟暫存清理
  backup <來源> <目標資料夾>    備份壓縮
  registry_read <key_path> [value_name]  讀取登錄檔
  registry_write <key_path> <value_name> <value>  寫入登錄檔
  video_screenshot <路徑> [秒數] [輸出]  影片截取畫面
  video_trim <路徑> <起始秒> <結束秒> [輸出]  影片剪輯
  monitor_list                  列出所有螢幕資訊
  email_read <host> <user> <pass> [資料夾] [數量]  讀取 IMAP 郵件
  gcal_list [天數]              列出 Google Calendar 行程
  gcal_add <標題> <開始> <結束> [說明]  新增行程（時間格式 2026-04-13T10:00:00）
  global_hotkey <熱鍵> <指令> [秒數]  監聽全域快捷鍵
  git_op <動作> [repo] [參數]   Git 操作（status/log/pull/add/commit/push/diff）
  hw_monitor                    硬體監控（GPU/電池/溫度）
  report_gen <標題> <資料JSON> [輸出路徑]  生成 HTML 報告
  dropbox_upload <local> <remote> [token]  上傳到 Dropbox
  dropbox_download <remote> <local> [token]  從 Dropbox 下載
  docker_op <動作> [容器名]     Docker 操作（list/start/stop/logs/images）
  pdf_to_images <路徑> [輸出資料夾] [dpi]  PDF 轉圖片
  barcode_scan [圖片路徑]       掃描條碼或 QR Code
  nlp_summarize <文字>          AI 文字摘要
  nlp_sentiment <文字>          AI 情緒分析
  vpn_control <list|connect|disconnect> [名稱] [帳號] [密碼]  VPN 控制
  sys_restore <create|list> [說明]  Windows 系統還原點
  disk_analyze [路徑] [數量]    磁碟空間分析
  face_detect [圖片路徑] [輸出路徑]  人臉偵測
  video_to_gif <路徑> [起始秒] [持續秒] [輸出] [fps]  影片轉 GIF
  excel_chart <路徑> <工作表> [類型] [標題]  Excel 生成圖表
  speedtest                     網路速度測試
  screenshot_compare [圖1] [圖2] [輸出]  截圖比對差異
  set_reminder <HH:MM或秒數> <訊息>  設定一次性提醒
  webpage_screenshot <網址> [輸出]  網頁全頁截圖
  web_monitor <網址> [selector] [間隔秒] [持續秒]  網頁變化監控
  batch_rename <資料夾> <正規表達式> <替換> [副檔名]  批次重新命名
  img_compress <路徑> [品質0-100] [輸出]  壓縮圖片
  batch_img_process <資料夾> <resize|compress> [寬] [高] [品質]  批次圖片處理
  ocr_translate [圖片路徑] [目標語言]  OCR辨識+翻譯
  ip_info [IP地址]              查詢 IP 地理位置（留空查自己）
  currency <金額> <來源幣> <目標幣>  匯率換算
  event_log [日誌名] [等級] [數量]  Windows 事件日誌
  tts_edge <文字> [語音] [輸出]  微軟 Edge TTS 語音合成
  tts_voices                    列出可用 Edge TTS 語音
  send_email_attach <收件人> <主旨> <內容> [附件路徑,...]  含附件郵件
  clipboard_img_get [輸出路徑]  讀取剪貼簿圖片
  clipboard_img_set <圖片路徑>  複製圖片到剪貼簿
  usb_list                      列出 USB 裝置
  firewall <list|add|remove> [名稱] [方向] [port] [協定]  防火牆管理
  todo <add|list|done|delete|clear> [任務] [id]  任務清單
  file_sync <來源> <目標> [dry_run]  資料夾同步
  sysres_chart [秒數] [輸出]    系統資源使用圖表
  password_save <網站> <帳號> <密碼> <主密碼>  加密儲存密碼
  password_get <網站> <主密碼>  查詢已儲存密碼
  rdp_connect <host> [user] [寬] [高]  RDP 遠端桌面
  chrome_bookmarks              列出 Chrome 書籤
  printer_list                  列出印表機
  printer_jobs                  列出列印佇列
  net_share <list|connect|disconnect> [共享路徑] [磁碟代號] [帳號] [密碼]  網路芳鄰
  font_list [關鍵字]            列出系統字型
  volume_get                    查詢系統音量
  volume_set <0-100>            設定系統音量
  volume_mute [true|false]      靜音/取消靜音
  brightness_get                查詢螢幕亮度
  brightness_set <0-100>        設定螢幕亮度
  resolution_list               查詢螢幕解析度
  media_control <動作>          媒體控制（play_pause/next/prev/stop/volume_up/volume_down/mute）
  audio_devices                 列出音訊裝置
  audio_switch <裝置名稱>       切換音訊輸出裝置
  software_list [關鍵字]        列出已安裝軟體
  software_install <名稱>       安裝軟體（winget）
  software_uninstall <名稱>     卸載軟體
  startup_list                  列出開機自啟動程式
  startup_add <名稱> <指令>     新增開機自啟動
  startup_remove <名稱>         移除開機自啟動
  env_get [變數名]              查詢環境變數
  env_set <名稱> <值> [true]    設定環境變數（true=永久）
  user_list                     列出使用者帳戶
  user_create <帳號> <密碼>     建立使用者
  user_delete <帳號>            刪除使用者
  win_update <list|install|check>  Windows 更新管理
  device_list [關鍵字]          裝置管理員列表
  device_toggle <名稱> [true|false]  啟用/停用裝置
  netadapter_list               列出網路介面卡
  netadapter_toggle <名稱> [true|false]  啟用/停用網路卡
  dns_get                       查詢 DNS 設定
  dns_set <介面> <DNS1> [DNS2]  設定 DNS
  ip_config <介面> <IP> [遮罩] [閘道]  設定靜態 IP
  hosts_list                    列出 hosts 設定
  hosts_add <IP> <domain>       新增 hosts
  hosts_remove <domain>         移除 hosts
  net_traffic [秒數]            監控網路流量
  if_then <條件類型> <條件值> <指令> [秒數]  條件式自動化
  window_arrange <side_by_side|quad|stack|maximize_all>  多視窗排列
  region_ocr <x> <y> <w> <h> [語言]  指定區域 OCR
  window_screenshot <視窗標題關鍵字> [輸出路徑]  指定視窗截圖

  ── 新增工具（同步 bot.py）──
  analyze_pdf <路徑> [最大字數]    分析 PDF 文件
  audio_process <action> <輸入> [輸出] [起始ms] [結束ms]  音訊處理
  auto_skill <action> [目標] [名稱] [code]  自動技能管理
  auto_trade <action> [代號] [數量] [價格] [類型]  自動交易
  automation <action> [參數...]    自動化操作
  barcode [圖片路徑]               掃描條碼
  bluetooth <action> [MAC]         藍牙操作
  browser_advanced <action> [參數]  進階瀏覽器控制
  browser_control <action> [url] [selector] [text]  瀏覽器控制
  calendar <action> [天數] [標題] [開始] [結束]  Google 日曆
  clipboard <action> [文字]        剪貼簿操作
  clipboard_image <action> [路徑]  剪貼簿圖片
  cloud_storage <action> <路徑> [drive_id]  雲端儲存
  compare_stocks <代號,...> [指標]  比較股票
  database <type> <db> <SQL>       資料庫操作
  ddg_search <關鍵字> [地區] [數量]  DuckDuckGo 搜尋
  desktop_control <action> [x] [y] [text] [app]  桌面控制
  device_manager <action> [名稱]   裝置管理員
  disk_backup <action> [來源] [目標]  磁碟備份
  display <action> [亮度]          顯示設定
  docker <action> [容器名]         Docker 操作
  document_control <action> <路徑> [內容]  文件控制
  download_file <URL> [路徑]       下載檔案
  dropbox <action> <本地> <遠端> [token]  Dropbox
  email_control <host> <user> <pass>  Email 控制
  emotion_detect <action> [文字] [圖片]  情緒偵測
  encrypt_file <action> <路徑> <密碼>  加解密檔案
  env_var <action> [名稱] [值]     環境變數
  file_system <action> [路徑] [目標]  檔案系統操作
  file_tools <action> <路徑>       檔案工具
  file_transfer <action> <來源> [目標]  檔案傳輸
  find_image_on_screen <圖片> [信心值]  螢幕找圖
  generate_image <prompt> [寬] [高] [疊加文字]  生成圖片（含尺寸）
  get_candlestick_chart <代號> [期間]  K線圖分析
  get_crypto <幣種> [vs貨幣]       查詢加密貨幣
  get_earnings <代號>              查詢財報
  get_etf <代號>                   查詢 ETF
  get_finance_news [來源] [數量]   財經新聞
  get_forex <貨幣對>               查詢匯率
  get_fundamentals <代號>          查詢基本面
  get_macro <指標>                 總經指標
  get_market_sentiment             市場情緒
  get_stock <代號> [期間]          查詢股票（進階版）
  get_stock_advanced <代號> [指標]  進階技術分析
  get_weather <城市>               查詢天氣（進階版）
  git <action> [repo] [message]    Git 操作
  goal_manager <action> [目標]     目標管理
  google_trends <關鍵字,...> [時間] [地區]  Google Trends
  hardware                         硬體監控
  image_edit <action> <路徑> [參數]  圖片編輯
  image_tools <action> [路徑] [quality]  圖片工具
  knowledge_base <action> [content] [query]  知識庫
  long_term_memory <action> [chat_id] [內容]  長期記憶
  lookup <ip|currency> [參數]      查詢工具
  manage_schedule <action> [名稱] [時間]  排程管理
  media <action> [裝置名]          媒體控制
  monitor_config                   螢幕設定資訊
  multi_deploy <action> [host] [user]  多機部署
  multi_perspective <主題>         多角度分析
  network_config <action> [名稱]   網路設定
  network_diag <action> <host>     網路診斷
  news_monitor <action> [關鍵字]   新聞監控
  news_search <關鍵字> [語言] [數量]  新聞搜尋
  nlp <action> <文字>              NLP 工具
  osint_search <action> [query]    OSINT 搜尋
  password_mgr <action> <站> <主密碼>  密碼管理
  pdf_edit <action> [路徑] [輸出]  PDF 編輯
  pdf_image <路徑> [輸出] [dpi]    PDF 轉圖片
  pentest <action> [target]        滲透測試
  portfolio <action> [chat_id] [symbol]  投資組合
  power_control <action>           電源控制
  pptx_control <action> <路徑>     PowerPoint 控制
  proactive_alert <action> [名稱]  主動警報
  ptt_search <關鍵字> [看板] [數量]  PTT 搜尋
  push_notify <platform> <訊息> <webhook>  推播通知
  qr_code <action> [content] [path]  QR Code 操作
  read_screen [問題] [螢幕號]      讀取螢幕內容
  read_webpage <URL> [最大字數]    讀取網頁內容
  registry <action> <key> [value_name]  登錄檔操作
  reminder <時間> <訊息>           設定提醒
  report <標題> <資料JSON> [輸出]  報告生成
  restore_point <action> [說明]    系統還原點
  run_code <type> <code>           執行程式碼
  screen_vision [問題]             螢幕視覺分析
  scroll_at [方向] [格數] [x] [y] [螢幕]  指定位置滾動
  self_benchmark <action>          自我評測
  send_email <收件人> <主旨> <內容>  發送 Email（進階版）
  send_voice <文字> [語音]         生成語音檔
  smart_home <action> [device]     智慧家居
  software <action> [名稱]         軟體管理
  ssh_sftp <action> <host> <user> <pass>  SSH/SFTP
  startup <action> [名稱] [指令]   開機自啟
  system_monitor <action> [target]  系統監控
  system_tools <action>            系統工具集
  think_as <人物> <問題>           角色思考
  threat_intel <action> [target]   威脅情報
  todo_list <action> [task] [id]   任務清單
  tts_advanced <action> [文字] [語音]  進階 TTS
  user_account <action> [帳號]     使用者帳戶
  video_gif <路徑> [起始秒] [持續秒]  影片轉 GIF
  video_process <action> <路徑>    影片處理
  virtual_desktop <action>         虛擬桌面
  voice_cmd <action> [持續秒]      語音命令
  voice_id <action> [名稱] [音檔]  聲紋辨識
  volume <action> [音量]           音量控制
  vpn <action> [名稱]              VPN 控制
  wait_seconds <秒數>              等待指定秒數
  web_scrape <action> [url] [selector]  網頁爬取
  webpage_shot <action> <url>      網頁截圖
  wikipedia_search <關鍵字> [語言]  Wikipedia 搜尋
  win_notify_relay <action>        Windows 通知轉發
  window_control <action> [關鍵字]  視窗控制
  window_manager <action> [視窗名]  視窗管理員
  windows_update <action>          Windows 更新管理
  workflow <action> [名稱] [steps]  工作流程
  youtube_summary <URL>            YouTube 字幕摘要
"""

import sys
import os
import io
import time
import sqlite3
import subprocess
import requests
import pyautogui
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

# 強制 stdout 使用 UTF-8
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

load_dotenv(Path("C:/Users/blue_/claude-telegram-bot/.env"))

pyautogui.FAILSAFE = True
SCREENSHOT_DIR = Path.home() / "Desktop"

# ── 天氣 ────────────────────────────────────────────

def get_weather(city: str):
    try:
        res = requests.get(
            f"https://wttr.in/{city}?format=j1",
            headers={"User-Agent": "curl/7.68.0"},
            timeout=10
        )
        data = res.json()
        current = data["current_condition"][0]
        area = data["nearest_area"][0]
        location = area["areaName"][0]["value"] + ", " + area["country"][0]["value"]
        desc = current["lang_zh-tw"][0]["value"] if "lang_zh-tw" in current else current["weatherDesc"][0]["value"]
        temp = current["temp_C"]
        feels = current["FeelsLikeC"]
        humidity = current["humidity"]
        wind = current["windspeedKmph"]
        wind_dir = current["winddir16Point"]
        print(
            f"📍 {location}\n"
            f"🌤 {desc}\n"
            f"🌡 氣溫：{temp}°C（體感 {feels}°C）\n"
            f"💧 濕度：{humidity}%\n"
            f"💨 風速：{wind} km/h（{wind_dir}）"
        )
    except Exception as e:
        print(f"查不到「{city}」的天氣：{e}")


# ── 股票 ────────────────────────────────────────────

def calc_rsi(closes, period=14):
    deltas = closes.diff().dropna()
    gains = deltas.clip(lower=0)
    losses = -deltas.clip(upper=0)
    avg_gain = gains.rolling(period).mean().iloc[-1]
    avg_loss = losses.rolling(period).mean().iloc[-1]
    if avg_loss == 0:
        return 100.0
    rs = avg_gain / avg_loss
    return round(100 - (100 / (1 + rs)), 1)

def get_stock(symbol: str, period: str = "1mo"):
    try:
        import yfinance as yf
        ticker = yf.Ticker(symbol)
        info = ticker.info
        hist = ticker.history(period="3mo")

        if hist.empty:
            print(f"找不到「{symbol}」的股票數據，請確認代號是否正確。")
            return

        name = info.get("longName") or info.get("shortName") or symbol
        currency = info.get("currency", "")
        current = hist["Close"].iloc[-1]
        prev = hist["Close"].iloc[-2] if len(hist) > 1 else current
        change = current - prev
        change_pct = (change / prev) * 100 if prev else 0
        arrow = "▲" if change >= 0 else "▼"
        volume = hist["Volume"].iloc[-1]
        avg_volume = hist["Volume"].tail(20).mean()
        volume_ratio = volume / avg_volume if avg_volume else 1

        ma5 = hist["Close"].tail(5).mean()
        ma20 = hist["Close"].tail(20).mean()
        ma60 = hist["Close"].tail(60).mean() if len(hist) >= 60 else hist["Close"].mean()
        rsi = calc_rsi(hist["Close"]) if len(hist) >= 15 else None

        if ma5 > ma20 > ma60:
            trend = "強勢多頭（MA5>MA20>MA60）📈"
        elif ma5 < ma20 < ma60:
            trend = "強勢空頭（MA5<MA20<MA60）📉"
        elif ma5 > ma20:
            trend = "短線偏多（MA5>MA20）🔼"
        else:
            trend = "短線偏空（MA5<MA20）🔽"

        rsi_note = ""
        if rsi is not None:
            if rsi >= 80:
                rsi_note = "（嚴重超買 ⚠️）"
            elif rsi >= 70:
                rsi_note = "（超買區間）"
            elif rsi <= 20:
                rsi_note = "（嚴重超賣 💡）"
            elif rsi <= 30:
                rsi_note = "（超賣區間）"
            else:
                rsi_note = "（中性）"

        market_cap = info.get("marketCap")
        pe_ratio = info.get("trailingPE")
        week52_high = info.get("fiftyTwoWeekHigh")
        week52_low = info.get("fiftyTwoWeekLow")
        high_period = hist["High"].tail(20).max()
        low_period = hist["Low"].tail(20).min()
        pct_from_high = ((current - high_period) / high_period * 100) if high_period else 0
        pct_from_low = ((current - low_period) / low_period * 100) if low_period else 0

        result = (
            f"📊 {name} ({symbol})\n"
            f"💰 現價：{current:.2f} {currency}  {arrow} {abs(change):.2f} ({change_pct:+.2f}%)\n"
            f"📦 成交量：{volume:,}（均量 {volume_ratio:.1f}x）\n"
            f"\n── 技術指標 ──\n"
            f"MA5：{ma5:.2f}　MA20：{ma20:.2f}　MA60：{ma60:.2f}\n"
            f"趨勢：{trend}\n"
        )
        if rsi is not None:
            result += f"RSI(14)：{rsi}{rsi_note}\n"
        result += (
            f"近20日高點：{high_period:.2f}（距高 {pct_from_high:.1f}%）\n"
            f"近20日低點：{low_period:.2f}（距低 +{pct_from_low:.1f}%）\n"
        )
        if week52_high and week52_low:
            result += f"52週高低：{week52_low:.2f} ~ {week52_high:.2f}\n"
        result += "\n── 基本面 ──\n"
        if market_cap:
            mc_str = f"{market_cap/1e12:.2f}兆" if market_cap >= 1e12 else f"{market_cap/1e8:.0f}億"
            result += f"市值：{mc_str} {currency}\n"
        if pe_ratio:
            result += f"本益比：{pe_ratio:.1f}\n"

        print(result.strip())

    except Exception as e:
        print(f"查詢「{symbol}」失敗：{e}")


# ── 圖片生成 ────────────────────────────────────────

def add_text_overlay(image_bytes: bytes, text: str) -> bytes:
    try:
        from PIL import Image, ImageDraw, ImageFont
        img = Image.open(io.BytesIO(image_bytes)).convert("RGBA")
        w, h = img.size
        overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(overlay)
        font_size = max(36, w // 12)
        font_path = "C:/Windows/Fonts/msjhbd.ttc"
        try:
            font = ImageFont.truetype(font_path, font_size)
        except Exception:
            font = ImageFont.load_default()
        bbox = draw.textbbox((0, 0), text, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        x = (w - tw) // 2
        y = h - th - int(h * 0.08)
        for dx, dy in [(-2,-2),(2,-2),(-2,2),(2,2)]:
            draw.text((x+dx, y+dy), text, font=font, fill=(0, 0, 0, 180))
        draw.text((x, y), text, font=font, fill=(255, 255, 255, 255))
        result = Image.alpha_composite(img, overlay).convert("RGB")
        buf = io.BytesIO()
        result.save(buf, format="JPEG", quality=92)
        return buf.getvalue()
    except Exception:
        return image_bytes

def generate_image(prompt: str, overlay_text: str = ""):
    hf_token = os.getenv("HF_TOKEN")
    if not hf_token:
        print("未設定 HF_TOKEN，請在 .env 加入 HF_TOKEN=your_token")
        return
    headers = {"Authorization": f"Bearer {hf_token}"}
    payload = {"inputs": prompt}
    models = [
        "black-forest-labs/FLUX.1-schnell",
        "stabilityai/stable-diffusion-xl-base-1.0",
    ]
    image_bytes = None
    for model in models:
        for attempt in range(2):
            try:
                print(f"嘗試模型：{model}...")
                res = requests.post(
                    f"https://router.huggingface.co/hf-inference/models/{model}",
                    headers=headers,
                    json=payload,
                    timeout=120
                )
                if res.status_code == 200 and res.headers.get("content-type", "").startswith("image"):
                    image_bytes = res.content
                    break
                if res.status_code == 503:
                    time.sleep(10)
                    continue
            except Exception:
                pass
        if image_bytes:
            break

    if not image_bytes:
        print("圖片生成失敗")
        return

    if overlay_text:
        image_bytes = add_text_overlay(image_bytes, overlay_text)

    filename = SCREENSHOT_DIR / f"generated_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
    with open(filename, "wb") as f:
        f.write(image_bytes)
    print(f"圖片已儲存：{filename}")


# ── 桌面控制 ────────────────────────────────────────

def screenshot():
    img = pyautogui.screenshot()
    filename = SCREENSHOT_DIR / f"screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    img.save(filename)
    print(f"截圖已儲存：{filename}")
    return str(filename)

def click(x, y):
    pyautogui.click(int(x), int(y))
    print(f"已點擊 ({x}, {y})")

def double_click(x, y):
    pyautogui.doubleClick(int(x), int(y))
    print(f"已雙擊 ({x}, {y})")

def right_click(x, y):
    pyautogui.rightClick(int(x), int(y))
    print(f"已右鍵點擊 ({x}, {y})")

def move(x, y):
    pyautogui.moveTo(int(x), int(y), duration=0.3)
    print(f"滑鼠已移到 ({x}, {y})")

def type_text(text):
    pyautogui.write(text, interval=0.05)
    print(f"已輸入：{text}")

def press_key(key):
    pyautogui.press(key)
    print(f"已按下：{key}")

def open_app(app):
    subprocess.Popen(app, shell=True)
    print(f"已開啟：{app}")

def scroll(direction, amount=3):
    amount = int(amount)
    pyautogui.scroll(amount if direction == "up" else -amount)
    print(f"已向{direction}滾動 {amount} 格")

def pos():
    x, y = pyautogui.position()
    print(f"目前滑鼠位置：({x}, {y})")


# ── 長期記憶 ────────────────────────────────────────

MEMORY_DB = Path("C:/Users/blue_/claude-telegram-bot/memory.db")

def memory_save(chat_id: int, content: str):
    conn = sqlite3.connect(MEMORY_DB)
    conn.execute("CREATE TABLE IF NOT EXISTS long_term_memory (id INTEGER PRIMARY KEY AUTOINCREMENT, chat_id INTEGER NOT NULL, content TEXT NOT NULL, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)")
    conn.execute("INSERT INTO long_term_memory (chat_id, content) VALUES (?, ?)", (chat_id, content))
    conn.commit()
    conn.close()
    print(f"✅ 已儲存記憶：{content}")

def memory_list(chat_id: int):
    conn = sqlite3.connect(MEMORY_DB)
    conn.execute("CREATE TABLE IF NOT EXISTS long_term_memory (id INTEGER PRIMARY KEY AUTOINCREMENT, chat_id INTEGER NOT NULL, content TEXT NOT NULL, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)")
    rows = conn.execute("SELECT id, content, created_at FROM long_term_memory WHERE chat_id=? ORDER BY id DESC", (chat_id,)).fetchall()
    conn.close()
    if not rows:
        print("目前沒有長期記憶。")
        return
    print(f"📝 長期記憶（chat_id={chat_id}）：")
    for r in rows:
        print(f"  [{r[0]}] {r[1]}  ({r[2]})")

def memory_del(chat_id: int, memory_id: int):
    conn = sqlite3.connect(MEMORY_DB)
    conn.execute("DELETE FROM long_term_memory WHERE id=? AND chat_id=?", (memory_id, chat_id))
    conn.commit()
    conn.close()
    print(f"✅ 已刪除記憶 ID {memory_id}")


# ── Vision 截圖理解 ──────────────────────────────────

def vision(question: str = "請描述這個畫面上有什麼，以及目前電腦在做什麼事。"):
    from anthropic import Anthropic
    import base64
    client = Anthropic()
    img = pyautogui.screenshot()
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    img_b64 = base64.standard_b64encode(buf.getvalue()).decode("utf-8")
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": img_b64}},
                {"type": "text", "text": question}
            ]
        }]
    )
    print(response.content[0].text)


# ── 圖像定位 ─────────────────────────────────────────

def find_image(template_path: str, confidence: float = 0.8):
    try:
        location = pyautogui.locateOnScreen(template_path, confidence=confidence)
        if location:
            cx, cy = pyautogui.center(location)
            print(f"✅ 找到圖片：中心座標 ({cx}, {cy})，區域 {location}")
            return cx, cy
        else:
            print("❌ 畫面上找不到該圖片")
            return None
    except Exception as e:
        print(f"❌ 搜尋失敗：{e}")
        return None


# ── 瀏覽器自動化 ──────────────────────────────────────

_browser_context = {}

def browser(action: str, *args):
    from playwright.sync_api import sync_playwright

    if action == "open":
        url = args[0] if args else "https://www.google.com"
        def _open():
            pw = sync_playwright().start()
            b = pw.chromium.launch(headless=False)
            page = b.new_page()
            page.goto(url)
            _browser_context["pw"] = pw
            _browser_context["browser"] = b
            _browser_context["page"] = page
            print(f"✅ 已開啟：{url}")
        _open()

    elif action == "click":
        page = _browser_context.get("page")
        if not page:
            print("❌ 瀏覽器未開啟，請先執行 browser open <url>")
            return
        selector = args[0]
        page.click(selector)
        print(f"✅ 已點擊：{selector}")

    elif action == "type":
        page = _browser_context.get("page")
        if not page:
            print("❌ 瀏覽器未開啟")
            return
        selector, text = args[0], " ".join(args[1:])
        page.fill(selector, text)
        print(f"✅ 已輸入到 {selector}：{text}")

    elif action == "get_text":
        page = _browser_context.get("page")
        if not page:
            print("❌ 瀏覽器未開啟")
            return
        selector = args[0] if args else "body"
        text = page.inner_text(selector)
        print(text[:2000])

    elif action == "screenshot":
        page = _browser_context.get("page")
        if not page:
            print("❌ 瀏覽器未開啟")
            return
        filename = SCREENSHOT_DIR / f"browser_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        page.screenshot(path=str(filename))
        print(f"✅ 截圖已儲存：{filename}")

    elif action == "goto":
        page = _browser_context.get("page")
        if not page:
            print("❌ 瀏覽器未開啟")
            return
        page.goto(args[0])
        print(f"✅ 已前往：{args[0]}")

    elif action == "close":
        if "browser" in _browser_context:
            _browser_context["browser"].close()
            _browser_context["pw"].stop()
            _browser_context.clear()
            print("✅ 瀏覽器已關閉")
    else:
        print(f"❌ 未知動作：{action}。可用：open / click / type / get_text / screenshot / goto / close")


# ── 視窗管理 ─────────────────────────────────────────

def window_list():
    import pygetwindow as gw
    wins = [w for w in gw.getAllWindows() if w.title.strip()]
    for w in wins:
        print(f"  [{w._hWnd}] {w.title}")

def _find_window(keyword):
    import pygetwindow as gw
    wins = [w for w in gw.getAllWindows() if keyword.lower() in w.title.lower()]
    return wins[0] if wins else None

def window_focus(keyword):
    w = _find_window(keyword)
    if w:
        w.activate()
        print(f"✅ 已切換到：{w.title}")
    else:
        print(f"❌ 找不到視窗：{keyword}")

def window_close(keyword):
    w = _find_window(keyword)
    if w:
        w.close()
        print(f"✅ 已關閉：{w.title}")
    else:
        print(f"❌ 找不到視窗：{keyword}")

def window_min(keyword):
    w = _find_window(keyword)
    if w:
        w.minimize()
        print(f"✅ 已最小化：{w.title}")
    else:
        print(f"❌ 找不到視窗：{keyword}")

def window_max(keyword):
    w = _find_window(keyword)
    if w:
        w.maximize()
        print(f"✅ 已最大化：{w.title}")
    else:
        print(f"❌ 找不到視窗：{keyword}")


# ── 組合鍵 / 剪貼簿 ──────────────────────────────────

def hotkey(*keys):
    combo = "+".join(keys)
    pyautogui.hotkey(*keys)
    print(f"✅ 已執行組合鍵：{combo}")

def clipboard_get():
    import pyperclip
    text = pyperclip.paste()
    print(text)

def clipboard_set(text):
    import pyperclip
    pyperclip.copy(text)
    print(f"✅ 已寫入剪貼簿：{text}")


# ── 檔案系統 ─────────────────────────────────────────

def file_list(path="."):
    p = Path(path)
    if not p.exists():
        print(f"❌ 路徑不存在：{path}")
        return
    for item in sorted(p.iterdir()):
        tag = "📁" if item.is_dir() else "📄"
        print(f"  {tag} {item.name}")

def file_read(path):
    p = Path(path)
    if not p.exists():
        print(f"❌ 檔案不存在：{path}")
        return
    print(p.read_text(encoding="utf-8", errors="replace"))

def file_write(path, content):
    Path(path).write_text(content, encoding="utf-8")
    print(f"✅ 已寫入：{path}")

def file_delete(path):
    import shutil
    p = Path(path)
    if not p.exists():
        print(f"❌ 不存在：{path}")
        return
    if p.is_dir():
        shutil.rmtree(p)
    else:
        p.unlink()
    print(f"✅ 已刪除：{path}")

def file_copy(src, dst):
    import shutil
    shutil.copy2(src, dst)
    print(f"✅ 已複製：{src} → {dst}")

def file_move(src, dst):
    import shutil
    shutil.move(src, dst)
    print(f"✅ 已移動：{src} → {dst}")

def file_search(folder, keyword):
    results = list(Path(folder).rglob(f"*{keyword}*"))
    if not results:
        print(f"找不到包含「{keyword}」的檔案")
        return
    for r in results:
        print(f"  {r}")


# ── 系統監控 ─────────────────────────────────────────

def sysinfo():
    import psutil
    cpu = psutil.cpu_percent(interval=1)
    mem = psutil.virtual_memory()
    disk = psutil.disk_usage("C:/")
    print(
        f"💻 系統狀態\n"
        f"CPU：{cpu}%\n"
        f"記憶體：{mem.percent}%（已用 {mem.used//1024//1024}MB / 共 {mem.total//1024//1024}MB）\n"
        f"磁碟 C：{disk.percent}%（已用 {disk.used//1024//1024//1024}GB / 共 {disk.total//1024//1024//1024}GB）"
    )

def process_list():
    import psutil
    print(f"{'PID':<8} {'CPU%':<7} {'記憶體MB':<10} 程序名稱")
    print("-" * 45)
    procs = sorted(psutil.process_iter(["pid","name","cpu_percent","memory_info"]),
                   key=lambda p: p.info["memory_info"].rss if p.info["memory_info"] else 0, reverse=True)
    for p in procs[:30]:
        mem_mb = p.info["memory_info"].rss // 1024 // 1024 if p.info["memory_info"] else 0
        print(f"{p.info['pid']:<8} {p.info['cpu_percent']:<7} {mem_mb:<10} {p.info['name']}")

def process_kill(name_or_pid):
    import psutil
    try:
        pid = int(name_or_pid)
        psutil.Process(pid).kill()
        print(f"✅ 已結束 PID {pid}")
    except ValueError:
        killed = 0
        for p in psutil.process_iter(["pid","name"]):
            if name_or_pid.lower() in p.info["name"].lower():
                p.kill()
                killed += 1
        print(f"✅ 已結束 {killed} 個「{name_or_pid}」程序")


# ── Windows 通知 ──────────────────────────────────────

def notify(title, message):
    try:
        from win10toast import ToastNotifier
        ToastNotifier().show_toast(title, message, duration=5, threaded=True)
        print(f"✅ 通知已送出：{title} - {message}")
    except Exception as e:
        print(f"❌ 通知失敗：{e}")


# ── 語音合成 TTS ──────────────────────────────────────

def clean_for_tts(text: str) -> str:
    """清理文字讓 TTS 更口語：去除 emoji、Markdown、特殊符號"""
    import re
    # 移除 emoji（精確範圍，避免誤刪中文字）
    emoji_pattern = re.compile(
        "["
        "\U0001F600-\U0001F64F"
        "\U0001F300-\U0001F5FF"
        "\U0001F680-\U0001F6FF"
        "\U0001F700-\U0001F9FF"
        "\U0001FA00-\U0001FAFF"
        "\u2600-\u26FF"
        "\u2700-\u27BF"
        "\u2300-\u23FF"
        "\u25A0-\u25FF"
        "\u2B00-\u2BFF"
        "\uFE00-\uFE0F"
        "\u200B-\u200D\uFEFF"
        "]+", flags=re.UNICODE
    )
    text = emoji_pattern.sub("", text)
    # 移除 Markdown
    text = re.sub(r"\*{1,3}(.*?)\*{1,3}", r"\1", text)
    text = re.sub(r"_{1,2}(.*?)_{1,2}", r"\1", text)
    text = re.sub(r"`{1,3}.*?`{1,3}", "", text, flags=re.DOTALL)
    text = re.sub(r"~~(.*?)~~", r"\1", text)
    text = re.sub(r"#{1,6}\s*", "", text)
    text = re.sub(r">\s*", "", text)
    text = re.sub(r"\[([^\]]+)\]\([^\)]+\)", r"\1", text)
    # 移除特殊符號
    text = re.sub(r"[|\\/<>{}=+\[\]~^@#$%&*_]", "", text)
    text = re.sub(r"^\s*[-•·]\s+", "", text, flags=re.MULTILINE)
    # 整理標點
    text = re.sub(r"\.{3,}", "，", text)
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    text = "，".join(lines)
    text = re.sub(r"\s{2,}", " ", text).strip()
    return text

def tts(text):
    import pyttsx3
    engine = pyttsx3.init()
    engine.setProperty("rate", 180)
    engine.say(clean_for_tts(text))
    engine.runAndWait()
    print(f"✅ 已朗讀：{text}")


# ── 動作錄製/重播 ─────────────────────────────────────

_record_events = []
_record_listener = None
RECORD_DIR = Path("C:/Users/blue_/recordings")
RECORD_DIR.mkdir(exist_ok=True)

def record_start():
    from pynput import mouse, keyboard
    global _record_events, _record_listener
    _record_events = []
    start_time = time.time()

    def on_move(x, y):
        _record_events.append({"t": time.time()-start_time, "type": "move", "x": x, "y": y})
    def on_click(x, y, button, pressed):
        _record_events.append({"t": time.time()-start_time, "type": "click", "x": x, "y": y, "button": str(button), "pressed": pressed})
    def on_key(key):
        try:
            _record_events.append({"t": time.time()-start_time, "type": "key", "key": key.char})
        except AttributeError:
            _record_events.append({"t": time.time()-start_time, "type": "key", "key": str(key)})

    ml = mouse.Listener(on_move=on_move, on_click=on_click)
    kl = keyboard.Listener(on_press=on_key)
    ml.start(); kl.start()
    _record_listener = (ml, kl)
    print("✅ 開始錄製，執行 record_stop 停止")

def record_stop():
    global _record_listener
    if _record_listener:
        for l in _record_listener:
            l.stop()
        _record_listener = None
    filename = RECORD_DIR / f"rec_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    import json
    filename.write_text(json.dumps(_record_events, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"✅ 錄製完成，已儲存：{filename}（共 {len(_record_events)} 個事件）")

def record_play(filepath):
    import json
    events = json.loads(Path(filepath).read_text(encoding="utf-8"))
    prev_t = 0
    for e in events:
        delay = e["t"] - prev_t
        if delay > 0:
            time.sleep(min(delay, 2))
        prev_t = e["t"]
        if e["type"] == "move":
            pyautogui.moveTo(e["x"], e["y"])
        elif e["type"] == "click" and e["pressed"]:
            pyautogui.click(e["x"], e["y"])
        elif e["type"] == "key":
            try:
                pyautogui.press(e["key"])
            except Exception:
                pass
    print(f"✅ 重播完成")


# ── Email ─────────────────────────────────────────────

def send_email(to: str, subject: str, body: str):
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    smtp_host = os.getenv("SMTP_HOST", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    smtp_user = os.getenv("SMTP_USER", "")
    smtp_pass = os.getenv("SMTP_PASS", "")
    if not smtp_user or not smtp_pass:
        print("❌ 請在 .env 設定 SMTP_USER 和 SMTP_PASS")
        return
    msg = MIMEMultipart()
    msg["From"] = smtp_user
    msg["To"] = to
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))
    with smtplib.SMTP(smtp_host, smtp_port) as s:
        s.starttls()
        s.login(smtp_user, smtp_pass)
        s.send_message(msg)
    print(f"✅ Email 已寄出到：{to}")


# ── 排程管理 ────────────────────────────────────────

SCHTASKS = "C:\\Windows\\System32\\schtasks.exe"
BOT_SCRIPT = r"C:\Users\blue_\claude-telegram-bot\bot.py"

def bot_status():
    import psutil
    bot_pids = []
    for p in psutil.process_iter(['pid', 'cmdline']):
        try:
            cmd = ' '.join(p.info.get('cmdline') or [])
            if 'bot.py' in cmd and 'psutil' not in cmd:
                bot_pids.append(p.pid)
        except Exception:
            pass
    if bot_pids:
        print(f"✅ Bot 執行中（PID: {', '.join(map(str, bot_pids))}）")
    else:
        print("❌ Bot 未執行")

def bot_restart():
    import psutil, time
    # 停掉所有 bot.py 進程
    for p in psutil.process_iter(['pid', 'cmdline']):
        try:
            cmd = ' '.join(p.info.get('cmdline') or [])
            if 'bot.py' in cmd and 'psutil' not in cmd:
                p.kill()
        except Exception:
            pass
    time.sleep(1)
    pythonw = r"C:\Users\blue_\AppData\Local\Microsoft\WindowsApps\pythonw3.12.exe"
    subprocess.Popen([pythonw, BOT_SCRIPT], cwd=str(Path(BOT_SCRIPT).parent))
    print("✅ Bot 已重啟")

def schedule_list():
    result = subprocess.run(
        [SCHTASKS, "/Query", "/FO", "CSV", "/NH"],
        capture_output=True, text=True, encoding="cp950", errors="replace"
    )
    lines = [l for l in result.stdout.strip().splitlines() if l.strip()]
    print(f"{'任務名稱':<35} {'下次執行':<25} {'狀態'}")
    print("-" * 75)
    for line in lines:
        parts = line.strip('"').split('","')
        if len(parts) >= 3:
            name = parts[0].replace("\\", "").strip()
            next_run = parts[1].strip()
            status = parts[2].strip()
            print(f"{name:<35} {next_run:<25} {status}")

def schedule_add(name: str, time_hhmm: str, script_path: str):
    # 建立排程
    subprocess.run([SCHTASKS, "/Create", "/TN", name,
                    "/TR", f"pythonw {script_path}",
                    "/SC", "DAILY", "/ST", time_hhmm, "/F"],
                   capture_output=True)
    # 設定喚醒
    ps = (
        f"$t = Get-ScheduledTask -TaskName '{name}';"
        f"$t.Settings.WakeToRun = $true;"
        f"$t.Settings.DisallowStartIfOnBatteries = $false;"
        f"$t.Settings.StopIfGoingOnBatteries = $false;"
        f"Set-ScheduledTask -TaskName '{name}' -Settings $t.Settings | Out-Null;"
        f"Write-Host '已建立'"
    )
    subprocess.run(["powershell.exe", "-Command", ps], capture_output=True, text=True)
    print(f"✅ 排程 [{name}] 已建立，每天 {time_hhmm} 執行，可喚醒電腦")

def schedule_del(name: str):
    result = subprocess.run([SCHTASKS, "/Delete", "/TN", name, "/F"],
                            capture_output=True, text=True, encoding="cp950", errors="replace")
    if result.returncode == 0:
        print(f"✅ 排程 [{name}] 已刪除")
    else:
        print(f"❌ 刪除失敗：{result.stderr.strip()}")


# ── 語音辨識 STT ─────────────────────────────────────

def stt(duration=5, file_path="", language="zh-TW"):
    """語音辨識：麥克風錄音或直接辨識音訊檔案（支援 WAV/OGG/MP3）"""
    import speech_recognition as sr, tempfile
    r = sr.Recognizer()
    try:
        if file_path:
            # 若為非 WAV 格式先用 ffmpeg 轉換
            p = Path(file_path)
            if p.suffix.lower() != ".wav":
                import imageio_ffmpeg, subprocess
                ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
                wav = tempfile.mktemp(suffix=".wav")
                subprocess.run([ffmpeg, "-y", "-i", str(p), "-ar", "16000", "-ac", "1", wav], capture_output=True)
                src_path = wav
                cleanup = True
            else:
                src_path = file_path
                cleanup = False
            with sr.AudioFile(src_path) as source:
                audio = r.record(source)
            if cleanup:
                Path(src_path).unlink(missing_ok=True)
        else:
            import sounddevice as sd, soundfile as sf
            sample_rate = 16000
            print(f"🎤 錄音 {duration} 秒，請說話...")
            recording = sd.rec(int(int(duration) * sample_rate), samplerate=sample_rate, channels=1, dtype="int16")
            sd.wait()
            with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as f:
                tmp_path = f.name
            sf.write(tmp_path, recording, sample_rate)
            with sr.AudioFile(tmp_path) as source:
                audio = r.record(source)
            Path(tmp_path).unlink(missing_ok=True)

        text = r.recognize_google(audio, language=language)
        print(f"✅ 辨識結果：{text}")
        return text
    except sr.UnknownValueError:
        print("❌ 無法辨識語音")
    except sr.RequestError as e:
        print(f"❌ STT 服務錯誤：{e}")
    except Exception as e:
        print(f"❌ STT 失敗：{e}")


# ── OCR 文字辨識 ──────────────────────────────────────

def ocr(image_path: str = ""):
    import easyocr
    reader = easyocr.Reader(["ch_tra", "en"], gpu=False)
    if image_path:
        source = image_path
    else:
        img = pyautogui.screenshot()
        source = str(SCREENSHOT_DIR / f"ocr_tmp_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
        img.save(source)
    results = reader.readtext(source)
    texts = [r[1] for r in results]
    output = "\n".join(texts)
    print(output)
    return output


# ── 自動化工作流程 ────────────────────────────────────

WORKFLOW_DIR = Path("C:/Users/blue_/workflows")
WORKFLOW_DIR.mkdir(exist_ok=True)

def workflow_save(name: str, json_content: str):
    import json
    path = WORKFLOW_DIR / f"{name}.json"
    data = json.loads(json_content)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"✅ 工作流程已儲存：{path}")

def workflow_run(json_path: str):
    import json
    path = Path(json_path)
    if not path.exists():
        path = WORKFLOW_DIR / f"{json_path}.json"
    steps = json.loads(path.read_text(encoding="utf-8"))
    print(f"▶ 執行工作流程：{path.name}（共 {len(steps)} 步）")
    for i, step in enumerate(steps, 1):
        tool = step.get("tool")
        args = step.get("args", [])
        delay = step.get("delay", 0)
        print(f"  [{i}] {tool} {args}")
        if delay:
            time.sleep(delay)
        try:
            # 動態呼叫已有工具
            tool_map = {
                "click": lambda a: pyautogui.click(int(a[0]), int(a[1])),
                "type": lambda a: pyautogui.write(" ".join(a), interval=0.05),
                "press": lambda a: pyautogui.press(a[0]),
                "hotkey": lambda a: pyautogui.hotkey(*a),
                "open": lambda a: subprocess.Popen(" ".join(a), shell=True),
                "screenshot": lambda a: screenshot(),
                "move": lambda a: pyautogui.moveTo(int(a[0]), int(a[1]), duration=0.3),
                "scroll": lambda a: pyautogui.scroll(int(a[1]) if a[0]=="up" else -int(a[1])),
                "wait": lambda a: time.sleep(float(a[0])),
                "notify": lambda a: notify(a[0], " ".join(a[1:])),
                "browser": lambda a: browser(a[0], *a[1:]),
                "file_write": lambda a: file_write(a[0], " ".join(a[1:])),
            }
            if tool in tool_map:
                tool_map[tool](args)
            else:
                print(f"    ⚠️ 未知工具：{tool}")
        except Exception as e:
            print(f"    ❌ 步驟失敗：{e}")
    print("✅ 工作流程執行完畢")


# ── 螢幕監控觸發 ──────────────────────────────────────

def screen_watch(template_path: str, command: str, interval: float = 2.0, timeout: float = 60.0):
    print(f"👁 監控中：等待 [{template_path}] 出現，超時 {timeout}s...")
    start = time.time()
    while time.time() - start < timeout:
        try:
            loc = pyautogui.locateOnScreen(template_path, confidence=0.8)
            if loc:
                print(f"✅ 偵測到目標！執行：{command}")
                subprocess.run(command, shell=True)
                return
        except Exception:
            pass
        time.sleep(interval)
    print("⏰ 監控逾時，未偵測到目標")


# ── 多螢幕支援 ───────────────────────────────────────

def monitors():
    try:
        from screeninfo import get_monitors
        for i, m in enumerate(get_monitors()):
            print(f"螢幕 {i}: {m.width}x{m.height} 位置({m.x},{m.y}) {'主螢幕' if m.is_primary else ''}")
    except Exception as e:
        print(f"❌ {e}")


# ── ZIP 壓縮 ─────────────────────────────────────────

def zip_files(source: str, dest: str):
    import zipfile
    src = Path(source)
    with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zf:
        if src.is_dir():
            for f in src.rglob("*"):
                zf.write(f, f.relative_to(src.parent))
        else:
            zf.write(src, src.name)
    print(f"✅ 已壓縮：{source} → {dest}")

def unzip(zip_path: str, dest: str):
    import zipfile
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(dest)
    print(f"✅ 已解壓縮：{zip_path} → {dest}")


# ── 網路下載 ─────────────────────────────────────────

def download(url: str, save_path: str = ""):
    if not save_path:
        save_path = str(SCREENSHOT_DIR / url.split("/")[-1].split("?")[0])
    r = requests.get(url, stream=True, timeout=60)
    r.raise_for_status()
    with open(save_path, "wb") as f:
        for chunk in r.iter_content(chunk_size=8192):
            f.write(chunk)
    print(f"✅ 已下載：{save_path}")


# ── 印表機 ───────────────────────────────────────────

def print_file(path: str):
    try:
        import win32api
        win32api.ShellExecute(0, "print", path, None, ".", 0)
        print(f"✅ 已送印：{path}")
    except Exception as e:
        print(f"❌ 列印失敗：{e}")


# ── WiFi 管理 ────────────────────────────────────────

def wifi_list():
    result = subprocess.run(
        ["powershell.exe", "-Command", "netsh wlan show networks mode=bssid"],
        capture_output=True, text=True, encoding="cp950", errors="replace"
    )
    print(result.stdout[:3000])

def wifi_connect(ssid: str, password: str):
    profile = f"""<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
    <name>{ssid}</name>
    <SSIDConfig><SSID><name>{ssid}</name></SSID></SSIDConfig>
    <connectionType>ESS</connectionType>
    <connectionMode>auto</connectionMode>
    <MSM><security><authEncryption>
        <authentication>WPA2PSK</authentication>
        <encryption>AES</encryption>
    </authEncryption>
    <sharedKey><keyType>passPhrase</keyType><protected>false</protected>
        <keyMaterial>{password}</keyMaterial>
    </sharedKey></security></MSM>
</WLANProfile>"""
    profile_path = Path("C:/Users/blue_/wifi_tmp.xml")
    profile_path.write_text(profile, encoding="utf-8")
    subprocess.run(["powershell.exe", "-Command",
                    f"netsh wlan add profile filename='{profile_path}'; netsh wlan connect name='{ssid}'"],
                   capture_output=True)
    profile_path.unlink(missing_ok=True)
    print(f"✅ 已嘗試連線：{ssid}")


# ── 螢幕串流 ─────────────────────────────────────────

def screen_stream(duration: int = 10, interval: float = 1.0):
    print(f"📹 開始串流 {duration} 秒（每 {interval} 秒截圖）...")
    end = time.time() + duration
    count = 0
    while time.time() < end:
        img = pyautogui.screenshot()
        filename = SCREENSHOT_DIR / f"stream_{datetime.now().strftime('%H%M%S')}_{count:03d}.png"
        img.save(filename)
        count += 1
        time.sleep(interval)
    print(f"✅ 串流完成，共 {count} 張截圖存至桌面")


# ── 語音喚醒 ─────────────────────────────────────────

def wake_listen(keyword: str = "小牛馬", duration: int = 5):
    import sounddevice as sd
    import soundfile as sf
    import speech_recognition as sr
    import tempfile
    print(f"👂 監聽中，等待說出「{keyword}」...")
    sample_rate = 16000
    while True:
        recording = sd.rec(int(duration * sample_rate), samplerate=sample_rate, channels=1, dtype="int16")
        sd.wait()
        with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as f:
            tmp_path = f.name
        sf.write(tmp_path, recording, sample_rate)
        r = sr.Recognizer()
        try:
            with sr.AudioFile(tmp_path) as source:
                audio = r.record(source)
            text = r.recognize_google(audio, language="zh-TW")
            Path(tmp_path).unlink(missing_ok=True)
            if keyword in text:
                print(f"✅ 偵測到喚醒詞：{text}")
                return text
        except Exception:
            Path(tmp_path).unlink(missing_ok=True)


# ── 拖曳 / 右鍵選單 ──────────────────────────────────

def drag(x1, y1, x2, y2, duration=0.5):
    pyautogui.moveTo(int(x1), int(y1))
    pyautogui.dragTo(int(x2), int(y2), duration=float(duration), button="left")
    print(f"✅ 已拖曳 ({x1},{y1}) → ({x2},{y2})")

def right_menu(x, y, item: str = ""):
    pyautogui.rightClick(int(x), int(y))
    time.sleep(0.3)
    if item:
        import pygetwindow as gw
        pyautogui.write(item, interval=0.05)
        time.sleep(0.2)
        pyautogui.press("enter")
        print(f"✅ 已右鍵點擊並選擇：{item}")
    else:
        print(f"✅ 已右鍵點擊 ({x},{y})")


# ── 多步驟 AI 規劃 ────────────────────────────────────

def ai_plan(goal: str):
    from anthropic import Anthropic
    client = Anthropic()
    print(f"🧠 AI 規劃目標：{goal}")
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        system="""你是一個電腦自動化規劃師。用戶給你一個目標，你要把它拆解成可執行的步驟。
每個步驟使用以下工具之一：click/type/press/hotkey/open/screenshot/browser/file_write/wait/notify
以 JSON 陣列格式回傳，例如：
[
  {"tool": "open", "args": ["notepad"], "delay": 1},
  {"tool": "type", "args": ["Hello World"]},
  {"tool": "hotkey", "args": ["ctrl","s"]}
]
只回傳 JSON，不要其他文字。""",
        messages=[{"role": "user", "content": f"目標：{goal}"}]
    )
    import json
    plan_text = response.content[0].text.strip()
    try:
        steps = json.loads(plan_text)
        print(f"📋 規劃完成，共 {len(steps)} 步：")
        for i, s in enumerate(steps, 1):
            print(f"  [{i}] {s.get('tool')} {s.get('args', [])}")
        print("\n▶ 開始執行...")
        tool_map = {
            "click": lambda a: pyautogui.click(int(a[0]), int(a[1])),
            "type": lambda a: pyautogui.write(" ".join(str(x) for x in a), interval=0.05),
            "press": lambda a: pyautogui.press(a[0]),
            "hotkey": lambda a: pyautogui.hotkey(*a),
            "open": lambda a: subprocess.Popen(" ".join(str(x) for x in a), shell=True),
            "screenshot": lambda a: screenshot(),
            "wait": lambda a: time.sleep(float(a[0])),
            "notify": lambda a: notify(a[0], a[1] if len(a) > 1 else ""),
            "move": lambda a: pyautogui.moveTo(int(a[0]), int(a[1]), duration=0.3),
            "scroll": lambda a: pyautogui.scroll(int(a[1]) if a[0] == "up" else -int(a[1])),
            "file_write": lambda a: file_write(a[0], " ".join(str(x) for x in a[1:])),
            "browser": lambda a: browser(a[0], *a[1:]),
        }
        for i, step in enumerate(steps, 1):
            t = step.get("tool")
            a = step.get("args", [])
            d = step.get("delay", 0)
            if d: time.sleep(d)
            if t in tool_map:
                tool_map[t](a)
                print(f"  ✅ 步驟 {i} 完成")
            else:
                print(f"  ⚠️ 未知工具：{t}")
        print("✅ 全部執行完畢")
    except json.JSONDecodeError:
        print(f"規劃結果：\n{plan_text}")


# ── 剪貼簿歷史 ───────────────────────────────────────

_clipboard_history = []

def clipboard_history():
    import pyperclip
    current = pyperclip.paste()
    if current and (not _clipboard_history or _clipboard_history[-1] != current):
        _clipboard_history.append(current)
    if not _clipboard_history:
        print("剪貼簿歷史是空的")
        return
    for i, item in enumerate(_clipboard_history[-20:], 1):
        preview = item[:80].replace("\n", "↵")
        print(f"  [{i}] {preview}")


# ── 虛擬桌面 ─────────────────────────────────────────

def vdesktop(action: str):
    if action == "left":
        pyautogui.hotkey("ctrl", "win", "left")
        print("✅ 切換到左邊虛擬桌面")
    elif action == "right":
        pyautogui.hotkey("ctrl", "win", "right")
        print("✅ 切換到右邊虛擬桌面")
    elif action == "new":
        pyautogui.hotkey("ctrl", "win", "d")
        print("✅ 已建立新虛擬桌面")
    else:
        print(f"❌ 未知動作：{action}（可用：left/right/new）")


# ── 電源管理 ─────────────────────────────────────────

def power(action: str):
    cmds = {
        "sleep":    "rundll32.exe powrprof.dll,SetSuspendState 0,1,0",
        "restart":  "shutdown /r /t 5",
        "shutdown": "shutdown /s /t 5",
    }
    if action not in cmds:
        print(f"❌ 未知動作：{action}（可用：sleep/restart/shutdown）")
        return
    subprocess.run(["powershell.exe", "-Command", cmds[action]])
    print(f"✅ 已執行：{action}")


# ── 藍牙管理 ─────────────────────────────────────────

def bt_scan():
    try:
        import asyncio
        import bleak
        async def _scan():
            devices = await bleak.BleakScanner.discover(timeout=5.0)
            return devices
        devices = asyncio.run(_scan())
        if not devices:
            print("找不到藍牙裝置")
            return
        for d in devices:
            print(f"  {d.address}  {d.name or '(未知)'}")
    except Exception as e:
        print(f"❌ 藍牙掃描失敗：{e}")

def bt_connect(mac: str):
    try:
        import asyncio
        import bleak
        async def _connect():
            async with bleak.BleakClient(mac) as client:
                print(f"✅ 已連線：{mac}（服務數：{len(client.services)}）")
                await asyncio.sleep(3)
        asyncio.run(_connect())
    except Exception as e:
        print(f"❌ 連線失敗：{e}")


# ── 程式碼執行 ───────────────────────────────────────

def run_python(code: str):
    import traceback, contextlib
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            exec(code, {"__builtins__": __builtins__})
        result = buf.getvalue()
        print(result if result else "✅ 執行完成（無輸出）")
    except Exception:
        print(f"❌ 執行錯誤：\n{traceback.format_exc()}")

def run_shell(cmd: str):
    result = subprocess.run(
        ["powershell.exe", "-Command", cmd],
        capture_output=True, text=True, encoding="cp950", errors="replace"
    )
    output = result.stdout + result.stderr
    print(output.strip() or "✅ 執行完成")


# ── 文件處理 ─────────────────────────────────────────

def word_read(path: str):
    from docx import Document
    doc = Document(path)
    text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    print(text[:3000])

def word_write(path: str, content: str):
    from docx import Document
    doc = Document()
    for line in content.split("\n"):
        doc.add_paragraph(line)
    doc.save(path)
    print(f"✅ 已寫入：{path}")

def excel_read(path: str, sheet: str = ""):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.active
    for row in ws.iter_rows(values_only=True):
        print("\t".join(str(c) if c is not None else "" for c in row))

def excel_write(path: str, sheet: str, data_json: str):
    import openpyxl, json
    data = json.loads(data_json)
    try:
        wb = openpyxl.load_workbook(path)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
    ws = wb[sheet] if sheet in wb.sheetnames else wb.create_sheet(sheet)
    for row in data:
        ws.append(row)
    wb.save(path)
    print(f"✅ 已寫入 Excel：{path} [{sheet}]")

def pdf_read(path: str):
    import fitz
    doc = fitz.open(path)
    text = "\n".join(page.get_text() for page in doc)
    print(text[:3000])


# ── 螢幕變化偵測 ──────────────────────────────────────

def screen_diff(interval: float = 1.0, duration: float = 30.0, region=None):
    import numpy as np
    import cv2
    print(f"👁 監控螢幕變化（{duration}秒）...")
    prev = None
    end = time.time() + duration
    changes = 0
    while time.time() < end:
        img = pyautogui.screenshot(region=region)
        arr = np.array(img.convert("L"))
        if prev is not None:
            diff = cv2.absdiff(prev, arr)
            score = diff.mean()
            if score > 5:
                changes += 1
                ts = datetime.now().strftime("%H:%M:%S")
                print(f"  [{ts}] 偵測到變化！差異分數：{score:.1f}")
        prev = arr
        time.sleep(interval)
    print(f"✅ 監控結束，共偵測到 {changes} 次變化")


# ── 網頁爬蟲 ─────────────────────────────────────────

def scrape(url: str, selector: str = "body"):
    from bs4 import BeautifulSoup
    res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
    res.raise_for_status()
    soup = BeautifulSoup(res.text, "html.parser")
    elements = soup.select(selector)
    if not elements:
        print(f"找不到選擇器：{selector}")
        return
    for el in elements[:10]:
        text = el.get_text(strip=True)
        if text:
            print(text[:200])


# ── 圖片編輯 ─────────────────────────────────────────

def img_edit(action: str, path: str, *args):
    from PIL import Image, ImageDraw, ImageFont
    img = Image.open(path)

    if action == "crop":
        x1, y1, x2, y2 = int(args[0]), int(args[1]), int(args[2]), int(args[3])
        img = img.crop((x1, y1, x2, y2))
    elif action == "resize":
        w, h = int(args[0]), int(args[1])
        img = img.resize((w, h), Image.LANCZOS)
    elif action == "text":
        text = args[0]
        x, y = int(args[1]) if len(args) > 1 else img.width // 2, int(args[2]) if len(args) > 2 else img.height - 50
        draw = ImageDraw.Draw(img)
        try:
            font = ImageFont.truetype("C:/Windows/Fonts/msjhbd.ttc", 36)
        except Exception:
            font = ImageFont.load_default()
        draw.text((x, y), text, font=font, fill=(255, 255, 255))
    elif action == "merge":
        img2 = Image.open(args[0])
        new = Image.new("RGB", (img.width + img2.width, max(img.height, img2.height)))
        new.paste(img, (0, 0))
        new.paste(img2, (img.width, 0))
        img = new
    elif action == "grayscale":
        img = img.convert("L")

    out_path = args[-1] if args and str(args[-1]).endswith((".png",".jpg",".jpeg")) else path
    img.save(out_path)
    print(f"✅ 圖片已儲存：{out_path}")


# ── Google Drive ──────────────────────────────────────

def gdrive_upload(file_path: str, folder_id: str = ""):
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
        from google.oauth2.credentials import Credentials
        creds_path = Path("C:/Users/blue_/gdrive_token.json")
        if not creds_path.exists():
            print("❌ 找不到 C:/Users/blue_/gdrive_token.json，請先完成 Google 授權")
            return
        creds = Credentials.from_authorized_user_file(str(creds_path))
        service = build("drive", "v3", credentials=creds)
        meta = {"name": Path(file_path).name}
        if folder_id:
            meta["parents"] = [folder_id]
        media = MediaFileUpload(file_path)
        f = service.files().create(body=meta, media_body=media, fields="id").execute()
        print(f"✅ 已上傳，檔案 ID：{f.get('id')}")
    except Exception as e:
        print(f"❌ 上傳失敗：{e}")

def gdrive_download(file_id: str, save_path: str):
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload
        from google.oauth2.credentials import Credentials
        creds = Credentials.from_authorized_user_file("C:/Users/blue_/gdrive_token.json")
        service = build("drive", "v3", credentials=creds)
        req = service.files().get_media(fileId=file_id)
        with open(save_path, "wb") as f:
            downloader = MediaIoBaseDownload(f, req)
            done = False
            while not done:
                _, done = downloader.next_chunk()
        print(f"✅ 已下載：{save_path}")
    except Exception as e:
        print(f"❌ 下載失敗：{e}")


# ── 資料庫操作 ───────────────────────────────────────

def db_query(db_path: str, sql: str):
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.execute(sql)
        if cur.description:
            headers = [d[0] for d in cur.description]
            print("\t".join(headers))
            print("-" * 60)
            for row in cur.fetchall():
                print("\t".join(str(c) for c in row))
        else:
            conn.commit()
            print(f"✅ 執行成功，影響 {cur.rowcount} 列")
    finally:
        conn.close()

def db_mysql(host: str, database: str, sql: str, user: str = "root", password: str = ""):
    import pymysql
    conn = pymysql.connect(host=host, user=user, password=password, database=database, charset="utf8mb4")
    try:
        with conn.cursor() as cur:
            cur.execute(sql)
            if cur.description:
                headers = [d[0] for d in cur.description]
                print("\t".join(headers))
                for row in cur.fetchall():
                    print("\t".join(str(c) for c in row))
            else:
                conn.commit()
                print(f"✅ 執行成功，影響 {cur.rowcount} 列")
    finally:
        conn.close()


# ── 加密/解密 ────────────────────────────────────────

def encrypt(file_path: str, password: str):
    from cryptography.fernet import Fernet
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    import base64
    salt = b"xinyang501_salt_"
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=100000)
    key = base64.urlsafe_b64encode(kdf.derive(password.encode()))
    f = Fernet(key)
    data = Path(file_path).read_bytes()
    out_path = file_path + ".enc"
    Path(out_path).write_bytes(f.encrypt(data))
    print(f"✅ 已加密：{out_path}")

def decrypt(file_path: str, password: str):
    from cryptography.fernet import Fernet
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    import base64
    salt = b"xinyang501_salt_"
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=100000)
    key = base64.urlsafe_b64encode(kdf.derive(password.encode()))
    f = Fernet(key)
    data = Path(file_path).read_bytes()
    out_path = file_path.replace(".enc", ".dec")
    Path(out_path).write_bytes(f.decrypt(data))
    print(f"✅ 已解密：{out_path}")


# ── 剪貼簿監控 ───────────────────────────────────────

def clipboard_watch(duration: float = 30.0):
    import pyperclip
    print(f"📋 監控剪貼簿 {duration} 秒...")
    prev = pyperclip.paste()
    end = time.time() + duration
    changes = []
    while time.time() < end:
        current = pyperclip.paste()
        if current != prev and current:
            ts = datetime.now().strftime("%H:%M:%S")
            preview = current[:80].replace("\n", "↵")
            print(f"  [{ts}] 新內容：{preview}")
            changes.append(current)
            prev = current
        time.sleep(0.5)
    print(f"✅ 監控結束，共偵測到 {len(changes)} 次變化")


# ── QR Code ──────────────────────────────────────────

def qr_gen(content: str, save_path: str = ""):
    import qrcode
    qr = qrcode.make(content)
    if not save_path:
        save_path = str(SCREENSHOT_DIR / f"qr_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
    qr.save(save_path)
    print(f"✅ QR Code 已生成：{save_path}")

def qr_scan(image_path: str = ""):
    try:
        from pyzbar.pyzbar import decode
        from PIL import Image
        if not image_path:
            img = pyautogui.screenshot()
        else:
            img = Image.open(image_path)
        results = decode(img)
        if not results:
            print("❌ 未偵測到 QR Code")
            return
        for r in results:
            print(f"✅ 掃描結果：{r.data.decode('utf-8')}")
    except Exception as e:
        print(f"❌ 掃描失敗：{e}")


# ── 螢幕錄影 ────────────────────────────────────────

def screen_record(duration: float = 10.0, output: str = ""):
    try:
        import mss, cv2, numpy as np, time as t
        out_path = output or str(Path.home() / "Desktop" / f"record_{datetime.now().strftime('%H%M%S')}.mp4")
        with mss.mss() as sct:
            mon = sct.monitors[1]
            w, h = mon["width"], mon["height"]
            fps = 10
            writer = cv2.VideoWriter(out_path, cv2.VideoWriter_fourcc(*"mp4v"), fps, (w, h))
            end = t.time() + duration
            while t.time() < end:
                frame = np.array(sct.grab(mon))
                frame = cv2.cvtColor(frame, cv2.COLOR_BGRA2BGR)
                writer.write(frame)
                t.sleep(1 / fps)
            writer.release()
        print(f"✅ 錄影完成：{out_path}")
    except Exception as e:
        print(f"❌ 錄影失敗：{e}")


def webcam_capture(output: str = ""):
    try:
        import cv2
        out_path = output or str(Path.home() / "Desktop" / f"webcam_{datetime.now().strftime('%H%M%S')}.jpg")
        cap = cv2.VideoCapture(0)
        ret, frame = cap.read()
        cap.release()
        if ret:
            cv2.imwrite(out_path, frame)
            print(f"✅ 已拍照：{out_path}")
        else:
            print("❌ 無法存取攝影機")
    except Exception as e:
        print(f"❌ 攝影機失敗：{e}")


# ── 翻譯 ────────────────────────────────────────────

def translate(text: str, target: str = "zh-TW", source: str = "auto"):
    try:
        from deep_translator import GoogleTranslator
        result = GoogleTranslator(source=source, target=target).translate(text)
        print(result)
    except Exception as e:
        print(f"❌ 翻譯失敗：{e}")


# ── 資料視覺化 ───────────────────────────────────────

def chart(chart_type: str, data_json: str, title: str = "", output: str = ""):
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import json
        data = json.loads(data_json)
        out_path = output or str(Path.home() / "Desktop" / f"chart_{datetime.now().strftime('%H%M%S')}.png")
        fig, ax = plt.subplots()
        if chart_type == "line":
            for label, values in data.items():
                ax.plot(values, label=label)
            ax.legend()
        elif chart_type == "bar":
            ax.bar(list(data.keys()), list(data.values()))
        elif chart_type == "pie":
            ax.pie(list(data.values()), labels=list(data.keys()), autopct="%1.1f%%")
        if title:
            ax.set_title(title)
        plt.tight_layout()
        plt.savefig(out_path)
        plt.close()
        print(f"✅ 圖表已存：{out_path}")
    except Exception as e:
        print(f"❌ 圖表生成失敗：{e}")


# ── PowerPoint ───────────────────────────────────────

def pptx_read(path: str):
    try:
        from pptx import Presentation
        prs = Presentation(path)
        for i, slide in enumerate(prs.slides, 1):
            texts = [sh.text for sh in slide.shapes if sh.has_text_frame]
            print(f"[投影片 {i}] " + " | ".join(t for t in texts if t.strip()))
    except Exception as e:
        print(f"❌ 讀取失敗：{e}")

def pptx_create(path: str, slides_json: str):
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt
        import json
        slides = json.loads(slides_json)
        prs = Presentation()
        for s in slides:
            layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(layout)
            if "title" in s:
                slide.shapes.title.text = s["title"]
            if "body" in s and slide.placeholders[1]:
                slide.placeholders[1].text = s["body"]
        prs.save(path)
        print(f"✅ 已建立簡報：{path}")
    except Exception as e:
        print(f"❌ 建立失敗：{e}")


# ── REST API ─────────────────────────────────────────

def api_call(method: str, url: str, headers_json: str = "{}", body_json: str = "{}"):
    try:
        import json
        headers = json.loads(headers_json)
        body = json.loads(body_json)
        resp = requests.request(method.upper(), url, headers=headers, json=body if body else None, timeout=30)
        try:
            print(json.dumps(resp.json(), ensure_ascii=False, indent=2)[:3000])
        except Exception:
            print(resp.text[:3000])
    except Exception as e:
        print(f"❌ API 呼叫失敗：{e}")


# ── 程序守護 ─────────────────────────────────────────

def watchdog(process_name: str, script: str, duration: float = 60.0):
    try:
        import psutil, time as t
        print(f"🐕 開始監控 [{process_name}]，持續 {duration} 秒...")
        end = t.time() + duration
        restarts = 0
        while t.time() < end:
            running = any(p.name().lower() == process_name.lower() for p in psutil.process_iter())
            if not running:
                subprocess.Popen(["pythonw" if script.endswith(".py") else script, script] if script.endswith(".py") else [script])
                restarts += 1
                print(f"⚠️ [{process_name}] 已重啟（第 {restarts} 次）")
            t.sleep(5)
        print(f"✅ 守護結束，共重啟 {restarts} 次")
    except Exception as e:
        print(f"❌ 守護失敗：{e}")


# ── SSH ──────────────────────────────────────────────

def ssh_run(host: str, user: str, password: str, command: str, port: int = 22):
    try:
        import paramiko
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(host, port=port, username=user, password=password, timeout=15)
        _, stdout, stderr = client.exec_command(command)
        out = stdout.read().decode(errors="replace")
        err = stderr.read().decode(errors="replace")
        client.close()
        print((out + err).strip() or "（執行完畢，無輸出）")
    except Exception as e:
        print(f"❌ SSH 失敗：{e}")

def sftp_upload(host: str, user: str, password: str, local: str, remote: str, port: int = 22):
    try:
        import paramiko
        t = paramiko.Transport((host, port))
        t.connect(username=user, password=password)
        sftp = paramiko.SFTPClient.from_transport(t)
        sftp.put(local, remote)
        sftp.close(); t.close()
        print(f"✅ 已上傳：{local} → {remote}")
    except Exception as e:
        print(f"❌ SFTP 上傳失敗：{e}")

def sftp_download(host: str, user: str, password: str, remote: str, local: str, port: int = 22):
    try:
        import paramiko
        t = paramiko.Transport((host, port))
        t.connect(username=user, password=password)
        sftp = paramiko.SFTPClient.from_transport(t)
        sftp.get(remote, local)
        sftp.close(); t.close()
        print(f"✅ 已下載：{remote} → {local}")
    except Exception as e:
        print(f"❌ SFTP 下載失敗：{e}")


# ── 網路診斷 ─────────────────────────────────────────

def net_ping(host: str, count: int = 4):
    result = subprocess.run(["ping", "-n", str(count), host], capture_output=True, text=True, encoding="cp950", errors="replace")
    print(result.stdout.strip())

def net_traceroute(host: str):
    result = subprocess.run(["tracert", host], capture_output=True, text=True, encoding="cp950", errors="replace", timeout=60)
    print(result.stdout[:3000])

def net_portscan(host: str, ports: str = "22,80,443,3306,3389,8080"):
    import socket
    results = []
    for p in [int(x) for x in ports.split(",")]:
        try:
            s = socket.socket()
            s.settimeout(1)
            r = s.connect_ex((host, p))
            results.append(f"Port {p}: {'開放 ✅' if r == 0 else '關閉 ❌'}")
            s.close()
        except Exception:
            results.append(f"Port {p}: 錯誤")
    print("\n".join(results))


# ── Windows 服務 ─────────────────────────────────────

def win_service(action: str, name: str = ""):
    try:
        if action == "list":
            r = subprocess.run(["powershell.exe", "-Command",
                "Get-Service | Select-Object Name,Status | Format-Table -AutoSize"],
                capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout[:3000])
        elif action in ("start", "stop"):
            cmd = f"{'Start' if action=='start' else 'Stop'}-Service -Name '{name}' -Force"
            r = subprocess.run(["powershell.exe", "-Command", cmd],
                capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout or r.stderr or f"✅ {action} {name}")
    except Exception as e:
        print(f"❌ 服務操作失敗：{e}")


# ── PDF 編輯 ─────────────────────────────────────────

def pdf_merge(paths_json: str, output: str):
    try:
        import fitz, json
        paths = json.loads(paths_json)
        writer = fitz.open()
        for p in paths:
            writer.insert_pdf(fitz.open(p))
        writer.save(output)
        print(f"✅ 已合併 {len(paths)} 個 PDF：{output}")
    except Exception as e:
        print(f"❌ 合併失敗：{e}")

def pdf_split(path: str, output_dir: str):
    try:
        import fitz
        doc = fitz.open(path)
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        for i, page in enumerate(doc):
            out = fitz.open()
            out.insert_pdf(doc, from_page=i, to_page=i)
            out.save(str(Path(output_dir) / f"page_{i+1}.pdf"))
        print(f"✅ 已分割 {len(doc)} 頁到：{output_dir}")
    except Exception as e:
        print(f"❌ 分割失敗：{e}")

def pdf_watermark(path: str, text: str, output: str = ""):
    try:
        import fitz
        doc = fitz.open(path)
        out_path = output or path.replace(".pdf", "_wm.pdf")
        for page in doc:
            page.insert_text((page.rect.width/2 - 50, page.rect.height/2),
                text, fontsize=40, color=(0.8, 0.8, 0.8), rotate=45)
        doc.save(out_path)
        print(f"✅ 已加浮水印：{out_path}")
    except Exception as e:
        print(f"❌ 浮水印失敗：{e}")


# ── 音訊處理 ─────────────────────────────────────────

def audio_convert(input_path: str, output_path: str):
    try:
        from pydub import AudioSegment
        fmt = Path(output_path).suffix.lstrip(".")
        AudioSegment.from_file(input_path).export(output_path, format=fmt)
        print(f"✅ 已轉換：{output_path}")
    except Exception as e:
        print(f"❌ 音訊轉換失敗：{e}")

def audio_trim(input_path: str, start_ms: int, end_ms: int, output_path: str = ""):
    try:
        from pydub import AudioSegment
        audio = AudioSegment.from_file(input_path)[start_ms:end_ms]
        out = output_path or input_path.replace(".", "_trim.")
        audio.export(out, format=Path(out).suffix.lstrip("."))
        print(f"✅ 已剪輯：{out}")
    except Exception as e:
        print(f"❌ 音訊剪輯失敗：{e}")


# ── 推播通知 ─────────────────────────────────────────

def discord_notify(webhook_url: str, message: str):
    try:
        resp = requests.post(webhook_url, json={"content": message}, timeout=10)
        print(f"✅ Discord 已發送（{resp.status_code}）")
    except Exception as e:
        print(f"❌ Discord 發送失敗：{e}")

def line_notify(token: str, message: str):
    try:
        resp = requests.post(
            "https://notify-api.line.me/api/notify",
            headers={"Authorization": f"Bearer {token}"},
            data={"message": message}, timeout=10
        )
        print(f"✅ LINE 已發送（{resp.status_code}）")
    except Exception as e:
        print(f"❌ LINE 發送失敗：{e}")


# ── 磁碟清理 ─────────────────────────────────────────

def disk_clean(action: str = "list"):
    try:
        import tempfile, shutil
        tmp = Path(tempfile.gettempdir())
        if action == "list":
            files = list(tmp.rglob("*"))
            total = sum(f.stat().st_size for f in files if f.is_file())
            print(f"暫存資料夾：{tmp}\n檔案數：{len(files)}\n佔用空間：{total/1024/1024:.1f} MB")
        elif action == "clean":
            count = 0
            for f in tmp.iterdir():
                try:
                    if f.is_file():
                        f.unlink(); count += 1
                    elif f.is_dir():
                        shutil.rmtree(f, ignore_errors=True); count += 1
                except Exception:
                    pass
            print(f"✅ 已清理 {count} 個暫存項目")
    except Exception as e:
        print(f"❌ 磁碟清理失敗：{e}")

def backup(src: str, dest: str):
    try:
        import shutil
        out = Path(dest) / f"{Path(src).name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        shutil.make_archive(str(out).replace(".zip",""), "zip", src)
        print(f"✅ 備份完成：{out}")
    except Exception as e:
        print(f"❌ 備份失敗：{e}")


# ── Windows 登錄檔 ────────────────────────────────────

def registry_read(key_path: str, value_name: str = ""):
    try:
        import winreg
        parts = key_path.split("\\", 1)
        roots = {"HKEY_LOCAL_MACHINE": winreg.HKEY_LOCAL_MACHINE,
                 "HKEY_CURRENT_USER": winreg.HKEY_CURRENT_USER,
                 "HKLM": winreg.HKEY_LOCAL_MACHINE, "HKCU": winreg.HKEY_CURRENT_USER}
        root = roots[parts[0]]
        with winreg.OpenKey(root, parts[1]) as k:
            if value_name:
                val, _ = winreg.QueryValueEx(k, value_name)
                print(f"{value_name} = {val}")
            else:
                i = 0
                while True:
                    try:
                        n, v, _ = winreg.EnumValue(k, i)
                        print(f"{n} = {v}")
                        i += 1
                    except OSError:
                        break
    except Exception as e:
        print(f"❌ 讀取登錄檔失敗：{e}")

def registry_write(key_path: str, value_name: str, value: str):
    try:
        import winreg
        parts = key_path.split("\\", 1)
        roots = {"HKEY_LOCAL_MACHINE": winreg.HKEY_LOCAL_MACHINE,
                 "HKEY_CURRENT_USER": winreg.HKEY_CURRENT_USER,
                 "HKLM": winreg.HKEY_LOCAL_MACHINE, "HKCU": winreg.HKEY_CURRENT_USER}
        root = roots[parts[0]]
        with winreg.OpenKey(root, parts[1], 0, winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, value_name, 0, winreg.REG_SZ, value)
        print(f"✅ 已寫入：{value_name} = {value}")
    except Exception as e:
        print(f"❌ 寫入登錄檔失敗：{e}")


# ── 影片處理 ─────────────────────────────────────────

def video_screenshot(path: str, second: float = 0, output: str = ""):
    try:
        import cv2
        cap = cv2.VideoCapture(path)
        cap.set(cv2.CAP_PROP_POS_MSEC, second * 1000)
        ret, frame = cap.read()
        cap.release()
        out = output or path.replace(".mp4", f"_frame{int(second)}s.jpg")
        if ret:
            cv2.imwrite(out, frame)
            print(f"✅ 已擷取畫面：{out}")
        else:
            print("❌ 無法讀取影片")
    except Exception as e:
        print(f"❌ 影片截圖失敗：{e}")

def video_trim(path: str, start_sec: float, end_sec: float, output: str = ""):
    try:
        out = output or path.replace(".mp4", f"_trim.mp4")
        subprocess.run([
            "ffmpeg", "-y", "-i", path,
            "-ss", str(start_sec), "-to", str(end_sec),
            "-c", "copy", out
        ], capture_output=True)
        print(f"✅ 已剪輯：{out}")
    except Exception as e:
        print(f"❌ 影片剪輯失敗：{e}")


# ── 多螢幕管理 ───────────────────────────────────────

def monitor_list():
    try:
        from screeninfo import get_monitors
        for m in get_monitors():
            print(f"{'主螢幕 ' if m.is_primary else '副螢幕 '}{m.width}x{m.height} @({m.x},{m.y}) name={m.name}")
    except Exception as e:
        print(f"❌ 取得螢幕資訊失敗：{e}")


# ── Email 讀取 (IMAP) ────────────────────────────────

def email_read(host: str, user: str, password: str, folder: str = "INBOX", count: int = 5):
    try:
        import imapclient, email as _email
        from email.header import decode_header
        client = imapclient.IMAPClient(host, ssl=True)
        client.login(user, password)
        client.select_folder(folder)
        msgs = client.search(["ALL"])
        recent = msgs[-count:] if len(msgs) >= count else msgs
        results = []
        for uid in reversed(recent):
            raw = client.fetch([uid], ["RFC822"])[uid][b"RFC822"]
            msg = _email.message_from_bytes(raw)
            subj_raw, enc = decode_header(msg["Subject"])[0]
            subject = subj_raw.decode(enc or "utf-8") if isinstance(subj_raw, bytes) else subj_raw
            sender = msg["From"]
            date = msg["Date"]
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode(errors="replace")[:200]
                        break
            else:
                body = msg.get_payload(decode=True).decode(errors="replace")[:200]
            results.append(f"[{date}]\n寄件人：{sender}\n主旨：{subject}\n{body}\n{'─'*30}")
        client.logout()
        print("\n".join(results))
    except Exception as e:
        print(f"❌ 讀取郵件失敗：{e}")


# ── Google Calendar ───────────────────────────────────

def gcal_list(days: int = 7):
    try:
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build
        from datetime import timezone, timedelta
        creds_path = Path("C:/Users/blue_/claude-telegram-bot/gcal_token.json")
        if not creds_path.exists():
            print("❌ 未找到 Google Calendar 憑證（gcal_token.json）")
            return
        creds = Credentials.from_authorized_user_file(str(creds_path))
        service = build("calendar", "v3", credentials=creds)
        now = datetime.now(timezone.utc)
        end = now + timedelta(days=days)
        events = service.events().list(
            calendarId="primary",
            timeMin=now.isoformat(),
            timeMax=end.isoformat(),
            maxResults=20, singleEvents=True, orderBy="startTime"
        ).execute().get("items", [])
        if not events:
            print(f"未來 {days} 天沒有行程")
            return
        for e in events:
            start = e["start"].get("dateTime", e["start"].get("date"))
            print(f"📅 {start}  {e.get('summary','（無標題）')}")
    except Exception as e:
        print(f"❌ 讀取行事曆失敗：{e}")

def gcal_add(title: str, start: str, end: str, description: str = ""):
    try:
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build
        creds_path = Path("C:/Users/blue_/claude-telegram-bot/gcal_token.json")
        if not creds_path.exists():
            print("❌ 未找到 Google Calendar 憑證（gcal_token.json）")
            return
        creds = Credentials.from_authorized_user_file(str(creds_path))
        service = build("calendar", "v3", credentials=creds)
        event = {
            "summary": title,
            "description": description,
            "start": {"dateTime": start, "timeZone": "Asia/Taipei"},
            "end": {"dateTime": end, "timeZone": "Asia/Taipei"},
        }
        created = service.events().insert(calendarId="primary", body=event).execute()
        print(f"✅ 行程已新增：{created.get('htmlLink')}")
    except Exception as e:
        print(f"❌ 新增行程失敗：{e}")


# ── 全域快捷鍵 ───────────────────────────────────────

def global_hotkey_listen(hotkey: str, command: str, duration: float = 60.0):
    try:
        import keyboard as kb, time as t
        triggered = []
        def on_trigger():
            triggered.append(datetime.now().strftime("%H:%M:%S"))
            subprocess.run(command, shell=True)
        kb.add_hotkey(hotkey, on_trigger)
        print(f"🎹 監聽快捷鍵 [{hotkey}]，持續 {duration} 秒...")
        t.sleep(duration)
        kb.remove_all_hotkeys()
        print(f"✅ 共觸發 {len(triggered)} 次：{triggered}")
    except Exception as e:
        print(f"❌ 快捷鍵監聽失敗：{e}")


# ── Git 操作 ─────────────────────────────────────────

def git_op(action: str, repo: str = ".", message: str = "", remote: str = "origin", branch: str = "master"):
    try:
        import git as _git
        repo_obj = _git.Repo(repo)
        if action == "status":
            print(repo_obj.git.status())
        elif action == "log":
            for c in list(repo_obj.iter_commits())[:10]:
                print(f"{c.hexsha[:7]} [{c.authored_datetime.strftime('%Y-%m-%d %H:%M')}] {c.message.strip()[:60]}")
        elif action == "pull":
            result = repo_obj.remotes[remote].pull()
            print(f"✅ Pull 完成：{result[0].commit.hexsha[:7]}")
        elif action == "add":
            repo_obj.git.add(A=True)
            print("✅ 已 git add -A")
        elif action == "commit":
            repo_obj.index.commit(message or "auto commit")
            print(f"✅ 已 commit：{message}")
        elif action == "push":
            repo_obj.remotes[remote].push(branch)
            print(f"✅ 已 push 到 {remote}/{branch}")
        elif action == "diff":
            print(repo_obj.git.diff()[:3000] or "（無變更）")
    except Exception as e:
        print(f"❌ Git 操作失敗：{e}")


# ── 硬體監控進階 ─────────────────────────────────────

def hw_monitor():
    try:
        import psutil
        cpu_temp = ""
        try:
            temps = psutil.sensors_temperatures()
            if temps:
                for name, entries in temps.items():
                    for e in entries[:2]:
                        cpu_temp += f"{name}: {e.current}°C  "
        except Exception:
            cpu_temp = "（不支援溫度感測）"

        battery = psutil.sensors_battery()
        bat_str = f"{battery.percent:.0f}% {'充電中' if battery.power_plugged else '使用電池'}" if battery else "無電池"

        gpu_str = ""
        try:
            import GPUtil
            gpus = GPUtil.getGPUs()
            for g in gpus:
                gpu_str += f"\nGPU [{g.name}] 使用率：{g.load*100:.0f}% | 記憶體：{g.memoryUsed:.0f}/{g.memoryTotal:.0f}MB | 溫度：{g.temperature}°C"
        except Exception:
            gpu_str = "\nGPU：（未偵測到 NVIDIA GPU）"

        print(
            f"🌡 溫度：{cpu_temp}\n"
            f"🔋 電池：{bat_str}"
            f"{gpu_str}"
        )
    except Exception as e:
        print(f"❌ 硬體監控失敗：{e}")


# ── 自動報告生成 ─────────────────────────────────────

def report_gen(title: str, data_json: str, output: str = ""):
    try:
        import jinja2, json
        data = json.loads(data_json)
        out_path = output or str(Path.home() / "Desktop" / f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
        template_str = """<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>body{font-family:sans-serif;margin:40px}table{border-collapse:collapse;width:100%}
th,td{border:1px solid #ccc;padding:8px;text-align:left}th{background:#4472C4;color:white}
tr:nth-child(even){background:#f2f2f2}h1{color:#4472C4}</style></head>
<body><h1>{{ title }}</h1><p>生成時間：{{ time }}</p>
{% for section, rows in data.items() %}
<h2>{{ section }}</h2>
{% if rows is iterable and rows is not string %}
{% if rows[0] is mapping %}
<table><tr>{% for k in rows[0].keys() %}<th>{{ k }}</th>{% endfor %}</tr>
{% for row in rows %}<tr>{% for v in row.values() %}<td>{{ v }}</td>{% endfor %}</tr>{% endfor %}
</table>
{% else %}<ul>{% for item in rows %}<li>{{ item }}</li>{% endfor %}</ul>{% endif %}
{% else %}<p>{{ rows }}</p>{% endif %}
{% endfor %}</body></html>"""
        tmpl = jinja2.Template(template_str)
        html = tmpl.render(title=title, data=data, time=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        Path(out_path).write_text(html, encoding="utf-8")
        print(f"✅ 報告已生成：{out_path}")
    except Exception as e:
        print(f"❌ 報告生成失敗：{e}")


# ── Dropbox ──────────────────────────────────────────

def dropbox_upload(local: str, remote: str, token: str = ""):
    try:
        import dropbox as dbx
        tok = token or os.getenv("DROPBOX_TOKEN", "")
        if not tok:
            print("❌ 請設定 DROPBOX_TOKEN 環境變數")
            return
        d = dbx.Dropbox(tok)
        with open(local, "rb") as f:
            d.files_upload(f.read(), remote, mode=dbx.files.WriteMode.overwrite)
        print(f"✅ 已上傳到 Dropbox：{remote}")
    except Exception as e:
        print(f"❌ Dropbox 上傳失敗：{e}")

def dropbox_download(remote: str, local: str, token: str = ""):
    try:
        import dropbox as dbx
        tok = token or os.getenv("DROPBOX_TOKEN", "")
        if not tok:
            print("❌ 請設定 DROPBOX_TOKEN 環境變數")
            return
        d = dbx.Dropbox(tok)
        _, res = d.files_download(remote)
        Path(local).write_bytes(res.content)
        print(f"✅ 已從 Dropbox 下載：{local}")
    except Exception as e:
        print(f"❌ Dropbox 下載失敗：{e}")


# ── Docker ───────────────────────────────────────────

def docker_op(action: str, name: str = ""):
    try:
        import docker as _docker
        client = _docker.from_env()
        if action == "list":
            for c in client.containers.list(all=True):
                print(f"[{c.status}] {c.name}  {c.image.tags}")
        elif action == "start":
            client.containers.get(name).start()
            print(f"✅ 容器 [{name}] 已啟動")
        elif action == "stop":
            client.containers.get(name).stop()
            print(f"✅ 容器 [{name}] 已停止")
        elif action == "logs":
            logs = client.containers.get(name).logs(tail=50).decode(errors="replace")
            print(logs)
        elif action == "images":
            for img in client.images.list():
                print(f"{img.tags}  {img.short_id}")
    except Exception as e:
        print(f"❌ Docker 操作失敗：{e}")


# ── PDF 轉圖片 ───────────────────────────────────────

def pdf_to_images(path: str, output_dir: str = "", dpi: int = 150):
    try:
        import fitz
        doc = fitz.open(path)
        out = Path(output_dir) if output_dir else Path(path).parent / (Path(path).stem + "_imgs")
        out.mkdir(parents=True, exist_ok=True)
        for i, page in enumerate(doc):
            mat = fitz.Matrix(dpi/72, dpi/72)
            pix = page.get_pixmap(matrix=mat)
            img_path = str(out / f"page_{i+1}.png")
            pix.save(img_path)
        print(f"✅ 已轉換 {len(doc)} 頁到：{out}")
    except Exception as e:
        print(f"❌ PDF 轉圖片失敗：{e}")


# ── 條碼掃描 ─────────────────────────────────────────

def barcode_scan(image_path: str = ""):
    try:
        from pyzbar.pyzbar import decode
        from PIL import Image
        img = Image.open(image_path) if image_path else pyautogui.screenshot()
        results = decode(img)
        if not results:
            print("❌ 未偵測到條碼或 QR Code")
            return
        for r in results:
            print(f"類型：{r.type}  內容：{r.data.decode('utf-8', errors='replace')}")
    except Exception as e:
        print(f"❌ 條碼掃描失敗：{e}")


# ── NLP 文字分析 ─────────────────────────────────────

def nlp_summarize(text: str):
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=512,
            messages=[{"role": "user", "content": f"請用繁體中文摘要以下文字（100字以內）：\n\n{text}"}]
        )
        print(msg.content[0].text)
    except Exception as e:
        print(f"❌ 摘要失敗：{e}")

def nlp_sentiment(text: str):
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=128,
            messages=[{"role": "user", "content": f"分析以下文字的情緒，只回覆：正面/負面/中性 + 一句說明：\n\n{text}"}]
        )
        print(msg.content[0].text)
    except Exception as e:
        print(f"❌ 情緒分析失敗：{e}")


# ── VPN 控制 ─────────────────────────────────────────

def vpn_control(action: str, name: str = "", user: str = "", password: str = ""):
    try:
        if action == "list":
            r = subprocess.run(["powershell.exe", "-Command",
                "Get-VpnConnection | Select-Object Name,ConnectionStatus | Format-Table"],
                capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout or "（未設定 VPN）")
        elif action == "connect":
            r = subprocess.run(["rasdial", name, user, password],
                capture_output=True, text=True, encoding="cp950", errors="replace")
            print(r.stdout.strip())
        elif action == "disconnect":
            r = subprocess.run(["rasdial", name, "/disconnect"],
                capture_output=True, text=True, encoding="cp950", errors="replace")
            print(r.stdout.strip())
    except Exception as e:
        print(f"❌ VPN 操作失敗：{e}")


# ── 系統還原點 ───────────────────────────────────────

def sys_restore(action: str, description: str = ""):
    try:
        if action == "create":
            ps = f"Checkpoint-Computer -Description '{description or 'Claude Auto Restore'}' -RestorePointType MODIFY_SETTINGS"
            r = subprocess.run(["powershell.exe", "-Command", ps],
                capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout or "✅ 還原點已建立")
        elif action == "list":
            r = subprocess.run(["powershell.exe", "-Command",
                "Get-ComputerRestorePoint | Select-Object SequenceNumber,Description,CreationTime | Format-Table"],
                capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout or "（無還原點）")
    except Exception as e:
        print(f"❌ 系統還原點操作失敗：{e}")


# ── 磁碟分析 ─────────────────────────────────────────

def disk_analyze(path: str = "C:/", top: int = 10):
    try:
        import psutil
        usage = psutil.disk_usage(path)
        print(f"磁碟：{path}\n總容量：{usage.total/1024**3:.1f} GB\n已使用：{usage.used/1024**3:.1f} GB ({usage.percent}%)\n可用：{usage.free/1024**3:.1f} GB\n")
        sizes = []
        try:
            for item in Path(path).iterdir():
                try:
                    if item.is_dir():
                        size = sum(f.stat().st_size for f in item.rglob("*") if f.is_file())
                    else:
                        size = item.stat().st_size
                    sizes.append((size, str(item)))
                except Exception:
                    pass
        except Exception:
            pass
        sizes.sort(reverse=True)
        print(f"佔用最多的前 {top} 個項目：")
        for size, name in sizes[:top]:
            print(f"  {size/1024**3:.2f} GB  {name}")
    except Exception as e:
        print(f"❌ 磁碟分析失敗：{e}")


# ── 人臉偵測 ─────────────────────────────────────────

def face_detect(image_path: str = "", output: str = ""):
    try:
        import cv2, numpy as np
        if not image_path:
            img_pil = pyautogui.screenshot()
            img = cv2.cvtColor(np.array(img_pil), cv2.COLOR_RGB2BGR)
        else:
            img = cv2.imread(image_path)
        cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        faces = cascade.detectMultiScale(gray, 1.1, 4)
        for (x, y, w, h) in faces:
            cv2.rectangle(img, (x, y), (x+w, y+h), (0, 255, 0), 2)
        out_path = output or str(Path.home() / "Desktop" / f"faces_{datetime.now().strftime('%H%M%S')}.jpg")
        cv2.imwrite(out_path, img)
        print(f"✅ 偵測到 {len(faces)} 張人臉，已存：{out_path}")
    except Exception as e:
        print(f"❌ 人臉偵測失敗：{e}")


# ── 影片轉 GIF ───────────────────────────────────────

def video_to_gif(path: str, start: float = 0, duration: float = 5.0, output: str = "", fps: int = 10):
    try:
        import imageio
        out = output or path.replace(".mp4", ".gif")
        reader = imageio.get_reader(path)
        meta = reader.get_meta_data()
        video_fps = meta.get("fps", 30)
        start_frame = int(start * video_fps)
        end_frame = int((start + duration) * video_fps)
        frames = []
        for i, frame in enumerate(reader):
            if i < start_frame:
                continue
            if i >= end_frame:
                break
            frames.append(frame)
        imageio.mimsave(out, frames, fps=fps)
        print(f"✅ GIF 已生成：{out}（{len(frames)} 幀）")
    except Exception as e:
        print(f"❌ 影片轉 GIF 失敗：{e}")


# ── 影片生成 ─────────────────────────────────────────

def ai_video(prompt: str, provider: str = "replicate", model: str = "",
             image_url: str = "", duration: float = 5, output: str = ""):
    """
    用 AI API 生成影片
    provider: replicate / runway / kling
    需要 .env 或環境變數中對應的 API Key
    """
    import requests, time, traceback, os
    from pathlib import Path
    from dotenv import load_dotenv
    load_dotenv(Path("C:/Users/blue_/claude-telegram-bot/.env"))

    out = output or str(Path.home() / "Desktop" / f"ai_video_{int(time.time())}.mp4")

    def _download(url, dest):
        r = requests.get(url, timeout=120, stream=True)
        r.raise_for_status()
        with open(dest, "wb") as f:
            for chunk in r.iter_content(8192):
                f.write(chunk)
        return dest

    try:
        if provider == "replicate":
            api_key = os.getenv("REPLICATE_API_TOKEN", "")
            if not api_key:
                print("❌ 缺少 REPLICATE_API_TOKEN，請在 .env 加入"); return
            mdl = model or ("stability-ai/stable-video-diffusion" if image_url else "minimax/video-01")
            headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
            inputs = {"prompt": prompt}
            if image_url: inputs["image"] = image_url
            if duration:  inputs["duration"] = int(duration)
            r = requests.post(
                f"https://api.replicate.com/v1/models/{mdl}/predictions",
                json={"input": inputs}, headers=headers, timeout=30
            )
            r.raise_for_status()
            pred_id = r.json()["id"]
            for _ in range(60):
                time.sleep(5)
                resp = requests.get(f"https://api.replicate.com/v1/predictions/{pred_id}",
                                    headers=headers, timeout=15)
                data = resp.json()
                if data.get("status") == "succeeded":
                    url = data["output"]
                    if isinstance(url, list): url = url[0]
                    _download(url, out)
                    print(f"✅ Replicate 影片已生成：{out}")
                    return
                elif data.get("status") == "failed":
                    print(f"❌ Replicate 失敗：{data.get('error')}")
                    return
            print("❌ Replicate 逾時")

        elif provider == "runway":
            api_key = os.getenv("RUNWAY_API_KEY", "")
            if not api_key:
                print("❌ 缺少 RUNWAY_API_KEY，請在 .env 加入"); return
            import runwayml
            client = runwayml.RunwayML(api_key=api_key)
            if image_url:
                task = client.image_to_video.create(
                    model="gen4_turbo", prompt_image=image_url,
                    prompt_text=prompt, duration=int(min(duration, 10)), ratio="1280:720")
            else:
                task = client.text_to_video.create(
                    model="gen4_turbo", prompt_text=prompt,
                    duration=int(min(duration, 10)), ratio="1280:720")
            for _ in range(60):
                time.sleep(5)
                t = client.tasks.retrieve(task.id)
                if t.status == "SUCCEEDED":
                    _download(t.output[0], out)
                    print(f"✅ Runway 影片已生成：{out}"); return
                elif t.status in ("FAILED", "CANCELLED"):
                    print(f"❌ Runway 失敗：{t.failure_reason}"); return
            print("❌ Runway 逾時")

        elif provider == "kling":
            import jwt as _jwt
            access_key = os.getenv("KLING_ACCESS_KEY", "")
            secret_key = os.getenv("KLING_SECRET_KEY", "")
            if not access_key or not secret_key:
                print("❌ 缺少 KLING_ACCESS_KEY / KLING_SECRET_KEY"); return
            def _token():
                return _jwt.encode({"iss": access_key, "exp": int(time.time())+1800,
                                    "nbf": int(time.time())-5}, secret_key, algorithm="HS256")
            hdrs = {"Authorization": f"Bearer {_token()}", "Content-Type": "application/json"}
            body = {"model_name": "kling-v1", "prompt": prompt,
                    "duration": str(int(min(duration, 10))), "aspect_ratio": "16:9"}
            if image_url: body["image_url"] = image_url
            ep = "image2video" if image_url else "text2video"
            r = requests.post(f"https://api.klingai.com/v1/videos/{ep}",
                              json=body, headers=hdrs, timeout=30)
            r.raise_for_status()
            d = r.json()
            if d.get("code") != 0:
                print(f"❌ Kling 錯誤：{d.get('message')}"); return
            task_id = d["data"]["task_id"]
            for _ in range(60):
                time.sleep(5)
                hdrs["Authorization"] = f"Bearer {_token()}"
                resp = requests.get(f"https://api.klingai.com/v1/videos/{ep}/{task_id}",
                                    headers=hdrs, timeout=15)
                dd = resp.json().get("data", {})
                if dd.get("task_status") == "succeed":
                    _download(dd["task_result"]["videos"][0]["url"], out)
                    print(f"✅ Kling 影片已生成：{out}"); return
                elif dd.get("task_status") == "failed":
                    print(f"❌ Kling 失敗：{dd.get('task_status_msg')}"); return
            print("❌ Kling 逾時")
        else:
            print(f"❌ 未知 provider：{provider}")
    except Exception as e:
        print(f"❌ AI 影片生成失敗：{e}\n{traceback.format_exc()}")


def video_gen(mode: str = "slideshow", output: str = "", **kwargs):
    """
    生成影片
    mode: slideshow / text_video / tts_video / screen_record
    """
    import re as _re
    import numpy as np
    import subprocess
    import tempfile
    import traceback
    from pathlib import Path
    from PIL import Image, ImageDraw, ImageFont
    import imageio_ffmpeg

    ffmpeg_exe = imageio_ffmpeg.get_ffmpeg_exe()
    w, h = kwargs.get("size", (1280, 720))
    fps  = kwargs.get("fps", 24)
    out  = output or str(Path.home() / "Desktop" / f"video_{datetime.now().strftime('%H%M%S')}.mp4")

    def _write_frames(frames_iter, out_path, vid_fps, width, height):
        proc = subprocess.Popen([
            ffmpeg_exe, "-y",
            "-f", "rawvideo", "-vcodec", "rawvideo",
            "-s", f"{width}x{height}", "-pix_fmt", "rgb24",
            "-r", str(vid_fps), "-i", "pipe:0",
            "-vcodec", "libx264", "-pix_fmt", "yuv420p",
            "-preset", "fast", "-crf", "23",
            out_path
        ], stdin=subprocess.PIPE, stderr=subprocess.PIPE)
        for frame in frames_iter:
            proc.stdin.write(np.asarray(frame, dtype=np.uint8).tobytes())
        proc.stdin.close()
        rc  = proc.wait()
        err = proc.stderr.read().decode(errors="replace")
        if rc != 0 or not Path(out_path).exists():
            raise RuntimeError(f"ffmpeg 失敗 rc={rc}: {err[-300:]}")

    def _get_font(pt=60):
        for fp in ["C:/Windows/Fonts/msjh.ttc", "C:/Windows/Fonts/msyh.ttc",
                   "C:/Windows/Fonts/simhei.ttf", "C:/Windows/Fonts/arial.ttf"]:
            if Path(fp).exists():
                try:
                    return ImageFont.truetype(fp, pt)
                except Exception:
                    pass
        return ImageFont.load_default()

    def _tts_sync(text: str, voice: str, out_path: str):
        import asyncio, edge_tts
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            loop.run_until_complete(edge_tts.Communicate(text, voice).save(out_path))
        finally:
            loop.close()
        if not Path(out_path).exists():
            raise RuntimeError("TTS 語音檔案未生成")

    def _get_audio_dur(audio_path: str) -> float:
        r = subprocess.run([ffmpeg_exe, "-i", audio_path],
                           capture_output=True, text=True, errors="replace")
        m = _re.search(r"Duration:\s*(\d+):(\d+):(\d+\.?\d*)", r.stderr)
        return int(m.group(1)) * 3600 + int(m.group(2)) * 60 + float(m.group(3)) if m else 10.0

    try:
        if mode == "slideshow":
            images = kwargs.get("images", [])
            dur    = kwargs.get("duration", 3)
            sl_fps = kwargs.get("fps", 12)
            trans  = kwargs.get("transition", 0.5)
            if not images:
                print("❌ 請提供 images 參數（圖片路徑列表）"); return

            sf = int(sl_fps * dur)
            tf = int(sl_fps * trans)

            def _gen():
                loaded = []
                for p in images:
                    try:
                        loaded.append(np.array(Image.open(p).convert("RGB").resize((w, h), Image.LANCZOS)))
                    except Exception:
                        loaded.append(np.zeros((h, w, 3), dtype=np.uint8))
                for arr in loaded:
                    for i in range(tf):  yield (arr * (i / tf)).astype(np.uint8)
                    for _ in range(max(1, sf - tf * 2)):  yield arr
                    for i in range(tf):  yield (arr * (1.0 - i / tf)).astype(np.uint8)

            _write_frames(_gen(), out, sl_fps, w, h)
            print(f"✅ 投影片影片已生成：{out}")

        elif mode == "text_video":
            text   = kwargs.get("text", "Hello")
            dur    = kwargs.get("duration", 5)
            bg_col = tuple(kwargs.get("bg_color",   [30, 30, 40]))
            fg_col = tuple(kwargs.get("font_color", [255, 255, 255]))
            fsize  = kwargs.get("font_size", 60)
            font   = _get_font(fsize)
            total  = max(1, int(fps * dur))

            def _gen():
                for i in range(total):
                    p = i / total
                    a = p / 0.2 if p < 0.2 else (1.0 - p) / 0.2 if p > 0.8 else 1.0
                    img  = Image.new("RGB", (w, h), bg_col)
                    draw = ImageDraw.Draw(img)
                    c    = tuple(int(x * a) for x in fg_col)
                    lines = text.split("\n")
                    lh    = fsize + 10
                    y0    = (h - len(lines) * lh) // 2
                    for j, ln in enumerate(lines):
                        bb = draw.textbbox((0, 0), ln, font=font)
                        draw.text(((w - (bb[2] - bb[0])) // 2, y0 + j * lh), ln, font=font, fill=c)
                    yield np.array(img)

            _write_frames(_gen(), out, fps, w, h)
            print(f"✅ 文字動畫影片已生成：{out}")

        elif mode == "tts_video":
            text     = kwargs.get("text", "")
            img_path = kwargs.get("image", "")
            voice    = kwargs.get("voice", "zh-CN-YunjianNeural")
            subtitle = kwargs.get("subtitle", True)
            if not text:
                print("❌ 請提供 text 參數"); return

            tmp    = Path(tempfile.mkdtemp())
            audio  = str(tmp / "tts.mp3")
            vidtmp = str(tmp / "silent.mp4")

            _tts_sync(clean_for_tts(text), voice, audio)
            dur   = _get_audio_dur(audio)
            total = max(1, int(fps * dur))

            if img_path and Path(img_path).exists():
                bg = np.array(Image.open(img_path).convert("RGB").resize((w, h), Image.LANCZOS))
            else:
                bg = np.zeros((h, w, 3), dtype=np.uint8)
                for row in range(h):
                    v = int(20 + 30 * row / h)
                    bg[row, :] = [v, v, v + 20]

            font = _get_font(40)
            cpl  = 20
            subs = [text[i:i+cpl] for i in range(0, len(text), cpl)]

            def _gen():
                for i in range(total):
                    frame = bg.copy()
                    if subtitle and subs:
                        ln  = subs[min(int(i / total * len(subs)), len(subs) - 1)]
                        img = Image.fromarray(frame)
                        d   = ImageDraw.Draw(img)
                        bb  = d.textbbox((0, 0), ln, font=font)
                        tw  = bb[2] - bb[0]
                        tx  = (w - tw) // 2
                        ty  = int(h * 0.82)
                        d.rectangle([tx - 10, ty - 5, tx + tw + 10, ty + 45], fill=(0, 0, 0))
                        d.text((tx, ty), ln, font=font, fill=(255, 255, 255))
                        frame = np.array(img)
                    yield frame

            _write_frames(_gen(), vidtmp, fps, w, h)
            r = subprocess.run([
                ffmpeg_exe, "-y", "-i", vidtmp, "-i", audio,
                "-c:v", "copy", "-c:a", "aac", "-shortest", out
            ], capture_output=True)
            if not Path(out).exists():
                raise RuntimeError(f"音訊合併失敗：{r.stderr.decode(errors='replace')[-200:]}")
            print(f"✅ TTS 語音影片已生成：{out}")

        elif mode == "screen_record":
            import mss, time as _time
            dur      = kwargs.get("duration", 10)
            rec_fps  = kwargs.get("fps", 10)
            interval = 1.0 / rec_fps
            total    = max(1, int(rec_fps * dur))

            with mss.mss() as sct:
                mon    = sct.monitors[1]
                sw, sh = mon["width"], mon["height"]

                def _gen():
                    for _ in range(total):
                        t0  = _time.time()
                        arr = np.array(sct.grab(mon))[:, :, :3][:, :, ::-1]
                        yield arr
                        elapsed = _time.time() - t0
                        if elapsed < interval:
                            _time.sleep(interval - elapsed)

                _write_frames(_gen(), out, rec_fps, sw, sh)
            print(f"✅ 螢幕錄影完成：{out}（{dur} 秒）")

        else:
            print(f"❌ 未知 mode：{mode}")

    except Exception as e:
        print(f"❌ 影片生成失敗：{e}\n{traceback.format_exc()}")


# ── Excel 圖表 ───────────────────────────────────────

def excel_chart(path: str, sheet: str, chart_type: str = "bar", title: str = ""):
    try:
        import openpyxl
        from openpyxl.chart import BarChart, LineChart, PieChart, Reference
        wb = openpyxl.load_workbook(path)
        ws = wb[sheet] if sheet in wb.sheetnames else wb.active
        max_row = ws.max_row
        max_col = ws.max_column
        chart_map = {"bar": BarChart, "line": LineChart, "pie": PieChart}
        chart = chart_map.get(chart_type, BarChart)()
        chart.title = title or sheet
        chart.style = 10
        data = Reference(ws, min_col=2, min_row=1, max_row=max_row, max_col=max_col)
        cats = Reference(ws, min_col=1, min_row=2, max_row=max_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        ws.add_chart(chart, "A" + str(max_row + 2))
        wb.save(path)
        print(f"✅ 圖表已加入：{path}")
    except Exception as e:
        print(f"❌ Excel 圖表失敗：{e}")


# ── 網路速度測試 ─────────────────────────────────────

def speedtest_run():
    try:
        import speedtest as st
        print("⏳ 測速中（約需 30 秒）...")
        s = st.Speedtest()
        s.get_best_server()
        download = s.download() / 1_000_000
        upload = s.upload() / 1_000_000
        ping = s.results.ping
        server = s.results.server.get("name","")
        print(f"📶 網路速度測試結果\n下載：{download:.1f} Mbps\n上傳：{upload:.1f} Mbps\nPing：{ping:.0f} ms\n伺服器：{server}")
    except Exception as e:
        print(f"❌ 速度測試失敗：{e}")


# ── 螢幕截圖比對 ─────────────────────────────────────

def screenshot_compare(img1_path: str = "", img2_path: str = "", output: str = ""):
    try:
        import cv2, numpy as np
        from PIL import Image
        if not img1_path:
            img1 = np.array(pyautogui.screenshot())
        else:
            img1 = cv2.imread(img1_path)
        if not img2_path:
            import time as t; t.sleep(2)
            img2 = np.array(pyautogui.screenshot())
        else:
            img2 = cv2.imread(img2_path)
        img1_bgr = cv2.cvtColor(img1, cv2.COLOR_RGB2BGR) if img1_path == "" else img1
        img2_bgr = cv2.cvtColor(img2, cv2.COLOR_RGB2BGR) if img2_path == "" else img2
        h, w = min(img1_bgr.shape[0], img2_bgr.shape[0]), min(img1_bgr.shape[1], img2_bgr.shape[1])
        img1_bgr, img2_bgr = img1_bgr[:h,:w], img2_bgr[:h,:w]
        diff = cv2.absdiff(img1_bgr, img2_bgr)
        gray = cv2.cvtColor(diff, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 30, 255, cv2.THRESH_BINARY)
        contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        result = img2_bgr.copy()
        cv2.drawContours(result, contours, -1, (0, 0, 255), 2)
        out = output or str(Path.home() / "Desktop" / f"diff_{datetime.now().strftime('%H%M%S')}.png")
        cv2.imwrite(out, result)
        changed = cv2.countNonZero(thresh)
        total = h * w
        pct = changed / total * 100
        print(f"✅ 差異：{pct:.2f}%，已標記差異區域：{out}")
    except Exception as e:
        print(f"❌ 截圖比對失敗：{e}")


# ── 一次性提醒 ───────────────────────────────────────

def set_reminder(time_str: str, message: str):
    """time_str: HH:MM 或 秒數（如 '300' 代表5分鐘後）"""
    try:
        import threading, time as t
        def _remind():
            if time_str.isdigit():
                t.sleep(int(time_str))
            else:
                import datetime as dt
                now = dt.datetime.now()
                target = dt.datetime.strptime(time_str, "%H:%M").replace(
                    year=now.year, month=now.month, day=now.day)
                if target < now:
                    target = target.replace(day=now.day + 1)
                t.sleep((target - now).total_seconds())
            try:
                from win10toast import ToastNotifier
                ToastNotifier().show_toast("⏰ 提醒", message, duration=10)
            except Exception:
                pass
            try:
                import pyttsx3
                engine = pyttsx3.init()
                engine.say(message)
                engine.runAndWait()
            except Exception:
                pass
            print(f"⏰ 提醒：{message}")
        threading.Thread(target=_remind, daemon=True).start()
        print(f"✅ 提醒已設定：{time_str} → {message}")
    except Exception as e:
        print(f"❌ 設定提醒失敗：{e}")


# ── 網頁全頁截圖 ─────────────────────────────────────

def webpage_screenshot(url: str, output: str = ""):
    try:
        from playwright.sync_api import sync_playwright
        out = output or str(Path.home() / "Desktop" / f"webpage_{datetime.now().strftime('%H%M%S')}.png")
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page(viewport={"width": 1280, "height": 800})
            page.goto(url, timeout=30000)
            page.wait_for_load_state("networkidle")
            page.screenshot(path=out, full_page=True)
            browser.close()
        print(f"✅ 網頁截圖已存：{out}")
    except Exception as e:
        print(f"❌ 網頁截圖失敗：{e}")


# ── 網頁變化監控 ─────────────────────────────────────

def web_monitor(url: str, selector: str = "body", interval: float = 60.0, duration: float = 3600.0):
    try:
        import hashlib, time as t
        from bs4 import BeautifulSoup
        def _fetch():
            r = requests.get(url, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
            soup = BeautifulSoup(r.text, "html.parser")
            elements = soup.select(selector)
            text = "\n".join(e.get_text(strip=True) for e in elements)
            return hashlib.md5(text.encode()).hexdigest(), text[:200]
        last_hash, _ = _fetch()
        print(f"🔍 開始監控：{url}（每 {interval} 秒，持續 {duration} 秒）")
        end = t.time() + duration
        changes = 0
        while t.time() < end:
            t.sleep(interval)
            try:
                new_hash, snippet = _fetch()
                if new_hash != last_hash:
                    changes += 1
                    print(f"⚠️ [{datetime.now().strftime('%H:%M:%S')}] 網頁有變化！\n{snippet}")
                    last_hash = new_hash
            except Exception as e:
                print(f"檢查失敗：{e}")
        print(f"✅ 監控結束，共偵測到 {changes} 次變化")
    except Exception as e:
        print(f"❌ 網頁監控失敗：{e}")


# ── 批次重新命名 ─────────────────────────────────────

def batch_rename(folder: str, pattern: str, replacement: str, ext_filter: str = ""):
    try:
        import re
        folder_path = Path(folder)
        files = [f for f in folder_path.iterdir() if f.is_file()]
        if ext_filter:
            files = [f for f in files if f.suffix.lower() == ext_filter.lower()]
        count = 0
        for f in sorted(files):
            new_name = re.sub(pattern, replacement, f.stem) + f.suffix
            new_path = f.parent / new_name
            if new_path != f:
                f.rename(new_path)
                print(f"  {f.name} → {new_name}")
                count += 1
        print(f"✅ 已重新命名 {count} 個檔案")
    except Exception as e:
        print(f"❌ 批次重新命名失敗：{e}")


# ── 圖片壓縮 ─────────────────────────────────────────

def img_compress(path: str, quality: int = 75, output: str = ""):
    try:
        from PIL import Image
        img = Image.open(path)
        out = output or path.replace(".", f"_q{quality}.")
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        img.save(out, optimize=True, quality=quality)
        orig_size = Path(path).stat().st_size
        new_size = Path(out).stat().st_size
        print(f"✅ 壓縮完成：{orig_size//1024}KB → {new_size//1024}KB（{(1-new_size/orig_size)*100:.1f}% 節省）\n儲存至：{out}")
    except Exception as e:
        print(f"❌ 圖片壓縮失敗：{e}")

def batch_img_process(folder: str, action: str, width: int = 0, height: int = 0, quality: int = 75):
    try:
        from PIL import Image
        folder_path = Path(folder)
        out_dir = folder_path / f"output_{action}"
        out_dir.mkdir(exist_ok=True)
        count = 0
        for f in folder_path.glob("*.{jpg,jpeg,png,bmp,webp}"):
            try:
                img = Image.open(f)
                if action == "resize" and width and height:
                    img = img.resize((width, height))
                elif action == "compress":
                    if img.mode in ("RGBA", "P"): img = img.convert("RGB")
                img.save(str(out_dir / f.name), optimize=True, quality=quality)
                count += 1
            except Exception:
                pass
        print(f"✅ 批次處理完成：{count} 張圖片 → {out_dir}")
    except Exception as e:
        print(f"❌ 批次圖片處理失敗：{e}")


# ── OCR + 翻譯 pipeline ──────────────────────────────

def ocr_translate(image_path: str = "", target: str = "zh-TW"):
    try:
        import easyocr, numpy as np
        from PIL import Image
        from deep_translator import GoogleTranslator
        reader = easyocr.Reader(["ch_tra", "en"], gpu=False)
        img = Image.open(image_path) if image_path else pyautogui.screenshot()
        results = reader.readtext(np.array(img))
        text = " ".join([r[1] for r in results])
        if not text.strip():
            print("❌ 未辨識到文字")
            return
        print(f"原文：{text[:300]}")
        translated = GoogleTranslator(source="auto", target=target).translate(text)
        print(f"翻譯（{target}）：{translated}")
    except Exception as e:
        print(f"❌ OCR 翻譯失敗：{e}")


# ── IP 資訊查詢 ──────────────────────────────────────

def ip_info(ip: str = ""):
    try:
        target = ip or ""
        res = requests.get(f"http://ip-api.com/json/{target}?lang=zh-TW", timeout=10)
        d = res.json()
        if d.get("status") == "fail":
            print(f"❌ 查詢失敗：{d.get('message')}")
            return
        print(
            f"IP：{d.get('query')}\n"
            f"國家：{d.get('country')} ({d.get('countryCode')})\n"
            f"城市：{d.get('city')}  地區：{d.get('regionName')}\n"
            f"ISP：{d.get('isp')}\n"
            f"組織：{d.get('org')}\n"
            f"時區：{d.get('timezone')}\n"
            f"座標：{d.get('lat')}, {d.get('lon')}"
        )
    except Exception as e:
        print(f"❌ IP 查詢失敗：{e}")


# ── 匯率查詢 ─────────────────────────────────────────

def currency(amount: float, from_cur: str, to_cur: str):
    try:
        res = requests.get(
            f"https://api.frankfurter.app/latest?amount={amount}&from={from_cur.upper()}&to={to_cur.upper()}",
            timeout=10
        )
        data = res.json()
        rate = data["rates"].get(to_cur.upper())
        if rate is None:
            print(f"❌ 找不到 {to_cur} 匯率")
            return
        print(f"💱 {amount} {from_cur.upper()} = {rate:.4f} {to_cur.upper()}")
        print(f"（匯率基準日：{data.get('date')}）")
    except Exception as e:
        print(f"❌ 匯率查詢失敗：{e}")


# ── Windows 事件日誌 ─────────────────────────────────

def event_log(log_name: str = "System", level: str = "Error", count: int = 10):
    try:
        ps = (
            f"Get-WinEvent -LogName '{log_name}' -MaxEvents {count} "
            f"| Where-Object {{$_.LevelDisplayName -eq '{level}'}} "
            f"| Select-Object TimeCreated,Id,Message "
            f"| Format-List"
        )
        r = subprocess.run(["powershell.exe", "-Command", ps],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30)
        print(r.stdout[:3000] or f"（無 {level} 等級事件）")
    except Exception as e:
        print(f"❌ 事件日誌查詢失敗：{e}")


# ── 進階 TTS (edge-tts) ──────────────────────────────

def tts_edge(text: str, voice: str = "zh-TW-HsiaoChenNeural", output: str = ""):
    try:
        import edge_tts, asyncio
        out = output or str(Path.home() / "Desktop" / f"tts_{datetime.now().strftime('%H%M%S')}.mp3")
        _clean = clean_for_tts(text)
        async def _gen():
            communicate = edge_tts.Communicate(_clean, voice)
            await communicate.save(out)
        asyncio.run(_gen())
        subprocess.Popen(["powershell.exe", "-Command", f"Start-Process '{out}'"])
        print(f"✅ 語音已生成並播放：{out}")
    except Exception as e:
        print(f"❌ Edge TTS 失敗：{e}")

def tts_voices():
    try:
        import edge_tts, asyncio
        async def _list():
            return await edge_tts.list_voices()
        voices = asyncio.run(_list())
        zh_voices = [v for v in voices if v["Locale"].startswith("zh")]
        for v in zh_voices:
            print(f"{v['ShortName']}  {v['FriendlyName']}")
    except Exception as e:
        print(f"❌ 列出語音失敗：{e}")


# ── 郵件附件發送 ─────────────────────────────────────

def send_email_attach(to: str, subject: str, body: str, attachments: str = ""):
    try:
        import smtplib
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText
        from email.mime.base import MIMEBase
        from email import encoders
        smtp_user = os.getenv("SMTP_USER", "")
        smtp_pass = os.getenv("SMTP_PASS", "")
        smtp_host = os.getenv("SMTP_HOST", "smtp.gmail.com")
        smtp_port = int(os.getenv("SMTP_PORT", "587"))
        msg = MIMEMultipart()
        msg["From"] = smtp_user
        msg["To"] = to
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain", "utf-8"))
        if attachments:
            for path in attachments.split(","):
                path = path.strip()
                if Path(path).exists():
                    with open(path, "rb") as f:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename={Path(path).name}")
                    msg.attach(part)
        with smtplib.SMTP(smtp_host, smtp_port) as s:
            s.starttls()
            s.login(smtp_user, smtp_pass)
            s.send_message(msg)
        print(f"✅ 郵件已發送至 {to}（附件：{attachments or '無'}）")
    except Exception as e:
        print(f"❌ 發送失敗：{e}")


# ── 剪貼簿圖片 ───────────────────────────────────────

def clipboard_img_get(output: str = ""):
    try:
        import win32clipboard
        from PIL import Image
        import io as _io
        win32clipboard.OpenClipboard()
        try:
            data = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
        finally:
            win32clipboard.CloseClipboard()
        out = output or str(Path.home() / "Desktop" / f"clipboard_{datetime.now().strftime('%H%M%S')}.png")
        img = Image.open(_io.BytesIO(data))
        img.save(out)
        print(f"✅ 剪貼簿圖片已存：{out}")
    except Exception as e:
        print(f"❌ 讀取剪貼簿圖片失敗：{e}")

def clipboard_img_set(path: str):
    try:
        import win32clipboard
        from PIL import Image
        import io as _io
        img = Image.open(path).convert("RGB")
        buf = _io.BytesIO()
        img.save(buf, "BMP")
        data = buf.getvalue()[14:]
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
        finally:
            win32clipboard.CloseClipboard()
        print(f"✅ 圖片已複製到剪貼簿：{path}")
    except Exception as e:
        print(f"❌ 設定剪貼簿圖片失敗：{e}")


# ── USB 裝置管理 ─────────────────────────────────────

def usb_list():
    try:
        r = subprocess.run(["powershell.exe", "-Command",
            "Get-PnpDevice | Where-Object {$_.Class -eq 'USB' -and $_.Status -eq 'OK'} | Select-Object FriendlyName,InstanceId | Format-Table -AutoSize"],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(r.stdout[:3000] or "（無 USB 裝置）")
    except Exception as e:
        print(f"❌ USB 查詢失敗：{e}")


# ── Windows 防火牆 ───────────────────────────────────

def firewall(action: str, name: str = "", direction: str = "in", port: int = 0, protocol: str = "TCP"):
    try:
        if action == "list":
            r = subprocess.run(["powershell.exe", "-Command",
                "Get-NetFirewallRule | Where-Object {$_.Enabled -eq 'True'} | Select-Object DisplayName,Direction,Action | Format-Table -AutoSize"],
                capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout[:3000])
        elif action == "add":
            ps = f"New-NetFirewallRule -DisplayName '{name}' -Direction {direction} -Protocol {protocol} -LocalPort {port} -Action Allow"
            r = subprocess.run(["powershell.exe", "-Command", ps], capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout or f"✅ 防火牆規則已新增：{name}")
        elif action == "remove":
            r = subprocess.run(["powershell.exe", "-Command", f"Remove-NetFirewallRule -DisplayName '{name}'"],
                capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout or f"✅ 防火牆規則已刪除：{name}")
    except Exception as e:
        print(f"❌ 防火牆操作失敗：{e}")


# ── 任務清單 ─────────────────────────────────────────

TODO_DB = Path.home() / "claude-telegram-bot" / "todo.db"

def _todo_init():
    conn = sqlite3.connect(str(TODO_DB))
    conn.execute("CREATE TABLE IF NOT EXISTS todos (id INTEGER PRIMARY KEY AUTOINCREMENT, task TEXT, done INTEGER DEFAULT 0, created TEXT)")
    conn.commit(); conn.close()

def todo(action: str, task: str = "", todo_id: int = 0):
    try:
        _todo_init()
        conn = sqlite3.connect(str(TODO_DB))
        cur = conn.cursor()
        if action == "add":
            cur.execute("INSERT INTO todos (task, created) VALUES (?, ?)", (task, datetime.now().strftime("%Y-%m-%d %H:%M")))
            conn.commit()
            print(f"✅ 已新增任務：{task}")
        elif action == "list":
            rows = cur.execute("SELECT id, task, done, created FROM todos ORDER BY done, id").fetchall()
            if not rows: print("（清單為空）"); conn.close(); return
            for r in rows:
                status = "✅" if r[2] else "⬜"
                print(f"{status} [{r[0]}] {r[1]}  ({r[3]})")
        elif action == "done":
            cur.execute("UPDATE todos SET done=1 WHERE id=?", (todo_id,))
            conn.commit(); print(f"✅ 任務 #{todo_id} 已完成")
        elif action == "delete":
            cur.execute("DELETE FROM todos WHERE id=?", (todo_id,))
            conn.commit(); print(f"✅ 任務 #{todo_id} 已刪除")
        elif action == "clear":
            cur.execute("DELETE FROM todos WHERE done=1")
            conn.commit(); print("✅ 已清除所有已完成任務")
        conn.close()
    except Exception as e:
        print(f"❌ 任務清單操作失敗：{e}")


# ── 檔案同步 ─────────────────────────────────────────

def file_sync(src: str, dest: str, dry_run: bool = False):
    try:
        import filecmp, shutil
        src_path = Path(src); dest_path = Path(dest)
        dest_path.mkdir(parents=True, exist_ok=True)
        copied = deleted = 0
        for item in src_path.rglob("*"):
            rel = item.relative_to(src_path)
            dst = dest_path / rel
            if item.is_dir():
                dst.mkdir(parents=True, exist_ok=True)
            elif item.is_file():
                if not dst.exists() or not filecmp.cmp(str(item), str(dst), shallow=False):
                    if not dry_run:
                        shutil.copy2(str(item), str(dst))
                    print(f"  複製：{rel}")
                    copied += 1
        print(f"{'[預覽] ' if dry_run else ''}✅ 同步完成：{copied} 個檔案更新，{deleted} 個刪除")
    except Exception as e:
        print(f"❌ 檔案同步失敗：{e}")


# ── 系統資源圖表 ─────────────────────────────────────

def sysres_chart(duration: int = 10, output: str = ""):
    try:
        import psutil, matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt, time as t
        cpu_vals, mem_vals, times = [], [], []
        for i in range(duration):
            cpu_vals.append(psutil.cpu_percent(interval=1))
            mem_vals.append(psutil.virtual_memory().percent)
            times.append(i + 1)
        out = output or str(Path.home() / "Desktop" / f"sysres_{datetime.now().strftime('%H%M%S')}.png")
        fig, ax = plt.subplots()
        ax.plot(times, cpu_vals, label="CPU %", color="blue")
        ax.plot(times, mem_vals, label="RAM %", color="orange")
        ax.set_ylim(0, 100); ax.set_xlabel("秒"); ax.set_ylabel("%")
        ax.set_title("系統資源使用率"); ax.legend(); plt.tight_layout()
        plt.savefig(out); plt.close()
        print(f"✅ 資源圖表已存：{out}")
    except Exception as e:
        print(f"❌ 資源圖表失敗：{e}")


# ── 密碼管理 ─────────────────────────────────────────

PWD_DB = Path.home() / "claude-telegram-bot" / "passwords.db"

def _pwd_get_key(master: str) -> bytes:
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    import base64
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=b"pwd_manager_v1", iterations=100000)
    return base64.urlsafe_b64encode(kdf.derive(master.encode()))

def password_save(site: str, username: str, password: str, master: str):
    try:
        from cryptography.fernet import Fernet
        key = _pwd_get_key(master)
        f = Fernet(key)
        encrypted = f.encrypt(password.encode()).decode()
        conn = sqlite3.connect(str(PWD_DB))
        conn.execute("CREATE TABLE IF NOT EXISTS passwords (id INTEGER PRIMARY KEY AUTOINCREMENT, site TEXT, username TEXT, password TEXT)")
        conn.execute("INSERT OR REPLACE INTO passwords (site, username, password) VALUES (?, ?, ?)", (site, username, encrypted))
        conn.commit(); conn.close()
        print(f"✅ 密碼已加密儲存：{site} / {username}")
    except Exception as e:
        print(f"❌ 儲存密碼失敗：{e}")

def password_get(site: str, master: str):
    try:
        from cryptography.fernet import Fernet
        key = _pwd_get_key(master)
        f = Fernet(key)
        conn = sqlite3.connect(str(PWD_DB))
        rows = conn.execute("SELECT site, username, password FROM passwords WHERE site LIKE ?", (f"%{site}%",)).fetchall()
        conn.close()
        if not rows: print(f"（找不到 {site} 的密碼）"); return
        for site_, user, enc_pwd in rows:
            try:
                pwd = f.decrypt(enc_pwd.encode()).decode()
                print(f"🔑 {site_}\n帳號：{user}\n密碼：{pwd}")
            except Exception:
                print(f"❌ 解密失敗（主密碼錯誤？）")
    except Exception as e:
        print(f"❌ 讀取密碼失敗：{e}")


# ── RDP 遠端桌面 ─────────────────────────────────────

def rdp_connect(host: str, user: str = "", width: int = 1280, height: int = 720):
    try:
        args = ["/v:" + host, f"/w:{width}", f"/h:{height}"]
        if user:
            args.append(f"/u:{user}")
        subprocess.Popen(["mstsc"] + args)
        print(f"✅ 正在連線 RDP：{host}")
    except Exception as e:
        print(f"❌ RDP 連線失敗：{e}")


# ── Chrome 書籤 ──────────────────────────────────────

def chrome_bookmarks():
    try:
        import json
        bookmark_path = Path.home() / "AppData/Local/Google/Chrome/User Data/Default/Bookmarks"
        if not bookmark_path.exists():
            print("❌ 找不到 Chrome 書籤檔案")
            return
        data = json.loads(bookmark_path.read_text(encoding="utf-8"))
        def _print_node(node, indent=0):
            if node.get("type") == "url":
                print("  " * indent + f"🔗 {node['name']}  {node['url']}")
            elif node.get("type") == "folder":
                print("  " * indent + f"📁 {node['name']}")
                for child in node.get("children", []):
                    _print_node(child, indent + 1)
        for root in data["roots"].values():
            _print_node(root)
    except Exception as e:
        print(f"❌ 讀取書籤失敗：{e}")


# ── 印表機管理 ───────────────────────────────────────

def printer_list():
    try:
        r = subprocess.run(["powershell.exe", "-Command",
            "Get-Printer | Select-Object Name,DriverName,PrinterStatus | Format-Table -AutoSize"],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(r.stdout or "（無印表機）")
    except Exception as e:
        print(f"❌ 印表機查詢失敗：{e}")

def printer_jobs():
    try:
        r = subprocess.run(["powershell.exe", "-Command",
            "Get-PrintJob -PrinterName (Get-Printer | Select-Object -First 1 -ExpandProperty Name) | Format-Table"],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(r.stdout or "（列印佇列為空）")
    except Exception as e:
        print(f"❌ 列印佇列查詢失敗：{e}")


# ── 網路芳鄰 ─────────────────────────────────────────

def net_share(action: str, share_path: str = "", drive: str = "Z:", user: str = "", password: str = ""):
    try:
        if action == "list":
            r = subprocess.run(["net", "use"], capture_output=True, text=True, encoding="cp950", errors="replace")
            print(r.stdout)
        elif action == "connect":
            args = ["net", "use", drive, share_path]
            if user: args += [f"/user:{user}", password]
            r = subprocess.run(args, capture_output=True, text=True, encoding="cp950", errors="replace")
            print(r.stdout or f"✅ 已連線 {share_path} → {drive}")
        elif action == "disconnect":
            r = subprocess.run(["net", "use", drive, "/delete"], capture_output=True, text=True, encoding="cp950", errors="replace")
            print(r.stdout or f"✅ 已中斷 {drive}")
    except Exception as e:
        print(f"❌ 網路芳鄰操作失敗：{e}")


# ── 字型列表 ─────────────────────────────────────────

def font_list(keyword: str = ""):
    try:
        r = subprocess.run(["powershell.exe", "-Command",
            "[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') | Out-Null; "
            "[System.Drawing.FontFamily]::Families | Select-Object -ExpandProperty Name"],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        fonts = [f for f in r.stdout.strip().splitlines() if keyword.lower() in f.lower()] if keyword else r.stdout.strip().splitlines()
        print("\n".join(fonts[:50]))
        if len(fonts) > 50:
            print(f"...（共 {len(fonts)} 個字型）")
    except Exception as e:
        print(f"❌ 字型列表失敗：{e}")


# ── 音量控制 ─────────────────────────────────────────

def volume_get():
    try:
        from ctypes import cast, POINTER
        from comtypes import CLSCTX_ALL
        from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        muted = volume.GetMute()
        level = round(volume.GetMasterVolumeLevelScalar() * 100)
        print(f"🔊 音量：{level}%  {'（靜音）' if muted else ''}")
    except Exception as e:
        print(f"❌ 音量查詢失敗：{e}")

def volume_set(level: int):
    try:
        from ctypes import cast, POINTER
        from comtypes import CLSCTX_ALL
        from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevelScalar(max(0.0, min(1.0, level / 100)), None)
        print(f"✅ 音量已設定為 {level}%")
    except Exception as e:
        print(f"❌ 設定音量失敗：{e}")

def volume_mute(mute: bool = True):
    try:
        from ctypes import cast, POINTER
        from comtypes import CLSCTX_ALL
        from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMute(1 if mute else 0, None)
        print(f"✅ {'靜音' if mute else '取消靜音'}")
    except Exception as e:
        print(f"❌ 靜音操作失敗：{e}")


# ── 螢幕亮度/解析度 ──────────────────────────────────

def brightness_get():
    try:
        import screen_brightness_control as sbc
        b = sbc.get_brightness()
        print(f"💡 亮度：{b}%")
    except Exception as e:
        print(f"❌ 亮度查詢失敗：{e}")

def brightness_set(level: int):
    try:
        import screen_brightness_control as sbc
        sbc.set_brightness(max(0, min(100, level)))
        print(f"✅ 亮度已設定為 {level}%")
    except Exception as e:
        print(f"❌ 設定亮度失敗：{e}")

def resolution_list():
    try:
        r = subprocess.run(["powershell.exe", "-Command",
            "Get-CimInstance -ClassName Win32_VideoController | Select-Object CurrentHorizontalResolution,CurrentVerticalResolution,CurrentRefreshRate | Format-List"],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(r.stdout.strip())
    except Exception as e:
        print(f"❌ 解析度查詢失敗：{e}")


# ── 媒體播放控制 ─────────────────────────────────────

def media_control(action: str):
    """action: play_pause / next / prev / stop / volume_up / volume_down / mute"""
    try:
        import keyboard as kb
        key_map = {
            "play_pause": "play/pause media",
            "next": "next track",
            "prev": "previous track",
            "stop": "stop media",
            "volume_up": "volume up",
            "volume_down": "volume down",
            "mute": "volume mute",
        }
        key = key_map.get(action)
        if not key:
            print(f"❌ 未知動作：{action}，可用：{list(key_map.keys())}")
            return
        kb.send(key)
        print(f"✅ 媒體控制：{action}")
    except Exception as e:
        print(f"❌ 媒體控制失敗：{e}")


# ── 音訊裝置切換 ─────────────────────────────────────

def audio_devices():
    try:
        from pycaw.pycaw import AudioUtilities
        devices = AudioUtilities.GetAllDevices()
        for i, d in enumerate(devices):
            print(f"[{i}] {d.FriendlyName}")
    except Exception as e:
        print(f"❌ 音訊裝置查詢失敗：{e}")

def audio_switch(device_name: str):
    try:
        r = subprocess.run(["powershell.exe", "-Command",
            f"Get-AudioDevice -List | Where-Object {{$_.Name -like '*{device_name}*'}} | Set-AudioDevice"],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(r.stdout or f"✅ 已切換至：{device_name}")
    except Exception as e:
        print(f"❌ 切換音訊裝置失敗（需安裝 AudioDeviceCmdlets）：{e}")


# ── 已安裝軟體管理 ───────────────────────────────────

def software_list(keyword: str = ""):
    try:
        r = subprocess.run(["powershell.exe", "-Command",
            "Get-Package | Select-Object Name,Version | Sort-Object Name | Format-Table -AutoSize"],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30)
        lines = r.stdout.strip().splitlines()
        if keyword:
            lines = [l for l in lines if keyword.lower() in l.lower()]
        print("\n".join(lines[:50]))
        if len(lines) > 50:
            print(f"...（共 {len(lines)} 個）")
    except Exception as e:
        print(f"❌ 軟體列表查詢失敗：{e}")

def software_install(name: str):
    try:
        r = subprocess.run(["powershell.exe", "-Command", f"winget install --id '{name}' --silent --accept-source-agreements --accept-package-agreements"],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=300)
        print(r.stdout or r.stderr or f"安裝完成：{name}")
    except Exception as e:
        print(f"❌ 安裝失敗：{e}")

def software_uninstall(name: str):
    try:
        r = subprocess.run(["powershell.exe", "-Command", f"winget uninstall --name '{name}' --silent"],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=300)
        print(r.stdout or r.stderr or f"✅ 已卸載：{name}")
    except Exception as e:
        print(f"❌ 卸載失敗：{e}")


# ── 開機自啟動管理 ───────────────────────────────────

def startup_list():
    try:
        r = subprocess.run(["powershell.exe", "-Command",
            "Get-CimInstance Win32_StartupCommand | Select-Object Name,Command,Location | Format-Table -AutoSize"],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(r.stdout.strip() or "（無開機自啟動程式）")
    except Exception as e:
        print(f"❌ 查詢失敗：{e}")

def startup_add(name: str, command: str):
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, name, 0, winreg.REG_SZ, command)
        winreg.CloseKey(key)
        print(f"✅ 已新增開機自啟動：{name}")
    except Exception as e:
        print(f"❌ 新增失敗：{e}")

def startup_remove(name: str):
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
        winreg.DeleteValue(key, name)
        winreg.CloseKey(key)
        print(f"✅ 已移除開機自啟動：{name}")
    except Exception as e:
        print(f"❌ 移除失敗：{e}")


# ── 環境變數管理 ─────────────────────────────────────

def env_get(name: str = ""):
    try:
        if name:
            val = os.environ.get(name, "（未設定）")
            print(f"{name} = {val}")
        else:
            for k, v in sorted(os.environ.items()):
                print(f"{k} = {v[:80]}")
    except Exception as e:
        print(f"❌ 環境變數查詢失敗：{e}")

def env_set(name: str, value: str, permanent: bool = False):
    try:
        os.environ[name] = value
        if permanent:
            r = subprocess.run(["powershell.exe", "-Command",
                f"[System.Environment]::SetEnvironmentVariable('{name}','{value}','User')"],
                capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(f"✅ 環境變數已永久設定：{name} = {value}")
        else:
            print(f"✅ 環境變數已設定（本次）：{name} = {value}")
    except Exception as e:
        print(f"❌ 設定環境變數失敗：{e}")


# ── 使用者帳戶管理 ───────────────────────────────────

def user_list():
    try:
        r = subprocess.run(["powershell.exe", "-Command",
            "Get-LocalUser | Select-Object Name,Enabled,LastLogon | Format-Table -AutoSize"],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(r.stdout.strip())
    except Exception as e:
        print(f"❌ 查詢失敗：{e}")

def user_create(username: str, password: str):
    try:
        ps = f"New-LocalUser -Name '{username}' -Password (ConvertTo-SecureString '{password}' -AsPlainText -Force) -FullName '{username}'"
        r = subprocess.run(["powershell.exe", "-Command", ps],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(r.stdout or f"✅ 使用者已建立：{username}")
    except Exception as e:
        print(f"❌ 建立使用者失敗：{e}")

def user_delete(username: str):
    try:
        r = subprocess.run(["powershell.exe", "-Command", f"Remove-LocalUser -Name '{username}'"],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(r.stdout or f"✅ 使用者已刪除：{username}")
    except Exception as e:
        print(f"❌ 刪除使用者失敗：{e}")


# ── Windows 更新 ─────────────────────────────────────

def win_update(action: str = "list"):
    try:
        if action == "list":
            r = subprocess.run(["powershell.exe", "-Command",
                "Get-WindowsUpdate -AcceptAll 2>$null | Select-Object Title,Size | Format-Table -AutoSize"],
                capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=60)
            print(r.stdout or "（需安裝 PSWindowsUpdate 模組）\n執行：Install-Module PSWindowsUpdate -Force")
        elif action == "install":
            r = subprocess.run(["powershell.exe", "-Command",
                "Install-WindowsUpdate -AcceptAll -AutoReboot:$false 2>&1"],
                capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=600)
            print(r.stdout or "✅ 更新完成")
        elif action == "check":
            r = subprocess.run(["powershell.exe", "-Command",
                "(New-Object -ComObject Microsoft.Update.AutoUpdate).DetectNow()"],
                capture_output=True, text=True, encoding="utf-8", errors="replace")
            print("✅ 已觸發 Windows Update 檢查")
    except Exception as e:
        print(f"❌ Windows 更新失敗：{e}")


# ── 裝置管理員 ───────────────────────────────────────

def device_list(keyword: str = ""):
    try:
        ps = "Get-PnpDevice | Select-Object Status,Class,FriendlyName | Format-Table -AutoSize"
        r = subprocess.run(["powershell.exe", "-Command", ps],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        lines = r.stdout.strip().splitlines()
        if keyword:
            lines = [l for l in lines if keyword.lower() in l.lower()]
        print("\n".join(lines[:50]))
    except Exception as e:
        print(f"❌ 裝置查詢失敗：{e}")

def device_toggle(name: str, enable: bool = True):
    try:
        action = "Enable" if enable else "Disable"
        r = subprocess.run(["powershell.exe", "-Command",
            f"{action}-PnpDevice -FriendlyName '*{name}*' -Confirm:$false"],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(r.stdout or f"✅ 裝置已{'啟用' if enable else '停用'}：{name}")
    except Exception as e:
        print(f"❌ 裝置操作失敗：{e}")


# ── 網路介面卡控制 ───────────────────────────────────

def netadapter_list():
    try:
        r = subprocess.run(["powershell.exe", "-Command",
            "Get-NetAdapter | Select-Object Name,Status,MacAddress,LinkSpeed | Format-Table -AutoSize"],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(r.stdout.strip())
    except Exception as e:
        print(f"❌ 查詢失敗：{e}")

def netadapter_toggle(name: str, enable: bool = True):
    try:
        action = "Enable" if enable else "Disable"
        r = subprocess.run(["powershell.exe", "-Command", f"{action}-NetAdapter -Name '{name}' -Confirm:$false"],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(r.stdout or f"✅ 網路卡已{'啟用' if enable else '停用'}：{name}")
    except Exception as e:
        print(f"❌ 網路卡操作失敗：{e}")


# ── DNS/IP 設定 ──────────────────────────────────────

def dns_get():
    try:
        r = subprocess.run(["powershell.exe", "-Command",
            "Get-DnsClientServerAddress | Select-Object InterfaceAlias,ServerAddresses | Format-Table -AutoSize"],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(r.stdout.strip())
    except Exception as e:
        print(f"❌ 查詢失敗：{e}")

def dns_set(interface: str, dns1: str, dns2: str = ""):
    try:
        servers = f"'{dns1}'" + (f",'{dns2}'" if dns2 else "")
        r = subprocess.run(["powershell.exe", "-Command",
            f"Set-DnsClientServerAddress -InterfaceAlias '{interface}' -ServerAddresses ({servers})"],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(r.stdout or f"✅ DNS 已設定：{dns1}{', '+dns2 if dns2 else ''}")
    except Exception as e:
        print(f"❌ DNS 設定失敗：{e}")

def ip_config(interface: str, ip: str, mask: str = "255.255.255.0", gateway: str = ""):
    try:
        ps = f"New-NetIPAddress -InterfaceAlias '{interface}' -IPAddress '{ip}' -PrefixLength 24 -DefaultGateway '{gateway}'"
        r = subprocess.run(["powershell.exe", "-Command", ps],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(r.stdout or f"✅ IP 已設定：{ip}")
    except Exception as e:
        print(f"❌ IP 設定失敗：{e}")


# ── Hosts 檔案編輯 ───────────────────────────────────

HOSTS_PATH = r"C:\Windows\System32\drivers\etc\hosts"

def hosts_list():
    try:
        content = Path(HOSTS_PATH).read_text(encoding="utf-8", errors="replace")
        lines = [l for l in content.splitlines() if l.strip() and not l.strip().startswith("#")]
        print("\n".join(lines) or "（無自訂 hosts）")
    except Exception as e:
        print(f"❌ 讀取 hosts 失敗：{e}")

def hosts_add(ip: str, domain: str):
    try:
        content = Path(HOSTS_PATH).read_text(encoding="utf-8", errors="replace")
        entry = f"\n{ip}\t{domain}"
        if domain in content:
            print(f"⚠️ {domain} 已存在 hosts 中")
            return
        with open(HOSTS_PATH, "a", encoding="utf-8") as f:
            f.write(entry)
        print(f"✅ 已新增：{ip} → {domain}")
    except Exception as e:
        print(f"❌ 新增 hosts 失敗（需管理員權限）：{e}")

def hosts_remove(domain: str):
    try:
        content = Path(HOSTS_PATH).read_text(encoding="utf-8", errors="replace")
        lines = [l for l in content.splitlines() if domain not in l]
        Path(HOSTS_PATH).write_text("\n".join(lines), encoding="utf-8")
        print(f"✅ 已移除：{domain}")
    except Exception as e:
        print(f"❌ 移除 hosts 失敗（需管理員權限）：{e}")


# ── 網路流量監控 ─────────────────────────────────────

def net_traffic(duration: int = 5):
    try:
        import psutil, time as t
        before = psutil.net_io_counters(pernic=False)
        t.sleep(duration)
        after = psutil.net_io_counters(pernic=False)
        sent = (after.bytes_sent - before.bytes_sent) / duration / 1024
        recv = (after.bytes_recv - before.bytes_recv) / duration / 1024
        print(f"📡 網路流量（{duration}秒平均）\n上傳：{sent:.1f} KB/s\n下載：{recv:.1f} KB/s")
        print("\n各網路介面：")
        per_nic = psutil.net_io_counters(pernic=True)
        for name, stats in per_nic.items():
            print(f"  {name}: ↑{stats.bytes_sent/1024/1024:.1f}MB ↓{stats.bytes_recv/1024/1024:.1f}MB")
    except Exception as e:
        print(f"❌ 網路流量監控失敗：{e}")


# ── 條件式自動化 ─────────────────────────────────────

def if_then(condition_type: str, condition_value: str, action_cmd: str, duration: float = 300.0):
    """
    condition_type: cpu_above / mem_above / file_exists / time_is / process_running
    condition_value: 數值或字串
    action_cmd: shell 指令
    """
    try:
        import psutil, time as t
        print(f"🤖 條件式自動化啟動：if [{condition_type}={condition_value}] → [{action_cmd}]")
        end = t.time() + duration
        triggered = False
        while t.time() < end:
            met = False
            if condition_type == "cpu_above":
                met = psutil.cpu_percent(interval=1) > float(condition_value)
            elif condition_type == "mem_above":
                met = psutil.virtual_memory().percent > float(condition_value)
            elif condition_type == "file_exists":
                met = Path(condition_value).exists()
            elif condition_type == "time_is":
                met = datetime.now().strftime("%H:%M") == condition_value
            elif condition_type == "process_running":
                met = any(p.name().lower() == condition_value.lower() for p in psutil.process_iter())
            if met and not triggered:
                subprocess.run(action_cmd, shell=True)
                triggered = True
                print(f"✅ [{datetime.now().strftime('%H:%M:%S')}] 條件觸發，已執行：{action_cmd}")
                break
            t.sleep(2)
        if not triggered:
            print(f"⏰ 監控結束，條件未觸發")
    except Exception as e:
        print(f"❌ 條件式自動化失敗：{e}")


# ── 多視窗排列 ───────────────────────────────────────

def window_arrange(layout: str = "side_by_side"):
    """layout: side_by_side / quad / stack / maximize_all"""
    try:
        import win32gui, win32con
        sw = pyautogui.size().width
        sh = pyautogui.size().height
        hwnds = []
        def _enum(hwnd, _):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                hwnds.append(hwnd)
        win32gui.EnumWindows(_enum, None)
        hwnds = [h for h in hwnds if win32gui.GetWindowText(h).strip()][:8]
        if layout == "side_by_side" and len(hwnds) >= 2:
            w = sw // 2
            win32gui.MoveWindow(hwnds[0], 0, 0, w, sh, True)
            win32gui.MoveWindow(hwnds[1], w, 0, w, sh, True)
            print(f"✅ 左右分割：{win32gui.GetWindowText(hwnds[0])} | {win32gui.GetWindowText(hwnds[1])}")
        elif layout == "quad" and len(hwnds) >= 4:
            w, h = sw // 2, sh // 2
            positions = [(0,0),(w,0),(0,h),(w,h)]
            for i, (x,y) in enumerate(positions):
                win32gui.MoveWindow(hwnds[i], x, y, w, h, True)
            print(f"✅ 四格排列完成")
        elif layout == "stack":
            h = sh // len(hwnds)
            for i, hwnd in enumerate(hwnds[:4]):
                win32gui.MoveWindow(hwnd, 0, i * h, sw, h, True)
            print(f"✅ 堆疊排列完成")
        elif layout == "maximize_all":
            for hwnd in hwnds:
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            print(f"✅ 已最大化 {len(hwnds)} 個視窗")
    except Exception as e:
        print(f"❌ 視窗排列失敗：{e}")


# ── 指定區域 OCR ─────────────────────────────────────

def region_ocr(x: int, y: int, w: int, h: int, lang: str = "ch_tra"):
    try:
        import easyocr, numpy as np
        from PIL import Image
        img = pyautogui.screenshot(region=(x, y, w, h))
        reader = easyocr.Reader([lang, "en"], gpu=False)
        results = reader.readtext(np.array(img))
        text = "\n".join(r[1] for r in results)
        print(text or "（未辨識到文字）")
    except Exception as e:
        print(f"❌ 區域 OCR 失敗：{e}")


# ── 指定視窗截圖 ─────────────────────────────────────

def window_screenshot(title_keyword: str, output: str = ""):
    try:
        import win32gui, win32ui, win32con
        from ctypes import windll
        import numpy as np
        from PIL import Image
        hwnd = None
        def _find(h, _):
            nonlocal hwnd
            if title_keyword.lower() in win32gui.GetWindowText(h).lower():
                hwnd = h
        win32gui.EnumWindows(_find, None)
        if not hwnd:
            print(f"❌ 找不到視窗：{title_keyword}")
            return
        win32gui.SetForegroundWindow(hwnd)
        rect = win32gui.GetWindowRect(hwnd)
        x, y, x2, y2 = rect
        w, h = x2 - x, y2 - y
        hwndDC = win32gui.GetWindowDC(hwnd)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
        saveDC.SelectObject(saveBitMap)
        windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 3)
        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)
        img = Image.frombuffer("RGB", (bmpinfo["bmWidth"], bmpinfo["bmHeight"]), bmpstr, "raw", "BGRX", 0, 1)
        out = output or str(Path.home() / "Desktop" / f"win_{title_keyword[:10]}_{datetime.now().strftime('%H%M%S')}.png")
        img.save(out)
        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC(); mfcDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)
        print(f"✅ 視窗截圖已存：{out}")
    except Exception as e:
        print(f"❌ 視窗截圖失敗：{e}")


# ── Wave 13：檔案監聽/像素監控/AI物件偵測/滑鼠錄製/ADB/WiFi熱點/OneDrive/FTP/WSL/Hyper-V/檔案diff/螢幕串流 ──

def file_watcher(path: str, events: str = "all", command: str = "", duration: float = 3600.0):
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
    import time
    class _H(FileSystemEventHandler):
        def _h(self, e, t):
            if events != "all" and t not in events: return
            print(f"📁 {t}：{e.src_path}")
            if command: import subprocess; subprocess.Popen(command.replace("{path}", e.src_path), shell=True)
        def on_created(self, e): self._h(e, "created")
        def on_modified(self, e): self._h(e, "modified")
        def on_deleted(self, e): self._h(e, "deleted")
    obs = Observer(); obs.schedule(_H(), path, recursive=True); obs.start()
    print(f"👁️ 監聽 {path} 中（{duration}s）...")
    try: time.sleep(duration)
    except KeyboardInterrupt: pass
    finally: obs.stop(); obs.join()

def pixel_watch(x: int, y: int, tolerance: int = 10, interval: float = 1.0, duration: float = 60.0, command: str = ""):
    import pyautogui, time, subprocess
    sc = pyautogui.screenshot(); r0,g0,b0 = sc.getpixel((x,y))[:3]
    print(f"🎨 開始監控({x},{y}) 初始顏色：#{r0:02X}{g0:02X}{b0:02X}")
    end = time.time() + duration
    while time.time() < end:
        sc = pyautogui.screenshot(); r,g,b = sc.getpixel((x,y))[:3]
        if abs(r-r0)+abs(g-g0)+abs(b-b0) > tolerance*3:
            print(f"🔔 顏色變化！#{r:02X}{g:02X}{b:02X}")
            if command: subprocess.Popen(command, shell=True)
            r0,g0,b0 = r,g,b
        time.sleep(interval)

def object_detect(target: str, action: str = "find", region: str = ""):
    import pyautogui, anthropic, base64, io as _io, json, re, os
    reg = None
    if region:
        parts = [int(v) for v in region.split(",")]
        if len(parts) == 4: reg = tuple(parts)
    sc = pyautogui.screenshot(region=reg)
    buf = _io.BytesIO(); sc.save(buf, format="PNG")
    img_b64 = base64.standard_b64encode(buf.getvalue()).decode()
    ox = reg[0] if reg else 0; oy = reg[1] if reg else 0
    cl = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    resp = cl.messages.create(model="claude-sonnet-4-6", max_tokens=256, messages=[{"role":"user","content":[
        {"type":"image","source":{"type":"base64","media_type":"image/png","data":img_b64}},
        {"type":"text","text":f"找「{target}」位置，回 JSON：{{\"found\":bool,\"x\":int,\"y\":int,\"description\":\"\"}}"}
    ]}])
    m = re.search(r'\{.*\}', resp.content[0].text, re.DOTALL)
    if not m: print("⚠️ 無法解析"); return
    res = json.loads(m.group())
    if not res.get("found"): print(f"⚠️ 未找到：{target}"); return
    ax, ay = res["x"]+ox, res["y"]+oy
    if action == "click": pyautogui.click(ax, ay); print(f"✅ 已點擊「{target}」({ax},{ay})")
    elif action == "double_click": pyautogui.doubleClick(ax, ay); print(f"✅ 已雙擊({ax},{ay})")
    else: print(f"✅ 找到「{target}」：({ax},{ay}) {res.get('description','')}")

def mouse_record(action: str, name: str = "", duration: float = 10.0, repeat: int = 1, speed: float = 1.0):
    import json, time
    f = Path.home() / ".claude_mouse_macros.json"
    store = json.loads(f.read_text()) if f.exists() else {}
    if action == "list":
        print("\n".join(f"- {k}（{len(v)}事件）" for k,v in store.items()) if store else "⚠️ 無已儲存巨集")
    elif action == "delete":
        if name in store: del store[name]; f.write_text(json.dumps(store)); print(f"✅ 已刪除：{name}")
        else: print(f"⚠️ 找不到：{name}")
    elif action == "start":
        from pynput import mouse as m
        events = []; t0 = time.time()
        def on_move(x,y): events.append({"t":time.time()-t0,"type":"move","x":x,"y":y})
        def on_click(x,y,btn,pressed): events.append({"t":time.time()-t0,"type":"click","x":x,"y":y,"btn":str(btn),"pressed":pressed})
        def on_scroll(x,y,dx,dy): events.append({"t":time.time()-t0,"type":"scroll","x":x,"y":y,"dx":dx,"dy":dy})
        lis = m.Listener(on_move=on_move,on_click=on_click,on_scroll=on_scroll)
        lis.start(); time.sleep(duration); lis.stop()
        store[name] = events; f.write_text(json.dumps(store))
        print(f"✅ 已錄製 '{name}'（{len(events)}事件）")
    elif action == "play":
        if name not in store: print(f"⚠️ 找不到：{name}"); return
        from pynput import mouse as m
        ctrl = m.Controller(); evts = store[name]
        for _ in range(repeat):
            prev = 0.0
            for e in evts:
                d = (e["t"]-prev)/speed; time.sleep(min(max(d,0),0.5)); prev=e["t"]
                if e["type"]=="move": ctrl.position=(e["x"],e["y"])
                elif e["type"]=="click":
                    btn=m.Button.left if "left" in e["btn"] else m.Button.right
                    ctrl.press(btn) if e["pressed"] else ctrl.release(btn)
                elif e["type"]=="scroll": ctrl.scroll(e["dx"],e["dy"])
        print(f"✅ '{name}' 回放 {repeat} 次完成")

def adb(action: str, x: int=0, y: int=0, x2: int=0, y2: int=0, text: str="", path: str="", remote: str="", package: str="", command: str="", device: str=""):
    import subprocess
    p = ["adb"] + (["-s",device] if device else [])
    try:
        if action == "devices":
            r = subprocess.run(["adb","devices","-l"],capture_output=True,text=True); print(r.stdout.strip())
        elif action == "screenshot":
            out = path or str(Path.home()/"Desktop"/f"adb_{datetime.now().strftime('%H%M%S')}.png")
            subprocess.run(p+["shell","screencap","-p","/sdcard/screen.png"],capture_output=True)
            subprocess.run(p+["pull","/sdcard/screen.png",out],capture_output=True); print(f"✅ 截圖：{out}")
        elif action == "tap":
            subprocess.run(p+["shell","input","tap",str(x),str(y)],capture_output=True); print(f"✅ 點擊({x},{y})")
        elif action == "swipe":
            subprocess.run(p+["shell","input","swipe",str(x),str(y),str(x2),str(y2),"300"],capture_output=True); print(f"✅ 滑動")
        elif action == "type":
            subprocess.run(p+["shell","input","text",text.replace(" ","%s")],capture_output=True); print(f"✅ 輸入：{text}")
        elif action == "key":
            subprocess.run(p+["shell","input","keyevent",text],capture_output=True); print(f"✅ 按鍵：{text}")
        elif action == "install":
            r=subprocess.run(p+["install","-r",path],capture_output=True,text=True); print(r.stdout.strip())
        elif action == "push":
            r=subprocess.run(p+["push",path,remote or "/sdcard/"],capture_output=True,text=True); print(r.stdout.strip())
        elif action == "pull":
            out=path or str(Path.home()/"Desktop"/Path(remote).name)
            r=subprocess.run(p+["pull",remote,out],capture_output=True,text=True); print(r.stdout.strip())
        elif action == "shell":
            r=subprocess.run(p+["shell",command],capture_output=True,text=True); print(r.stdout.strip())
        elif action == "start_app":
            subprocess.run(p+["shell","monkey","-p",package,"-c","android.intent.category.LAUNCHER","1"],capture_output=True); print(f"✅ 啟動：{package}")
        elif action == "stop_app":
            subprocess.run(p+["shell","am","force-stop",package],capture_output=True); print(f"✅ 停止：{package}")
    except Exception as e: print(f"❌ ADB失敗：{e}")

def wifi_hotspot(action: str, ssid: str="", password: str=""):
    import subprocess
    try:
        if action == "set":
            r=subprocess.run(["netsh","wlan","set","hostednetwork",f"mode=allow",f"ssid={ssid}",f"key={password}"],capture_output=True,text=True,encoding="utf-8",errors="replace"); print(r.stdout.strip())
        elif action == "start":
            r=subprocess.run(["netsh","wlan","start","hostednetwork"],capture_output=True,text=True,encoding="utf-8",errors="replace"); print(r.stdout.strip())
        elif action == "stop":
            r=subprocess.run(["netsh","wlan","stop","hostednetwork"],capture_output=True,text=True,encoding="utf-8",errors="replace"); print(r.stdout.strip())
        elif action == "status":
            r=subprocess.run(["netsh","wlan","show","hostednetwork"],capture_output=True,text=True,encoding="utf-8",errors="replace"); print(r.stdout.strip())
    except Exception as e: print(f"❌ 熱點失敗：{e}")

def onedrive(action: str, path: str="", remote: str=""):
    import os, shutil
    od = os.path.expandvars(r"%USERPROFILE%\OneDrive")
    if not Path(od).exists(): od = str(Path.home()/"OneDrive")
    try:
        if action == "list":
            target = Path(od)/(remote or "")
            for p in sorted(target.iterdir()): print(f"{'📁' if p.is_dir() else '📄'} {p.name}")
        elif action == "upload":
            dest = Path(od)/(remote or Path(path).name); shutil.copy2(path, dest); print(f"✅ 已上傳：{dest}")
        elif action == "download":
            src = Path(od)/remote; out = path or str(Path.home()/"Desktop"/src.name); shutil.copy2(src, out); print(f"✅ 已下載：{out}")
        elif action == "status":
            size = sum(f.stat().st_size for f in Path(od).rglob("*") if f.is_file())/1024/1024/1024; print(f"OneDrive路徑：{od}\n使用：{size:.2f}GB")
        elif action == "open":
            os.startfile(od); print(f"✅ 已開啟：{od}")
    except Exception as e: print(f"❌ OneDrive失敗：{e}")

def ftp(action: str, host: str="", user: str="", password: str="", local: str="", remote: str="", port: int=21):
    from ftplib import FTP
    try:
        f = FTP(); f.connect(host, port, timeout=30); f.login(user, password)
        if action == "list":
            for item in f.nlst(remote or "."): print(item)
        elif action == "upload":
            with open(local,"rb") as fp: f.storbinary(f"STOR {remote or Path(local).name}", fp)
            print(f"✅ 已上傳：{local}")
        elif action == "download":
            out = local or str(Path.home()/"Desktop"/Path(remote).name)
            with open(out,"wb") as fp: f.retrbinary(f"RETR {remote}", fp.write)
            print(f"✅ 已下載：{out}")
        elif action == "delete":
            f.delete(remote); print(f"✅ 已刪除：{remote}")
        elif action == "mkdir":
            f.mkd(remote); print(f"✅ 已建立目錄：{remote}")
        f.quit()
    except Exception as e: print(f"❌ FTP失敗：{e}")

def wsl(action: str, distro: str="", command: str=""):
    import subprocess
    try:
        if action == "list":
            r=subprocess.run(["wsl","--list","--verbose"],capture_output=True,text=True,encoding="utf-16-le",errors="replace"); print(r.stdout.strip())
        elif action == "run":
            cmd=["wsl"]+(["-d",distro] if distro else [])+["--","bash","-c",command]
            r=subprocess.run(cmd,capture_output=True,text=True,encoding="utf-8",errors="replace"); print(r.stdout.strip())
        elif action == "start":
            subprocess.Popen(["wsl"]+(["-d",distro] if distro else [])); print(f"✅ WSL 已啟動")
        elif action == "stop":
            cmd=["wsl","--terminate",distro] if distro else ["wsl","--shutdown"]
            subprocess.run(cmd,capture_output=True); print(f"✅ WSL 已停止")
        elif action == "status":
            r=subprocess.run(["wsl","--status"],capture_output=True,text=True,encoding="utf-16-le",errors="replace"); print(r.stdout.strip())
    except Exception as e: print(f"❌ WSL失敗：{e}")

def hyperv(action: str, name: str="", snapshot: str=""):
    import subprocess
    def ps(cmd):
        r=subprocess.run(["powershell","-Command",cmd],capture_output=True,text=True,encoding="utf-8",errors="replace"); return r.stdout.strip(),r.returncode
    try:
        if action == "list": out,_=ps("Get-VM | Select-Object Name,State | Format-Table -AutoSize"); print(out)
        elif action == "start": ps(f"Start-VM -Name '{name}'"); print(f"✅ 已啟動：{name}")
        elif action == "stop": ps(f"Stop-VM -Name '{name}' -Force"); print(f"✅ 已停止：{name}")
        elif action == "pause": ps(f"Suspend-VM -Name '{name}'"); print(f"✅ 已暫停：{name}")
        elif action == "resume": ps(f"Resume-VM -Name '{name}'"); print(f"✅ 已繼續：{name}")
        elif action == "snapshot":
            sn=snapshot or datetime.now().strftime("snap_%Y%m%d_%H%M%S")
            ps(f"Checkpoint-VM -Name '{name}' -SnapshotName '{sn}'"); print(f"✅ 快照：{sn}")
        elif action == "restore": ps(f"Restore-VMSnapshot -VMName '{name}' -Name '{snapshot}' -Confirm:$false"); print(f"✅ 還原：{snapshot}")
        elif action == "status": out,_=ps(f"Get-VM -Name '{name}' | Format-List"); print(out)
    except Exception as e: print(f"❌ Hyper-V失敗：{e}")

def file_diff(file1: str, file2: str, output: str="", mode: str="unified"):
    import difflib
    try:
        t1=Path(file1).read_text(encoding="utf-8",errors="replace").splitlines(keepends=True)
        t2=Path(file2).read_text(encoding="utf-8",errors="replace").splitlines(keepends=True)
        diff = list(difflib.unified_diff(t1,t2,fromfile=file1,tofile=file2) if mode=="unified" else difflib.context_diff(t1,t2,fromfile=file1,tofile=file2))
        if not diff: print("✅ 兩個檔案完全相同"); return
        result = "".join(diff)
        if output: Path(output).write_text(result,encoding="utf-8"); print(f"✅ diff已存：{output}")
        else: print(result[:3000])
    except Exception as e: print(f"❌ diff失敗：{e}")

def screen_live(fps: float=0.5, duration: float=60.0, quality: int=50):
    import pyautogui, io as _io, time
    interval = 1.0/max(fps,0.1); end=time.time()+duration; count=0
    out_dir=Path.home()/"Desktop"/"screen_live"; out_dir.mkdir(exist_ok=True)
    print(f"📹 開始串流（{fps}FPS，{duration}s）...")
    while time.time()<end:
        sc=pyautogui.screenshot()
        out=str(out_dir/f"frame_{count:04d}.jpg")
        sc.save(out,format="JPEG",quality=quality); count+=1
        time.sleep(interval)
    print(f"✅ 串流完成，{count}張截圖存於：{out_dir}")


# ── Wave 12：AI視覺循環/監控告警/間隔排程/等待文字/瀏覽器進階/語音命令/通知攔截/資料處理/WOL/剪貼簿歷史 ──

def vision_loop(goal: str, max_steps: int = 20, interval: float = 3.0, timeout: float = 120.0):
    import pyautogui, anthropic, base64, io as _io, time, json, re, os
    _client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    steps = 0; start = time.time(); log = []
    while steps < max_steps and (time.time() - start) < timeout:
        screenshot = pyautogui.screenshot()
        buf = _io.BytesIO(); screenshot.save(buf, format="PNG")
        img_b64 = base64.standard_b64encode(buf.getvalue()).decode()
        resp = _client.messages.create(model="claude-sonnet-4-6", max_tokens=512, messages=[{"role":"user","content":[
            {"type":"image","source":{"type":"base64","media_type":"image/png","data":img_b64}},
            {"type":"text","text":f"目標：{goal}\n已執行：{log}\n回答 JSON：{{\"done\":bool,\"action\":\"說明\",\"type\":\"click/type/key/wait\",\"x\":0,\"y\":0,\"text\":\"\"}}"}
        ]}])
        m = re.search(r'\{.*\}', resp.content[0].text, re.DOTALL)
        if not m: steps += 1; time.sleep(interval); continue
        act = json.loads(m.group())
        if act.get("done"): print(f"✅ 目標達成！{steps} 步\n" + "\n".join(log)); return
        if act.get("type") == "click" and act.get("x"): pyautogui.click(act["x"], act["y"])
        elif act.get("type") == "type" and act.get("text"): pyautogui.typewrite(act["text"], interval=0.05)
        elif act.get("type") == "key" and act.get("text"): pyautogui.press(act["text"])
        elif act.get("type") == "wait": time.sleep(2)
        log.append(f"步驟{steps+1}: {act.get('action','')}")
        steps += 1; time.sleep(interval)
    print(f"⏳ 完成 {steps} 步\n" + "\n".join(log))

def alert_monitor(condition: str, threshold: str, target: str = "", interval: int = 30, duration: int = 3600):
    import psutil, time
    end = time.time() + duration
    print(f"📊 監控啟動：{condition} {threshold}，每 {interval}s 檢查")
    while time.time() < end:
        try:
            triggered = False; msg = ""
            if condition == "cpu_above":
                v = psutil.cpu_percent(1)
                if v > float(threshold): triggered = True; msg = f"CPU {v:.1f}% > {threshold}%"
            elif condition == "mem_above":
                v = psutil.virtual_memory().percent
                if v > float(threshold): triggered = True; msg = f"記憶體 {v:.1f}% > {threshold}%"
            elif condition == "disk_above":
                v = psutil.disk_usage("/").percent
                if v > float(threshold): triggered = True; msg = f"磁碟 {v:.1f}% > {threshold}%"
            elif condition == "process_missing":
                pnames = [p.name().lower() for p in psutil.process_iter(["name"])]
                if not any(target.lower() in n for n in pnames): triggered = True; msg = f"程序 {target} 已停止"
            elif condition == "screen_text_found":
                import pyautogui, easyocr, tempfile
                reader = easyocr.Reader(["ch_tra","en"], gpu=False)
                screenshot = pyautogui.screenshot()
                tmp = tempfile.mktemp(suffix=".png"); screenshot.save(tmp)
                results = reader.readtext(tmp, detail=0); Path(tmp).unlink(missing_ok=True)
                if target.lower() in " ".join(results).lower(): triggered = True; msg = f"螢幕出現文字：{target}"
            if triggered: print(f"🔔 告警觸發：{msg}")
        except Exception as e: print(f"❌ {e}")
        time.sleep(interval)
    print("✅ 監控結束")

def interval_schedule(command: str, every_minutes: float = 60.0, repeat: int = 0, duration_hours: float = 0.0):
    import subprocess, time
    end = time.time() + duration_hours * 3600 if duration_hours > 0 else float("inf")
    count = 0; max_count = repeat if repeat > 0 else float("inf")
    print(f"⏱️ 間隔排程啟動：{command}，每 {every_minutes} 分鐘")
    while time.time() < end and count < max_count:
        subprocess.Popen(command, shell=True)
        count += 1; print(f"✅ 第 {count} 次執行")
        time.sleep(every_minutes * 60)
    print(f"✅ 排程結束，共執行 {count} 次")

def wait_for_text(text: str, timeout: float = 60.0, interval: float = 2.0, region: str = ""):
    import pyautogui, easyocr, time, tempfile
    reader = easyocr.Reader(["ch_tra","en"], gpu=False)
    start = time.time()
    reg = None
    if region:
        parts = [int(v) for v in region.split(",")]
        if len(parts) == 4: reg = tuple(parts)
    while time.time() - start < timeout:
        screenshot = pyautogui.screenshot(region=reg)
        tmp = tempfile.mktemp(suffix=".png"); screenshot.save(tmp)
        results = reader.readtext(tmp, detail=0); Path(tmp).unlink(missing_ok=True)
        if text.lower() in " ".join(results).lower():
            print(f"✅ 偵測到文字「{text}」（{time.time()-start:.1f}s）"); return
        time.sleep(interval)
    print(f"⏳ 超時，未偵測到「{text}」")

def data_process(action: str, path: str = "", output: str = "", query: str = "", data: str = "", paths: str = ""):
    import json, csv
    try:
        if action == "read_json":
            content = json.loads(Path(path).read_text(encoding="utf-8"))
            if isinstance(content, list): print(f"JSON {len(content)} 筆：\n{json.dumps(content[:5], ensure_ascii=False, indent=2)}")
            else: print(json.dumps(content, ensure_ascii=False, indent=2)[:2000])
        elif action == "write_json":
            Path(output or path).write_text(json.dumps(json.loads(data), ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"✅ JSON 已儲存：{output or path}")
        elif action == "read_csv":
            with open(path, encoding="utf-8-sig", errors="replace") as f:
                rows = list(csv.DictReader(f))
            print(f"CSV {len(rows)} 筆：\n{json.dumps(rows[:5], ensure_ascii=False, indent=2)}")
        elif action == "write_csv":
            obj = json.loads(data)
            with open(output or path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.DictWriter(f, fieldnames=obj[0].keys()); w.writeheader(); w.writerows(obj)
            print(f"✅ CSV 已儲存：{output or path}")
        elif action == "filter":
            with open(path, encoding="utf-8-sig", errors="replace") as f:
                rows = list(csv.DictReader(f))
            filtered = [r for r in rows if eval(query, {"__builtins__":{}}, r)]
            print(f"過濾 {len(filtered)}/{len(rows)} 筆：\n{json.dumps(filtered[:10], ensure_ascii=False, indent=2)}")
        elif action == "stats":
            with open(path, encoding="utf-8-sig", errors="replace") as f:
                rows = list(csv.DictReader(f))
            for col in rows[0].keys():
                vals = [r[col] for r in rows if r[col]]
                try:
                    nums = [float(v) for v in vals]
                    print(f"{col}: min={min(nums)} max={max(nums)} avg={sum(nums)/len(nums):.2f}")
                except: print(f"{col}: {len(vals)} 筆，{len(set(vals))} 種")
        elif action == "convert":
            ext_in = Path(path).suffix.lower(); ext_out = Path(output).suffix.lower()
            if ext_in == ".json" and ext_out == ".csv":
                obj = json.loads(Path(path).read_text(encoding="utf-8"))
                if not isinstance(obj, list): obj = [obj]
                with open(output, "w", newline="", encoding="utf-8-sig") as f:
                    w = csv.DictWriter(f, fieldnames=obj[0].keys()); w.writeheader(); w.writerows(obj)
            elif ext_in == ".csv" and ext_out == ".json":
                with open(path, encoding="utf-8-sig", errors="replace") as f:
                    rows = list(csv.DictReader(f))
                Path(output).write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"✅ 已轉換：{output}")
        elif action == "merge":
            all_rows = []
            for p in paths.split(","):
                p = p.strip()
                if p.endswith(".csv"):
                    with open(p, encoding="utf-8-sig", errors="replace") as f: all_rows.extend(list(csv.DictReader(f)))
                elif p.endswith(".json"):
                    obj = json.loads(Path(p).read_text(encoding="utf-8"))
                    if isinstance(obj, list): all_rows.extend(obj)
            out = output or str(Path(paths.split(",")[0].strip()).parent / "merged.csv")
            with open(out, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.DictWriter(f, fieldnames=all_rows[0].keys()); w.writeheader(); w.writerows(all_rows)
            print(f"✅ 合併 {len(all_rows)} 筆 → {out}")
        elif action == "to_table":
            with open(path, encoding="utf-8-sig", errors="replace") as f:
                rows = list(csv.DictReader(f))
            cols = list(rows[0].keys())
            print(" | ".join(cols))
            print("-" * 60)
            for r in rows[:20]: print(" | ".join(str(r.get(c,"")) for c in cols))
    except Exception as e:
        print(f"❌ 資料處理失敗：{e}")

def wake_on_lan(mac: str, broadcast: str = "255.255.255.255", port: int = 9):
    import socket
    try:
        mac_clean = mac.replace(":","").replace("-","")
        magic = b"\xff" * 6 + bytes.fromhex(mac_clean) * 16
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
            s.sendto(magic, (broadcast, port))
        print(f"✅ WOL 封包已送出 → {mac}")
    except Exception as e:
        print(f"❌ WOL 失敗：{e}")

_cb_history = []

def clipboard_history(action: str = "list", index: int = 0):
    import pyperclip, threading, time
    global _cb_history
    if action == "start_watch":
        def _w():
            last = ""
            while True:
                try:
                    cur = pyperclip.paste()
                    if cur != last and cur:
                        _cb_history.insert(0, cur)
                        if len(_cb_history) > 50: _cb_history.pop()
                        last = cur
                except: pass
                time.sleep(1)
        threading.Thread(target=_w, daemon=True).start()
        print("✅ 剪貼簿歷史監控已啟動")
    elif action == "list":
        if not _cb_history: print("⚠️ 歷史為空（先執行 start_watch）"); return
        for i, item in enumerate(_cb_history[:20]): print(f"[{i}] {item[:80]}")
    elif action == "get":
        if index < len(_cb_history): print(_cb_history[index])
        else: print(f"⚠️ 索引 {index} 超出範圍")
    elif action == "set":
        if index < len(_cb_history): pyperclip.copy(_cb_history[index]); print(f"✅ 已復原 [{index}]")
        else: print(f"⚠️ 索引 {index} 超出範圍")
    elif action == "clear":
        _cb_history.clear(); print("✅ 已清除")


# ── Wave 11：防火牆/程序/電源/事件/時間/UI自動化/巨集/顏色/攝影機/多螢幕/印表機/WiFi/代理/鎖定/Defender ──

def firewall(action: str, name: str = "", port: int = None, protocol: str = "TCP", direction: str = "Inbound"):
    import subprocess
    try:
        if action == "status":
            r = subprocess.run(["powershell","-Command","Get-NetFirewallProfile | Select-Object Name,Enabled | Format-Table -AutoSize"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout.strip())
        elif action == "list":
            r = subprocess.run(["powershell","-Command","Get-NetFirewallRule | Where-Object {$_.Enabled -eq 'True'} | Select-Object DisplayName,Direction,Action | Format-Table -AutoSize"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout.strip()[:3000])
        elif action == "add":
            r = subprocess.run(["powershell","-Command",f"New-NetFirewallRule -DisplayName '{name}' -Direction {direction} -Protocol {protocol} -LocalPort {port} -Action Allow -Enabled True"], capture_output=True, text=True)
            print(f"✅ 已新增防火牆規則：{name}" if r.returncode==0 else f"❌ 失敗：{r.stderr.strip()}")
        elif action == "remove":
            r = subprocess.run(["powershell","-Command",f"Remove-NetFirewallRule -DisplayName '{name}' -Confirm:$false"], capture_output=True, text=True)
            print(f"✅ 已移除：{name}" if r.returncode==0 else f"❌ 失敗：{r.stderr.strip()}")
        elif action == "enable":
            subprocess.run(["powershell","-Command","Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled True"], capture_output=True)
            print("✅ 防火牆已啟用")
        elif action == "disable":
            subprocess.run(["powershell","-Command","Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False"], capture_output=True)
            print("✅ 防火牆已停用")
    except Exception as e:
        print(f"❌ 防火牆操作失敗：{e}")

def process_mgr(action: str, name: str = "", pid: int = None, level: str = "normal"):
    import psutil
    priority_map = {"realtime":psutil.REALTIME_PRIORITY_CLASS,"high":psutil.HIGH_PRIORITY_CLASS,"above_normal":psutil.ABOVE_NORMAL_PRIORITY_CLASS,"normal":psutil.NORMAL_PRIORITY_CLASS,"below_normal":psutil.BELOW_NORMAL_PRIORITY_CLASS,"idle":psutil.IDLE_PRIORITY_CLASS}
    try:
        if action == "list":
            procs = sorted(psutil.process_iter(["pid","name","cpu_percent","memory_info"]), key=lambda p: p.info["cpu_percent"] or 0, reverse=True)
            print(f"{'PID':>6} {'CPU%':>6} {'MEM MB':>8}  名稱")
            for p in procs[:25]:
                mem = (p.info["memory_info"].rss//1024//1024) if p.info["memory_info"] else 0
                print(f"{p.info['pid']:>6} {p.info['cpu_percent'] or 0:>6.1f} {mem:>8}  {p.info['name']}")
        elif action == "search":
            found = [p for p in psutil.process_iter(["pid","name","cpu_percent","memory_info"]) if name.lower() in p.info["name"].lower()]
            if not found: print(f"⚠️ 找不到：{name}"); return
            for p in found:
                print(f"PID:{p.info['pid']} CPU:{p.info['cpu_percent']}% MEM:{p.info['memory_info'].rss//1024//1024}MB {p.info['name']}")
        elif action == "kill":
            targets = [psutil.Process(int(pid))] if pid else [p for p in psutil.process_iter(["pid","name"]) if name.lower() in p.info["name"].lower()]
            if not targets: print(f"⚠️ 找不到：{name}"); return
            for p in targets: p.kill()
            print(f"✅ 已終止 {len(targets)} 個程序：{name or pid}")
        elif action == "priority":
            p = psutil.Process(int(pid)) if pid else next((x for x in psutil.process_iter(["pid","name"]) if name.lower() in x.info["name"].lower()), None)
            if not p: print(f"⚠️ 找不到：{name}"); return
            p.nice(priority_map.get(level, psutil.NORMAL_PRIORITY_CLASS))
            print(f"✅ 已設定 PID {p.pid} 優先權為 {level}")
    except Exception as e:
        print(f"❌ 程序管理失敗：{e}")

def power_plan(action: str, plan: str = "balanced"):
    import subprocess
    plan_guids = {"balanced":"381b4222-f694-41f0-9685-ff5bb260df2e","high_performance":"8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c","power_saver":"a1841308-3541-4fab-bc81-f71556f20b4a"}
    try:
        if action == "list":
            r = subprocess.run(["powercfg","/list"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout.strip())
        elif action == "get":
            r = subprocess.run(["powercfg","/getactivescheme"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout.strip())
        elif action == "set":
            guid = plan_guids.get(plan)
            if not guid: print(f"⚠️ 未知計畫：{plan}"); return
            subprocess.run(["powercfg","/setactive",guid], capture_output=True)
            print(f"✅ 電源計畫：{plan}")
    except Exception as e:
        print(f"❌ 電源計畫失敗：{e}")

def event_log(log: str = "System", level: str = "Error", count: int = 10):
    import win32evtlog, win32evtlogutil
    level_map = {"Error":1,"Warning":2,"Information":4,"All":7}
    event_type = level_map.get(level, 1)
    try:
        hand = win32evtlog.OpenEventLog(None, log)
        flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
        events = []
        while len(events) < count:
            batch = win32evtlog.ReadEventLog(hand, flags, 0)
            if not batch: break
            for e in batch:
                if level == "All" or e.EventType == event_type:
                    try: msg = win32evtlogutil.SafeFormatMessage(e, log)[:100]
                    except: msg = "(無法讀取)"
                    events.append(f"[{e.TimeGenerated.Format()}] {e.SourceName}: {msg}")
                if len(events) >= count: break
        win32evtlog.CloseEventLog(hand)
        print(f"📋 {log} 事件（{level}）：")
        print("\n".join(events) if events else "✅ 無事件")
    except Exception as e:
        print(f"❌ 事件記錄失敗：{e}")

def datetime_config(action: str, timezone: str = "", datetime_str: str = ""):
    import subprocess
    try:
        if action == "get":
            r = subprocess.run(["powershell","-Command","Get-Date | Format-List; (Get-TimeZone).DisplayName"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout.strip())
        elif action == "sync":
            subprocess.run(["powershell","-Command","Start-Service w32tm -ErrorAction SilentlyContinue; w32tm /resync /force"], capture_output=True)
            print("✅ 已同步網路時間")
        elif action == "set_timezone":
            r = subprocess.run(["powershell","-Command",f"Set-TimeZone -Id '{timezone}'"], capture_output=True, text=True)
            print(f"✅ 時區：{timezone}" if r.returncode==0 else f"❌ 失敗：{r.stderr.strip()}")
        elif action == "set_time":
            r = subprocess.run(["powershell","-Command",f"Set-Date -Date '{datetime_str}'"], capture_output=True, text=True)
            print(f"✅ 時間：{datetime_str}" if r.returncode==0 else f"❌ 失敗：{r.stderr.strip()}")
    except Exception as e:
        print(f"❌ 時間設定失敗：{e}")

def ui_auto(action: str, window: str = "", control: str = "", text: str = ""):
    from pywinauto import Desktop
    try:
        desktop = Desktop(backend="uia")
        if action == "get_windows":
            wins = [w.window_text() for w in desktop.windows() if w.window_text()]
            print("\n".join(f"- {w}" for w in wins[:30]))
            return
        win = next((w for w in desktop.windows() if window.lower() in w.window_text().lower()), None)
        if not win: print(f"⚠️ 找不到視窗：{window}"); return
        if action == "find":
            info = [f"[{c.control_type()}] {c.window_text()}" for c in win.descendants() if c.window_text()][:30]
            print("\n".join(info))
        elif action == "read":
            print("\n".join(c.window_text() for c in win.descendants() if c.window_text())[:50])
        elif action == "click":
            c = next((c for c in win.descendants() if control.lower() in c.window_text().lower()), None)
            if c: c.click_input(); print(f"✅ 已點擊：{c.window_text()}")
            else: print(f"⚠️ 找不到：{control}")
        elif action == "type":
            c = next((c for c in win.descendants() if control.lower() in c.window_text().lower() or c.control_type() in ("Edit","Document")), None)
            if c: c.type_keys(text); print(f"✅ 已輸入文字")
            else: print(f"⚠️ 找不到輸入框：{control}")
    except Exception as e:
        print(f"❌ UI 自動化失敗：{e}")

def macro(action: str, name: str = "", repeat: int = 1, duration: float = 10.0):
    import keyboard, json, time
    macro_file = Path.home() / ".claude_macros.json"
    store = json.loads(macro_file.read_text()) if macro_file.exists() else {}
    try:
        if action == "record_start":
            keyboard.start_recording()
            time.sleep(duration)
            recorded = keyboard.stop_recording()
            store[name] = [{"type":e.event_type,"name":e.name,"time":e.time} for e in recorded]
            macro_file.write_text(json.dumps(store))
            print(f"✅ 已錄製巨集 '{name}'（{len(recorded)} 事件）")
        elif action == "play":
            if name not in store: print(f"⚠️ 找不到巨集：{name}"); return
            events = store[name]
            for _ in range(repeat):
                prev = events[0]["time"] if events else 0
                for e in events:
                    time.sleep(min(max(0,e["time"]-prev),0.5)); prev=e["time"]
                    keyboard.press(e["name"]) if e["type"]=="down" else keyboard.release(e["name"])
            print(f"✅ 已回放巨集 '{name}' {repeat} 次")
        elif action == "list":
            print("\n".join(f"- {k}（{len(v)} 事件）" for k,v in store.items()) if store else "⚠️ 無已儲存巨集")
        elif action == "delete":
            if name in store:
                del store[name]; macro_file.write_text(json.dumps(store)); print(f"✅ 已刪除：{name}")
            else: print(f"⚠️ 找不到：{name}")
    except Exception as e:
        print(f"❌ 巨集操作失敗：{e}")

def color_pick(x: int, y: int, action: str = "pick", w: int = 100, h: int = 100):
    import pyautogui
    from collections import Counter
    try:
        screenshot = pyautogui.screenshot()
        if action == "pick":
            r,g,b = screenshot.getpixel((x,y))[:3]
            print(f"🎨 ({x},{y}) RGB:({r},{g},{b}) HEX:#{r:02X}{g:02X}{b:02X}")
        elif action == "dominant":
            region = screenshot.crop((x,y,x+w,y+h)).convert("RGB").resize((50,50))
            top5 = Counter(list(region.getdata())).most_common(5)
            for (r,g,b),cnt in top5:
                print(f"RGB({r},{g},{b}) #{r:02X}{g:02X}{b:02X}  出現{cnt}次")
    except Exception as e:
        print(f"❌ 顏色選取失敗：{e}")

def webcam(action: str, device: int = 0, duration: float = 5.0, output: str = ""):
    import cv2
    try:
        if action == "list":
            found = []
            for i in range(5):
                cap = cv2.VideoCapture(i)
                if cap.isOpened(): found.append(f"裝置 {i}"); cap.release()
            print("\n".join(found) if found else "⚠️ 無可用攝影機")
        elif action == "photo":
            cap = cv2.VideoCapture(device)
            if not cap.isOpened(): print(f"❌ 無法開啟攝影機 {device}"); return
            ret, frame = cap.read(); cap.release()
            if not ret: print("❌ 無法拍攝"); return
            out = output or str(Path.home()/"Desktop"/f"webcam_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg")
            cv2.imwrite(out, frame); print(f"✅ 已拍照：{out}")
        elif action == "video":
            cap = cv2.VideoCapture(device)
            if not cap.isOpened(): print(f"❌ 無法開啟攝影機 {device}"); return
            out = output or str(Path.home()/"Desktop"/f"webcam_{datetime.now().strftime('%Y%m%d_%H%M%S')}.avi")
            fw,fh = int(cap.get(3)),int(cap.get(4))
            writer = cv2.VideoWriter(out,cv2.VideoWriter_fourcc(*"XVID"),20,(fw,fh))
            import time; end=time.time()+duration
            while time.time()<end:
                ret,frame=cap.read()
                if ret: writer.write(frame)
            cap.release(); writer.release(); print(f"✅ 已錄影：{out}")
    except Exception as e:
        print(f"❌ 攝影機操作失敗：{e}")

def multi_monitor(action: str, monitor: int = 1, window: str = ""):
    import subprocess, win32gui
    try:
        if action == "list":
            r = subprocess.run(["powershell","-Command","Get-CimInstance Win32_DesktopMonitor | Select-Object Name,ScreenWidth,ScreenHeight | Format-Table -AutoSize"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout.strip())
        elif action in ("extend","clone"):
            subprocess.Popen(["displayswitch.exe", "/extend" if action=="extend" else "/clone"])
            print(f"✅ 螢幕模式：{action}")
        elif action == "move_window":
            import ctypes; sw=ctypes.windll.user32.GetSystemMetrics(0)
            hwnds=[]
            win32gui.EnumWindows(lambda h,l: l.append(h) if win32gui.IsWindowVisible(h) and window.lower() in win32gui.GetWindowText(h).lower() else None, hwnds)
            if not hwnds: print(f"⚠️ 找不到視窗：{window}"); return
            rect=win32gui.GetWindowRect(hwnds[0])
            win32gui.MoveWindow(hwnds[0],sw*(monitor-1)+100,100,rect[2]-rect[0],rect[3]-rect[1],True)
            print(f"✅ 視窗 '{window}' 移至螢幕 {monitor}")
    except Exception as e:
        print(f"❌ 多螢幕管理失敗：{e}")

def printer(action: str, path: str = "", printer_name: str = ""):
    import win32print
    try:
        if action == "list":
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL|win32print.PRINTER_ENUM_CONNECTIONS)
            default = win32print.GetDefaultPrinter()
            print(f"預設：{default}")
            for p in printers: print(f"- {p[2]}")
        elif action == "print":
            import win32api
            pname = printer_name or win32print.GetDefaultPrinter()
            win32api.ShellExecute(0,"print",path,f'/d:"{pname}"',".",0)
            print(f"✅ 已傳送列印：{path}")
        elif action == "queue":
            pname = printer_name or win32print.GetDefaultPrinter()
            h = win32print.OpenPrinter(pname)
            jobs = win32print.EnumJobs(h,0,-1,1)
            win32print.ClosePrinter(h)
            print("佇列為空" if not jobs else "\n".join(f"工作{j['JobId']}：{j['pDocument']}" for j in jobs))
        elif action == "clear_queue":
            pname = printer_name or win32print.GetDefaultPrinter()
            h = win32print.OpenPrinter(pname)
            for j in win32print.EnumJobs(h,0,-1,1):
                win32print.SetJob(h,j["JobId"],0,None,win32print.JOB_CONTROL_DELETE)
            win32print.ClosePrinter(h); print(f"✅ 佇列已清除")
        elif action == "set_default":
            win32print.SetDefaultPrinter(printer_name); print(f"✅ 預設印表機：{printer_name}")
    except Exception as e:
        print(f"❌ 印表機操作失敗：{e}")

def wifi(action: str, ssid: str = "", password: str = ""):
    import subprocess
    try:
        cmd_map = {"scan":["netsh","wlan","show","networks","mode=Bssid"],"status":["netsh","wlan","show","interfaces"],"saved":["netsh","wlan","show","profiles"],"disconnect":["netsh","wlan","disconnect"]}
        if action in cmd_map:
            r = subprocess.run(cmd_map[action], capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout.strip()[:2000])
        elif action == "password":
            r = subprocess.run(["netsh","wlan","show","profile",f"name={ssid}","key=clear"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout.strip())
        elif action == "connect":
            r = subprocess.run(["netsh","wlan","connect",f"name={ssid}"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout.strip())
    except Exception as e:
        print(f"❌ WiFi 操作失敗：{e}")

def proxy(action: str, host: str = ""):
    import winreg
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Internet Settings"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ|winreg.KEY_SET_VALUE) as k:
            if action == "get":
                try:
                    enabled = winreg.QueryValueEx(k,"ProxyEnable")[0]
                    server = winreg.QueryValueEx(k,"ProxyServer")[0]
                    print(f"代理：{'啟用' if enabled else '停用'}  伺服器：{server}")
                except: print("代理：未設定")
            elif action == "set":
                winreg.SetValueEx(k,"ProxyEnable",0,winreg.REG_DWORD,1)
                winreg.SetValueEx(k,"ProxyServer",0,winreg.REG_SZ,host)
                print(f"✅ 代理已設定：{host}")
            elif action == "disable":
                winreg.SetValueEx(k,"ProxyEnable",0,winreg.REG_DWORD,0)
                print("✅ 代理已停用")
    except Exception as e:
        print(f"❌ 代理設定失敗：{e}")

def lock_screen(action: str = "lock"):
    import subprocess
    try:
        if action == "lock":
            subprocess.Popen(["rundll32.exe","user32.dll,LockWorkStation"]); print("🔒 螢幕已鎖定")
        elif action == "logoff":
            subprocess.run(["shutdown","/l"],capture_output=True); print("✅ 已登出")
        elif action == "switch_user":
            subprocess.Popen(["tsdiscon.exe"]); print("✅ 已切換使用者")
    except Exception as e:
        print(f"❌ 鎖定/登出失敗：{e}")

def defender(action: str, path: str = ""):
    import subprocess
    try:
        if action == "status":
            r = subprocess.run(["powershell","-Command","Get-MpComputerStatus | Select-Object AMRunningMode,RealTimeProtectionEnabled,AntivirusSignatureLastUpdated | Format-List"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout.strip())
        elif action in ("quick_scan","full_scan"):
            scan_type = "QuickScan" if action=="quick_scan" else "FullScan"
            subprocess.Popen(["powershell","-Command",f"Start-MpScan -ScanType {scan_type}"])
            print(f"🛡️ {action} 已啟動（背景執行中）")
        elif action == "threats":
            r = subprocess.run(["powershell","-Command","Get-MpThreatDetection | Select-Object ThreatID,Resources,ActionSuccess | Format-Table -AutoSize"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout.strip() or "✅ 無威脅記錄")
        elif action == "add_exclusion":
            r = subprocess.run(["powershell","-Command",f"Add-MpPreference -ExclusionPath '{path}'"], capture_output=True, text=True)
            print(f"✅ 已新增排除：{path}" if r.returncode==0 else f"❌ 失敗：{r.stderr.strip()}")
        elif action == "remove_exclusion":
            r = subprocess.run(["powershell","-Command",f"Remove-MpPreference -ExclusionPath '{path}'"], capture_output=True, text=True)
            print(f"✅ 已移除排除：{path}" if r.returncode==0 else f"❌ 失敗：{r.stderr.strip()}")
        elif action == "list_exclusions":
            r = subprocess.run(["powershell","-Command","(Get-MpPreference).ExclusionPath"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout.strip() or "✅ 無排除項目")
    except Exception as e:
        print(f"❌ Defender 操作失敗：{e}")


# ── 缺口1：觸發驅動 ─────────────────────────────────

def email_trigger(action, host="", user="", password="", filter_from="",
                  filter_subject="", duration=300, to="", subject="", body=""):
    from bot import execute_email_trigger
    print(execute_email_trigger(action, host, user, password, filter_from,
                                filter_subject, duration, to, subject, body))

def file_trigger(folder, event, action, pattern="", target="", duration=60):
    from bot import execute_file_trigger
    print(execute_file_trigger(folder, event, action, pattern, target, duration))

def webhook_server(action, port=8765, secret=""):
    from bot import execute_webhook_server
    print(execute_webhook_server(action, int(port), secret))

# ── 缺口2：應用程式深度控制 ──────────────────────────

def com_auto(app, action, path="", sheet=None, cell="", value="", macro="", to="", subject=""):
    from bot import execute_com_auto
    print(execute_com_auto(app, action, path, sheet, cell, value, macro, to, subject))

def dialog_auto(action, button_text="", window_title="", timeout=30):
    from bot import execute_dialog_auto
    print(execute_dialog_auto(action, button_text, window_title, int(timeout)))

def ime_switch(action):
    from bot import execute_ime_switch
    print(execute_ime_switch(action))

# ── 缺口3：感知能力 ──────────────────────────────────

def wake_word(action, keyword="", duration=5, language="zh-TW"):
    from bot import execute_wake_word
    print(execute_wake_word(action, keyword, float(duration), language))

def sound_detect(action, threshold=20, duration=5, output=""):
    from bot import execute_sound_detect
    print(execute_sound_detect(action, float(threshold), float(duration), output))

def face_recognize(action, name="", image_path="", output=""):
    from bot import execute_face_recognize
    print(execute_face_recognize(action, name, image_path, output))

# ── 缺口4：跨裝置控制 ────────────────────────────────

def http_server(action, port=9876, password=""):
    from bot import execute_http_server
    print(execute_http_server(action, int(port), password))

def lan_scan(action, subnet="", host="", port=80):
    from bot import execute_lan_scan
    print(execute_lan_scan(action, subnet, host, int(port)))

def serial_port(action, port="", baudrate=9600, data="", timeout=2):
    from bot import execute_serial_port
    print(execute_serial_port(action, port, int(baudrate), data, float(timeout)))

def mqtt(action, broker, port=1883, topic="", message="", duration=10, username="", password=""):
    from bot import execute_mqtt
    print(execute_mqtt(action, broker, int(port), topic, message, float(duration), username, password))

# ── 缺口5：內容理解與處理 ────────────────────────────

def doc_ai(action, path="", path2="", fields="", question="", url=""):
    from bot import execute_doc_ai
    print(execute_doc_ai(action, path, path2, fields, question, url))

def web_monitor(action, url, selector="body", interval=60, duration=300, keyword=""):
    from bot import execute_web_monitor
    print(execute_web_monitor(action, url, selector, float(interval), float(duration), keyword))

def audio_transcribe(action, path="", duration=30, language="", output=""):
    from bot import execute_audio_transcribe
    print(execute_audio_transcribe(action, path, float(duration), language, output))

# ══════════════════════════════════════════════════════
# 奧創升級技能集
# ══════════════════════════════════════════════════════

def osint_search(action, query="", target="", limit=10):
    from bot import execute_osint_search
    print(execute_osint_search(action, query, target, int(limit)))

def news_monitor(action, keywords="", interval=300, duration=3600):
    from bot import execute_news_monitor
    print(execute_news_monitor(action, keywords, float(interval), float(duration)))

def threat_intel(action, target="", api_key=""):
    from bot import execute_threat_intel
    print(execute_threat_intel(action, target, api_key))

def auto_skill(action, goal="", skill_name="", code=""):
    from bot import execute_auto_skill
    print(execute_auto_skill(action, goal, skill_name, code))

def smart_home(action, device="", value="", host="", token=""):
    from bot import execute_smart_home
    print(execute_smart_home(action, device, value, host, token))

def goal_manager(action, goal="", goal_id="", steps="", priority="normal"):
    from bot import execute_goal_manager
    print(execute_goal_manager(action, goal, goal_id, steps, priority))

def auto_trade(action, symbol="", amount=0.0, price=0.0, order_type="market"):
    from bot import execute_auto_trade
    print(execute_auto_trade(action, symbol, float(amount), float(price), order_type))

def knowledge_base(action, content="", query="", tag="", kb_id=""):
    from bot import execute_knowledge_base
    print(execute_knowledge_base(action, content, query, tag, kb_id))

def emotion_detect(action, text="", image_path=""):
    from bot import execute_emotion_detect
    print(execute_emotion_detect(action, text, image_path))

def voice_id(action, name="", audio_path="", duration=5):
    from bot import execute_voice_id
    print(execute_voice_id(action, name, audio_path, float(duration)))

def pentest(action, target="", port_range="1-1000", timeout=2):
    from bot import execute_pentest
    print(execute_pentest(action, target, port_range, float(timeout)))

def proactive_alert(action, name="", condition="", threshold="", target="", interval=60):
    from bot import execute_proactive_alert
    print(execute_proactive_alert(action, name, condition, threshold, target, float(interval)))

def multi_deploy(action, remote_host="", remote_user="", remote_pass="", remote_path="/tmp/niu_bot"):
    from bot import execute_multi_deploy
    print(execute_multi_deploy(action, remote_host, remote_user, remote_pass, remote_path))

def self_benchmark(action):
    from bot import execute_self_benchmark
    print(execute_self_benchmark(action))


# ── 新增投資技能 ─────────────────────────────────────

def get_institutional(symbol="", date=""):
    """台股三大法人買賣超。symbol=股票代號（空=整體市場），date=YYYYMMDD（空=今天）"""
    from bot import fetch_institutional
    print(fetch_institutional(symbol, date))

def get_sector(market="us"):
    """產業類股表現。market=us（美股）或 tw（台股）"""
    from bot import fetch_sector
    print(fetch_sector(market))

def get_commodity(items="all"):
    """大宗商品報價。items=gold,oil,silver,copper,natgas,wheat,corn 或 all"""
    from bot import fetch_commodity
    item_list = [x.strip() for x in items.split(",")] if items != "all" else ["all"]
    print(fetch_commodity(item_list))

def get_bond_yield():
    """美國公債殖利率（2Y/5Y/10Y/30Y）及利差分析"""
    from bot import fetch_bond_yield
    print(fetch_bond_yield())

def get_dividend_calendar(symbol):
    """除權息資訊。symbol=股票代號（如 0056.TW、AAPL）"""
    from bot import fetch_dividend_calendar
    print(fetch_dividend_calendar(symbol))

def stock_screener(criteria, market="us"):
    """選股篩選。criteria=篩選條件（如「殖利率>5%」），market=us/tw"""
    from bot import fetch_stock_screener
    print(fetch_stock_screener(criteria, market))

def get_margin_trading(symbol, date=""):
    """台股融資融券餘額。symbol=台股代號，date=YYYYMMDD（空=今天）"""
    from bot import fetch_margin_trading
    print(fetch_margin_trading(symbol, date))

def get_options(symbol, expiry=""):
    """選擇權鏈。symbol=股票代號（如 AAPL），expiry=到期日（空=最近一個）"""
    from bot import fetch_options
    print(fetch_options(symbol, expiry))

def get_futures(items="all"):
    """期貨報價。items=sp500,nasdaq,dow,gold,oil,taiex 或 all"""
    from bot import fetch_futures
    item_list = [x.strip() for x in items.split(",")] if items != "all" else ["all"]
    print(fetch_futures(item_list))

def get_ipo(count=10):
    """近期 IPO 行事曆。count=顯示筆數（預設10）"""
    from bot import fetch_ipo
    print(fetch_ipo(int(count)))

def backtest(symbol, strategy="ma_cross", period="2y"):
    """回測投資策略。strategy=ma_cross/buy_hold/dca，period=1y/2y/3y/5y"""
    from bot import fetch_backtest
    print(fetch_backtest(symbol, strategy, period))

def get_ashare(code, period="1mo"):
    """A股/港股查詢。code=6位A股代號或4位港股代號，period=1mo/3mo/6mo/1y"""
    from bot import fetch_ashare
    print(fetch_ashare(code, period))

def get_cn_news(source="all", count=5):
    """中國大陸新聞。source=xinhua/people/36kr/caixin/all，count=顯示則數"""
    from bot import fetch_cn_news
    print(fetch_cn_news(source, int(count)))

def china_search(query, category="其他", count=6):
    """中國大陸全方位搜尋（旅遊/美食/文化/戲劇/演員/工作等）。"""
    from bot import fetch_china_search
    print(fetch_china_search(query, category, int(count)))

def get_global_market():
    """全球主要股市指數概覽"""
    from bot import fetch_global_market
    print(fetch_global_market())

def get_economic_calendar(count=10):
    """重要經濟數據行事曆（CPI/非農/GDP/Fed）"""
    from bot import fetch_economic_calendar
    print(fetch_economic_calendar(int(count)))

def get_earnings_calendar(days=7):
    """未來N天財報日曆"""
    from bot import fetch_earnings_calendar
    print(fetch_earnings_calendar(int(days)))

def get_analyst_ratings(symbol):
    """分析師評級升降評紀錄。symbol=股票代號"""
    from bot import fetch_analyst_ratings
    print(fetch_analyst_ratings(symbol))

def get_short_interest(symbol):
    """空頭比率/借券賣出資料。symbol=股票代號"""
    from bot import fetch_short_interest
    print(fetch_short_interest(symbol))

def get_correlation(symbols, period="1y"):
    """多股相關性矩陣。symbols=逗號分隔代號，period=3mo/6mo/1y/2y"""
    from bot import fetch_correlation
    sym_list = [s.strip() for s in symbols.split(",")] if isinstance(symbols, str) else symbols
    print(fetch_correlation(sym_list, period))

def get_risk_metrics(symbol, period="1y"):
    """風險指標：Beta/夏普/波動率/VaR。period=1y/2y/3y"""
    from bot import fetch_risk_metrics
    print(fetch_risk_metrics(symbol, period))

def get_money_flow(symbol):
    """個股資金流向分析。symbol=股票代號"""
    from bot import fetch_money_flow
    print(fetch_money_flow(symbol))

def get_concept_stocks(theme):
    """台股概念股查詢。theme=AI/電動車/軍工/低軌衛星/半導體等"""
    from bot import fetch_concept_stocks
    print(fetch_concept_stocks(theme))

def get_crypto_depth(coin="bitcoin"):
    """加密幣深度：鏈上數據/資金費率/DeFi。coin=bitcoin/ethereum/solana等"""
    from bot import fetch_crypto_depth
    print(fetch_crypto_depth(coin))

def drip_calculator(symbol, shares, years=10, monthly_invest=0):
    """DRIP股息再投資試算。shares=初始股數，years=持有年數，monthly_invest=每月追加"""
    from bot import fetch_drip_calculator
    print(fetch_drip_calculator(symbol, float(shares), int(years), float(monthly_invest)))

def get_forex_chart(pair, period="3mo"):
    """外匯技術分析。pair=USDTWD=X等，period=1mo/3mo/6mo/1y"""
    from bot import fetch_forex_chart
    print(fetch_forex_chart(pair, period))

def get_warrant(underlying):
    """台股認購/認售權證。underlying=標的股代號（如2330）"""
    from bot import fetch_warrant
    print(fetch_warrant(underlying))

def get_portfolio_risk(symbols_weights, period="1y"):
    """投資組合風險分析。symbols_weights=代號:權重,代號:權重（如AAPL:0.5,MSFT:0.5）"""
    from bot import fetch_portfolio_risk
    if isinstance(symbols_weights, str):
        holdings = []
        for item in symbols_weights.split(","):
            parts = item.strip().split(":")
            if len(parts) == 2:
                holdings.append({"symbol": parts[0].strip(), "weight": float(parts[1].strip())})
            else:
                holdings.append({"symbol": parts[0].strip(), "weight": 1.0})
    else:
        holdings = symbols_weights
    print(fetch_portfolio_risk(holdings, period))


def retirement_calculator(current_age, current_savings, monthly_save,
                           retire_age=65, annual_return=6.0, monthly_expense=50000):
    """退休規劃試算。用法：retirement_calculator <年齡> <現有資產萬元> <月儲蓄元> [退休年齡] [年報酬%] [月支出元]"""
    from bot import fetch_retirement_calculator
    print(fetch_retirement_calculator(int(current_age), float(current_savings), float(monthly_save),
                                      int(retire_age), float(annual_return), float(monthly_expense)))


def loan_calculator(principal, annual_rate, years, loan_type="等額本息"):
    """貸款試算。用法：loan_calculator <貸款萬元> <年利率%> <年數> [等額本息|等額本金]"""
    from bot import fetch_loan_calculator
    print(fetch_loan_calculator(float(principal), float(annual_rate), int(years), loan_type))


def compound_calculator(principal, annual_rate, years, monthly_add=0, compound_freq=12):
    """複利計算器。用法：compound_calculator <本金> <年報酬%> <年數> [每月追加] [複利頻率]"""
    from bot import fetch_compound_calculator
    print(fetch_compound_calculator(float(principal), float(annual_rate), int(years),
                                    float(monthly_add), int(compound_freq)))


def asset_allocation(age, risk_level="穩健", goal="退休", investment_horizon=None):
    """資產配置建議。用法：asset_allocation <年齡> [保守|穩健|積極] [目標] [投資年數]"""
    from bot import fetch_asset_allocation
    print(fetch_asset_allocation(int(age), risk_level, goal,
                                 int(investment_horizon) if investment_horizon else None))


def tw_tax_calculator(dividend_income, other_income=0, tax_bracket=None, sell_amount=0):
    """台股稅務試算。用法：tw_tax_calculator <股利所得元> [其他收入元] [稅率%] [賣出金額元]"""
    from bot import fetch_tw_tax_calculator
    print(fetch_tw_tax_calculator(float(dividend_income), float(other_income),
                                  float(tax_bracket) if tax_bracket else None, float(sell_amount)))


def currency_converter(amount, from_currency, to_currency):
    """外幣換算。用法：currency_converter <金額> <來源幣別> <目標幣別>"""
    from bot import fetch_currency_converter
    print(fetch_currency_converter(float(amount), from_currency, to_currency))


def get_fund(symbol):
    """基金查詢。用法：get_fund <基金代號>"""
    from bot import fetch_fund
    print(fetch_fund(symbol))


def get_reits(symbol):
    """REITs查詢。用法：get_reits <REITs代號>"""
    from bot import fetch_reits
    print(fetch_reits(symbol))


def inflation_adjusted(nominal_return, years, amount, inflation_rate=2.0):
    """通膨調整報酬。用法：inflation_adjusted <名目報酬%> <年數> <本金元> [通膨率%]"""
    from bot import fetch_inflation_adjusted
    print(fetch_inflation_adjusted(float(nominal_return), int(years), float(amount), float(inflation_rate)))


def defi_calculator(principal_usd, apy, days, compound=True, protocol=""):
    """DeFi收益試算。用法：defi_calculator <本金USD> <APY%> <天數> [複利true/false] [協議名]"""
    from bot import fetch_defi_calculator
    c = compound if isinstance(compound, bool) else str(compound).lower() != "false"
    print(fetch_defi_calculator(float(principal_usd), float(apy), int(days), c, str(protocol)))


def gold_calculator(weight, unit="公克", currency="TWD"):
    """黃金換算。用法：gold_calculator <重量> [公克|錢|兩|盎司] [TWD|USD]"""
    from bot import fetch_gold_calculator
    print(fetch_gold_calculator(float(weight), unit, currency))


def forex_deposit(amount_twd, currency, annual_rate, months, buy_rate=None, sell_rate=None):
    """外幣定存試算。用法：forex_deposit <台幣本金> <幣別> <年利率%> <月數> [買入匯率] [賣出匯率]"""
    from bot import fetch_forex_deposit
    print(fetch_forex_deposit(float(amount_twd), currency, float(annual_rate), int(months),
                               float(buy_rate) if buy_rate else None,
                               float(sell_rate) if sell_rate else None))


def financial_health(monthly_income, monthly_expense, total_assets, total_debt,
                     emergency_fund_months=0, has_insurance=False, investment_ratio=0):
    """財務健康診斷。用法：financial_health <月收入> <月支出> <總資產> <總負債> [備用金月數] [有保險y/n] [投資比例%]"""
    from bot import fetch_financial_health
    ins = has_insurance if isinstance(has_insurance, bool) else str(has_insurance).lower() in ("y", "true", "yes", "1")
    print(fetch_financial_health(float(monthly_income), float(monthly_expense),
                                 float(total_assets), float(total_debt),
                                 float(emergency_fund_months), ins, float(investment_ratio)))


def deep_research(topic, lang="zh-tw", depth=5):
    """深度研究。用法：deep_research <主題> [zh-tw|en] [深度3-8]"""
    from bot import fetch_deep_research
    print(fetch_deep_research(topic, lang, int(depth)))


def fact_check(claim, lang="zh-tw"):
    """事實查核。用法：fact_check <要查核的說法>"""
    from bot import fetch_fact_check
    print(fetch_fact_check(claim, lang))


def timeline_events(topic, lang="zh-tw"):
    """時間軸整理。用法：timeline_events <主題>"""
    from bot import fetch_timeline_events
    print(fetch_timeline_events(topic, lang))


def sentiment_scan(topic, lang="zh-tw"):
    """輿情掃描。用法：sentiment_scan <話題>"""
    from bot import fetch_sentiment_scan
    print(fetch_sentiment_scan(topic, lang))


def compare_analysis(items_str, context=""):
    """多項比較。用法：compare_analysis <A,B,C> [背景說明]"""
    from bot import fetch_compare_analysis
    items = [i.strip() for i in items_str.split(",")]
    print(fetch_compare_analysis(items, None, context))


def pros_cons_analysis(subject, context="", lang="zh-tw"):
    """優缺點分析。用法：pros_cons_analysis <主題> [背景] [語言]"""
    from bot import fetch_pros_cons_analysis
    print(fetch_pros_cons_analysis(subject, context, lang))


def research_report(topic, purpose="一般研究", lang="zh-tw"):
    """研究報告。用法：research_report <主題> [目的] [語言]"""
    from bot import fetch_research_report
    print(fetch_research_report(topic, purpose, lang))


def opinion_writer(topic, stance="中立", style="正式"):
    """觀點撰寫。用法：opinion_writer <主題> [支持|反對|中立|批判] [正式|輕鬆|犀利]"""
    from bot import fetch_opinion_writer
    print(fetch_opinion_writer(topic, stance, style))


def trend_forecast(topic, timeframe="全部", lang="zh-tw"):
    """趨勢預測。用法：trend_forecast <主題> [短期|中期|長期|全部]"""
    from bot import fetch_trend_forecast
    print(fetch_trend_forecast(topic, timeframe, lang))


def debate_simulator(motion, lang="zh-tw"):
    """辯論模擬。用法：debate_simulator <辯論題目>"""
    from bot import fetch_debate_simulator
    print(fetch_debate_simulator(motion, lang))


def academic_search(query, field="", lang="en"):
    """學術論文搜尋。用法：academic_search <關鍵字> [領域] [語言]"""
    from bot import fetch_academic_search
    print(fetch_academic_search(query, field, lang))


def health_research(topic, lang="zh-tw"):
    """健康資訊搜尋。用法：health_research <主題>"""
    from bot import fetch_health_research
    print(fetch_health_research(topic, lang))


def law_research(topic, jurisdiction="台灣", lang="zh-tw"):
    """法規查詢。用法：law_research <主題> [地區] [語言]"""
    from bot import fetch_law_research
    print(fetch_law_research(topic, jurisdiction, lang))


def person_research(name, context="", lang="zh-tw"):
    """人物研究。用法：person_research <姓名> [背景說明]"""
    from bot import fetch_person_research
    print(fetch_person_research(name, context, lang))


def company_research(company, lang="zh-tw"):
    """公司深度研究。用法：company_research <公司名稱或代號>"""
    from bot import fetch_company_research
    print(fetch_company_research(company, lang))


def product_review(product, category="", lang="zh-tw"):
    """產品評測彙整。用法：product_review <產品名稱> [類別]"""
    from bot import fetch_product_review
    print(fetch_product_review(product, category, lang))


def travel_research(destination, days=None, style="", lang="zh-tw"):
    """旅遊研究。用法：travel_research <目的地> [天數] [風格]"""
    from bot import fetch_travel_research
    print(fetch_travel_research(destination, int(days) if days else None, style, lang))


def job_market(job_title, location="台灣", lang="zh-tw"):
    """職涯市場分析。用法：job_market <職位名稱> [地區]"""
    from bot import fetch_job_market
    print(fetch_job_market(job_title, location, lang))


def impact_analysis(event, scope_str="", lang="zh-tw"):
    """影響力分析。用法：impact_analysis <事件> [個人,企業,社會,經濟]"""
    from bot import fetch_impact_analysis
    scope = [s.strip() for s in scope_str.split(",")] if scope_str else None
    print(fetch_impact_analysis(event, scope, lang))


def scenario_planning(topic, horizon="", lang="zh-tw"):
    """情境規劃。用法：scenario_planning <主題> [時間範圍]"""
    from bot import fetch_scenario_planning
    print(fetch_scenario_planning(topic, horizon, lang))


def decision_helper(question, options_str="", criteria_str=""):
    """決策輔助。用法：decision_helper <決策問題> [選項A,選項B] [考量1,考量2]"""
    from bot import fetch_decision_helper
    options = [o.strip() for o in options_str.split(",")] if options_str else None
    criteria = [c.strip() for c in criteria_str.split(",")] if criteria_str else None
    print(fetch_decision_helper(question, options, criteria))


def devil_advocate(position, lang="zh-tw"):
    """魔鬼代言人。用法：devil_advocate <要被挑戰的觀點>"""
    from bot import fetch_devil_advocate
    print(fetch_devil_advocate(position, lang))


def summary_writer(topic, max_points=7, lang="zh-tw"):
    """多來源摘要。用法：summary_writer <主題> [重點數] [語言]"""
    from bot import fetch_summary_writer
    print(fetch_summary_writer(topic, int(max_points), lang))


def key_insights(topic, count=5, lang="zh-tw"):
    """洞察萃取。用法：key_insights <主題> [數量] [語言]"""
    from bot import fetch_key_insights
    print(fetch_key_insights(topic, int(count), lang))


def bias_detector(topic, lang="zh-tw"):
    """偏見偵測。用法：bias_detector <議題>"""
    from bot import fetch_bias_detector
    print(fetch_bias_detector(topic, lang))


def second_opinion(question, experts_str="", lang="zh-tw"):
    """多專家視角。用法：second_opinion <問題> [專家A,專家B,...]"""
    from bot import fetch_second_opinion
    experts = [e.strip() for e in experts_str.split(",")] if experts_str else None
    print(fetch_second_opinion(question, experts, lang))


def brainstorm(problem, count=8, style="實用", lang="zh-tw"):
    """腦力激盪。用法：brainstorm <問題> [數量] [實用|創意|顛覆]"""
    from bot import fetch_brainstorm
    print(fetch_brainstorm(problem, int(count), style, lang))


def benchmark_analysis(subject, industry="", lang="zh-tw"):
    """標竿分析。用法：benchmark_analysis <對象> [產業]"""
    from bot import fetch_benchmark_analysis
    print(fetch_benchmark_analysis(subject, industry, lang))


def steel_man(opposing_view, own_position="", lang="zh-tw"):
    """鋼人論證。用法：steel_man <對立觀點> [自己的立場]"""
    from bot import fetch_steel_man
    print(fetch_steel_man(opposing_view, own_position, lang))


def socratic_questioning(topic, depth=5, lang="zh-tw"):
    """蘇格拉底式提問。用法：socratic_questioning <主題> [層數]"""
    from bot import fetch_socratic_questioning
    print(fetch_socratic_questioning(topic, int(depth), lang))


def analogy_maker(concept, audience="一般大眾", count=3, lang="zh-tw"):
    """類比說明。用法：analogy_maker <概念> [受眾] [數量]"""
    from bot import fetch_analogy_maker
    print(fetch_analogy_maker(concept, audience, int(count), lang))


def narrative_builder(topic, key_message="", audience="", lang="zh-tw"):
    """敘事架構。用法：narrative_builder <主題> [核心訊息] [受眾]"""
    from bot import fetch_narrative_builder
    print(fetch_narrative_builder(topic, key_message, audience, lang))


def critique_writer(subject, type_="觀點", lang="zh-tw"):
    """批判性評析。用法：critique_writer <對象> [文章|政策|作品|計劃|觀點]"""
    from bot import fetch_critique_writer
    print(fetch_critique_writer(subject, type_, lang))


def position_statement(issue, stance, lang="zh-tw"):
    """立場聲明。用法：position_statement <議題> <支持|反對|有條件支持>"""
    from bot import fetch_position_statement
    print(fetch_position_statement(issue, stance, lang))


def ocr_click(target_text, monitor=1, click_type="click"):
    """OCR找字點擊。用法：ocr_click <目標文字> [螢幕1/2/3] [click|double_click|right_click]"""
    from bot import fetch_ocr_click
    print(fetch_ocr_click(target_text, int(monitor), click_type))


def vision_locate(description, monitor=1, action="click"):
    """視覺定位點擊。用法：vision_locate <描述> [螢幕] [click|double_click|locate_only]"""
    from bot import fetch_vision_locate
    print(fetch_vision_locate(description, int(monitor), action))


def screen_workflow(steps_json):
    """螢幕工作流。用法：screen_workflow '<JSON步驟陣列>'"""
    import json
    from bot import fetch_screen_workflow
    steps = json.loads(steps_json) if isinstance(steps_json, str) else steps_json
    print(fetch_screen_workflow(steps))


def app_navigator(app, task, input_text="", monitor=1):
    """App導航。用法：app_navigator <App名> <任務描述> [輸入文字] [螢幕]"""
    from bot import fetch_app_navigator
    print(fetch_app_navigator(app, task, input_text, int(monitor)))


def wait_and_click(target_text, timeout=15, monitor=1, action_after="click"):
    """等待出現後點擊。用法：wait_and_click <目標文字> [超時秒數] [螢幕] [click|none]"""
    from bot import fetch_wait_and_click
    print(fetch_wait_and_click(target_text, int(timeout), int(monitor), action_after))


def drag_drop(from_x=None, from_y=None, to_x=None, to_y=None, from_text="", to_text="", monitor=1, duration=0.5):
    """拖曳操作。用法：drag_drop <from_x> <from_y> <to_x> <to_y> 或 drag_drop from_text=<文字> to_text=<文字>"""
    from bot import fetch_drag_drop
    print(fetch_drag_drop(
        int(from_x) if from_x else None, int(from_y) if from_y else None,
        int(to_x) if to_x else None, int(to_y) if to_y else None,
        str(from_text), str(to_text), int(monitor), float(duration)))


# ── 新增 116 個缺失工具（同步 bot.py）──────────────────

def analyze_pdf_tool(path, max_chars=4000):
    """分析 PDF。用法：analyze_pdf <路徑> [最大字數]"""
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            total_pages = len(pdf.pages)
            texts = []
            for page in pdf.pages[:20]:
                t = page.extract_text()
                if t:
                    texts.append(t)
        full_text = "\n".join(texts)
        if not full_text.strip():
            print("PDF 無法提取文字（可能是掃描圖片 PDF）"); return
        result = f"📄 PDF 分析（共 {total_pages} 頁）\n\n{full_text[:int(max_chars)]}"
        if len(full_text) > int(max_chars):
            result += f"\n\n（內容已截斷，共 {len(full_text)} 字）"
        print(result)
    except Exception as e:
        print(f"PDF 讀取失敗：{e}")


def audio_process_tool(action, input_path, output="", start_ms=0, end_ms=0):
    """音訊處理。用法：audio_process <convert|trim|info> <輸入> [輸出] [起始ms] [結束ms]"""
    from bot import execute_audio_process
    print(execute_audio_process(action, input_path, output, int(start_ms), int(end_ms)))


def auto_skill_tool(action, goal="", skill_name="", code=""):
    """自動技能。用法：auto_skill <create|list|run|delete> [目標] [名稱] [程式碼]"""
    from bot import execute_auto_skill
    print(execute_auto_skill(action, goal, skill_name, code))


def auto_trade_tool(action, symbol="", amount=0.0, price=0.0, order_type="market"):
    """自動交易。用法：auto_trade <add|remove|list|simulate> [代號] [數量] [價格] [類型]"""
    from bot import execute_auto_trade
    print(execute_auto_trade(action, symbol, float(amount), float(price), order_type))


def automation_tool(action, condition_type="", condition_value="", command="", duration=60.0, layout="side_by_side", x=0, y=0, w=0, h=0, keyword="", output=""):
    """自動化。用法：automation <if_then|window_arrange|region_ocr|...> [參數...]"""
    from bot import execute_automation
    print(execute_automation(action, condition_type, condition_value, command, float(duration), layout, int(x), int(y), int(w), int(h), keyword, output))


def barcode_tool(image_path=""):
    """掃描條碼/QR Code。用法：barcode [圖片路徑]"""
    from bot import execute_barcode
    print(execute_barcode(image_path))


def bluetooth_tool(action, mac=""):
    """藍牙操作。用法：bluetooth <scan|connect|disconnect|list> [MAC]"""
    from bot import execute_bluetooth
    print(execute_bluetooth(action, mac))


def browser_advanced_tool(action, selector="", value="", name="", tab_index=0, timeout=30.0, url_pattern=""):
    """進階瀏覽器。用法：browser_advanced <action> [selector] [value] [name] [tab] [timeout] [url_pattern]"""
    from bot import execute_browser_advanced
    print(execute_browser_advanced(action, selector, value, name, int(tab_index), float(timeout), url_pattern))


def browser_control_tool(action, url="", selector="", text=""):
    """瀏覽器控制。用法：browser_control <open|goto|click|type|get_text|screenshot|close> [url] [selector] [text]"""
    from bot import execute_browser_control
    print(execute_browser_control(action, url, selector, text))


def calendar_tool(action, days=7, title="", start="", end="", description=""):
    """Google 日曆。用法：calendar <list|add> [天數] [標題] [開始] [結束] [說明]"""
    from bot import execute_calendar
    print(execute_calendar(action, int(days), title, start, end, description))


def clipboard_tool(action, text=""):
    """剪貼簿。用法：clipboard <get|set|history> [文字]"""
    from bot import execute_clipboard
    print(execute_clipboard(action, text))


def clipboard_image_tool(action, path=""):
    """剪貼簿圖片。用法：clipboard_image <get|set> [路徑]"""
    from bot import execute_clipboard_image
    print(execute_clipboard_image(action, path))


def cloud_storage_tool(action, path, drive_id="root"):
    """雲端儲存。用法：cloud_storage <upload|download|list> <路徑> [drive_id]"""
    from bot import execute_cloud_storage
    print(execute_cloud_storage(action, path, drive_id))


def compare_stocks_tool(symbols_str, metrics_str="all"):
    """比較股票。用法：compare_stocks <代號1,代號2,...> [指標]"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import compare_stocks as _compare_stocks
    symbols = symbols_str.split(",")
    metrics = metrics_str.split(",") if metrics_str != "all" else None
    print(_compare_stocks(symbols, metrics))


def database_tool(type_, db, sql, name=""):
    """資料庫。用法：database <sqlite|mysql> <路徑/host> <SQL> [名稱]"""
    from bot import execute_database
    print(execute_database(type_, db, sql, name))


def ddg_search_tool(query, region="zh-tw", max_results=5):
    """DuckDuckGo 搜尋。用法：ddg_search <關鍵字> [地區] [數量]"""
    from bot import execute_ddg_search
    print(execute_ddg_search(query, region, int(max_results)))


def desktop_control_tool(action, x=None, y=None, text=None, app=None, direction="down", amount=3, monitor=None):
    """桌面控制。用法：desktop_control <action> [x] [y] [text] [app] [direction] [amount] [monitor]"""
    from bot import execute_desktop_control
    result = execute_desktop_control(action, x=int(x) if x else None, y=int(y) if y else None,
                                     text=text, app=app, direction=direction, amount=int(amount),
                                     monitor=int(monitor) if monitor else None)
    print(result.get("message", str(result)))


def device_manager_tool(action, name="", keyword=""):
    """裝置管理員。用法：device_manager <list|enable|disable> [名稱] [關鍵字]"""
    from bot import execute_device_manager
    print(execute_device_manager(action, name, keyword))


def disk_backup_tool(action, src="", dest=""):
    """磁碟備份。用法：disk_backup <backup|restore|list> [來源] [目標]"""
    from bot import execute_disk_backup
    print(execute_disk_backup(action, src, dest))


def display_tool(action, level=None):
    """顯示設定。用法：display <get_brightness|set_brightness|get_resolution|list_resolutions> [亮度]"""
    from bot import execute_display
    print(execute_display(action, int(level) if level else None))


def docker_tool(action, name=""):
    """Docker 操作。用法：docker <list|start|stop|logs|images> [容器名]"""
    from bot import execute_docker
    print(execute_docker(action, name))


def document_control_tool(action, path, content="", sheet=None):
    """文件控制。用法：document_control <read|write|list_sheets> <路徑> [內容] [工作表]"""
    from bot import execute_document
    print(execute_document(action, path, content, sheet))


def download_file_tool(url, save_path=""):
    """下載檔案。用法：download_file <URL> [儲存路徑]"""
    from bot import execute_download_file
    print(execute_download_file(url, save_path))


def dropbox_tool(action, local, remote, token=""):
    """Dropbox 操作。用法：dropbox <upload|download> <本地> <遠端> [token]"""
    from bot import execute_dropbox
    print(execute_dropbox(action, local, remote, token))


def email_control_tool(host, user, password, folder="INBOX", count=5):
    """Email 控制（讀取）。用法：email_control <host> <user> <password> [folder] [count]"""
    from bot import execute_email_read
    print(execute_email_read(host, user, password, folder, int(count)))


def emotion_detect_tool(action, text="", image_path=""):
    """情緒偵測。用法：emotion_detect <text|image|both> [文字] [圖片路徑]"""
    from bot import execute_emotion_detect
    print(execute_emotion_detect(action, text, image_path))


def encrypt_file_tool(action, path, password):
    """加解密檔案。用法：encrypt_file <encrypt|decrypt> <路徑> <密碼>"""
    from bot import execute_encrypt_file
    print(execute_encrypt_file(action, path, password))


def env_var_tool(action, name="", value="", permanent="false"):
    """環境變數。用法：env_var <get|set|list|delete> [名稱] [值] [permanent]"""
    from bot import execute_env_var
    print(execute_env_var(action, name, value, permanent.lower() == "true"))


def file_system_tool(action, path="", dest="", content="", keyword=""):
    """檔案系統。用法：file_system <list|read|write|delete|copy|move|search|info> [路徑] [目標] [內容] [關鍵字]"""
    from bot import execute_file_system
    print(execute_file_system(action, path, dest, content, keyword))


def file_tools_tool(action, path, dest="", pattern="", replacement="", ext=""):
    """檔案工具。用法：file_tools <batch_rename|sync> <路徑> [目標] [pattern] [replacement] [ext]"""
    from bot import execute_file_tools
    print(execute_file_tools(action, path, dest, pattern, replacement, ext))


def file_transfer_tool(action, source, dest=""):
    """檔案傳輸。用法：file_transfer <zip|unzip|download|upload> <來源> [目標]"""
    from bot import execute_file_transfer
    print(execute_file_transfer(action, source, dest))


def find_image_on_screen_tool(image_path, confidence=0.8):
    """在螢幕上找圖。用法：find_image_on_screen <圖片路徑> [信心值]"""
    from bot import execute_find_image
    print(execute_find_image(image_path, float(confidence)))


def generate_image_tool(prompt, width=512, height=512, overlay_text=""):
    """生成圖片。用法：generate_image <prompt> [寬] [高] [疊加文字]"""
    from bot import fetch_image, add_text_overlay
    image_bytes = fetch_image(prompt, int(width), int(height))
    if image_bytes and overlay_text:
        image_bytes = add_text_overlay(image_bytes, overlay_text)
    if image_bytes:
        out = str(Path.home() / "Desktop" / f"gen_{int(time.time())}.png")
        with open(out, "wb") as f:
            f.write(image_bytes)
        print(f"✅ 圖片已生成：{out}")
    else:
        print("圖片生成失敗")


def get_candlestick_chart_tool(symbol, period="3mo"):
    """K線圖。用法：get_candlestick_chart <代號> [期間]"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import generate_candlestick
    img_bytes, pattern_str = generate_candlestick(symbol, period)
    if img_bytes:
        out = str(Path.home() / "Desktop" / f"candlestick_{symbol}_{int(time.time())}.png")
        with open(out, "wb") as f:
            f.write(img_bytes)
        print(f"✅ K線圖已儲存：{out}\n\n{pattern_str}")
    else:
        print(pattern_str)


def get_crypto_tool(coin, vs_currency="usd"):
    """查詢加密貨幣。用法：get_crypto <幣種> [vs貨幣]"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import fetch_crypto
    print(fetch_crypto(coin, vs_currency))


def get_earnings_tool(symbol):
    """查詢財報。用法：get_earnings <代號>"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import fetch_earnings
    print(fetch_earnings(symbol))


def get_etf_tool(symbol):
    """查詢 ETF。用法：get_etf <代號>"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import fetch_etf
    print(fetch_etf(symbol))


def get_finance_news_tool(source="all", count=5):
    """財經新聞。用法：get_finance_news [來源] [數量]"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import fetch_finance_news
    print(fetch_finance_news(source, int(count)))


def get_forex_tool(pair):
    """查詢匯率。用法：get_forex <貨幣對>"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import fetch_forex
    print(fetch_forex(pair))


def get_fundamentals_tool(symbol):
    """查詢基本面。用法：get_fundamentals <代號>"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import fetch_fundamentals
    print(fetch_fundamentals(symbol))


def get_macro_tool(indicator):
    """查詢總經指標。用法：get_macro <cpi|gdp|unemployment|fed_rate|nonfarm>"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import fetch_macro
    print(fetch_macro(indicator))


def get_market_sentiment_tool():
    """查詢市場情緒。用法：get_market_sentiment"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import fetch_market_sentiment
    print(fetch_market_sentiment())


def get_stock_tool(symbol, period="1mo"):
    """查詢股票。用法：get_stock <代號> [期間]"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import fetch_stock
    print(fetch_stock(symbol, period))


def get_stock_advanced_tool(symbol, indicators="macd,bb,kd"):
    """進階技術分析。用法：get_stock_advanced <代號> [指標]"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import fetch_stock_advanced
    ind_list = indicators.split(",") if indicators else None
    print(fetch_stock_advanced(symbol, ind_list))


def get_weather_tool(city):
    """查詢天氣。用法：get_weather <城市>"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import fetch_weather
    print(fetch_weather(city))


def git_tool(action, repo=".", message="", branch="master"):
    """Git 操作。用法：git <status|log|pull|add|commit|push|diff> [repo] [message] [branch]"""
    from bot import execute_git
    print(execute_git(action, repo, message, branch))


def goal_manager_tool(action, goal="", goal_id="", steps="", priority="normal"):
    """目標管理。用法：goal_manager <add|list|update|delete|progress> [目標] [id] [步驟] [優先]"""
    from bot import execute_goal_manager
    print(execute_goal_manager(action, goal, goal_id, steps, priority))


def google_trends_tool(keywords_str, timeframe="today 3-m", geo="TW"):
    """Google Trends。用法：google_trends <關鍵字1,關鍵字2,...> [時間] [地區]"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import fetch_google_trends
    keywords = keywords_str.split(",")
    print(fetch_google_trends(keywords, timeframe, geo))


def hardware_tool():
    """硬體監控。用法：hardware"""
    from bot import execute_hardware
    print(execute_hardware())


def image_edit_tool(action, path, *params):
    """圖片編輯。用法：image_edit <crop|resize|text|merge|rotate|flip> <路徑> [參數...]"""
    from bot import execute_image_edit
    print(execute_image_edit(action, path, *params))


def image_tools_tool(action, path="", quality=75, width=0, height=0, target_lang="zh-TW"):
    """圖片工具。用法：image_tools <compress|batch|ocr_translate> [路徑] [quality] [width] [height] [lang]"""
    from bot import execute_image_tools
    print(execute_image_tools(action, path, int(quality), int(width), int(height), target_lang))


def knowledge_base_tool(action, content="", query="", tag="", kb_id=""):
    """知識庫。用法：knowledge_base <add|search|list|delete|export> [content] [query] [tag] [id]"""
    from bot import execute_knowledge_base
    print(execute_knowledge_base(action, content, query, tag, kb_id))


def long_term_memory_tool(action, chat_id="0", content="", memory_id=""):
    """長期記憶。用法：long_term_memory <save|list|delete> [chat_id] [內容] [id]"""
    if action == "save":
        memory_save(int(chat_id), content)
    elif action == "list":
        memory_list(int(chat_id))
    elif action == "delete":
        memory_del(int(chat_id), int(memory_id))
    else:
        print(f"可用操作：save, list, delete")


def lookup_tool(action, ip="", amount=1.0, from_cur="USD", to_cur="TWD"):
    """查詢工具。用法：lookup <ip|currency> [ip] [amount] [from] [to]"""
    from bot import execute_lookup
    print(execute_lookup(action, ip, float(amount), from_cur, to_cur))


def manage_schedule_tool(action, name="", time_str="", script=""):
    """排程管理。用法：manage_schedule <list|add|remove|enable|disable> [名稱] [時間] [腳本]"""
    from bot import execute_manage_schedule
    print(execute_manage_schedule(action, name, time_str, script))


def media_tool(action, device_name=""):
    """媒體控制。用法：media <play_pause|next|prev|stop|volume_up|volume_down|mute|list_devices|switch> [裝置名]"""
    from bot import execute_media
    print(execute_media(action, device_name))


def monitor_config_tool():
    """螢幕設定。用法：monitor_config"""
    from bot import execute_monitor_config
    print(execute_monitor_config())


def multi_deploy_tool(action, remote_host="", remote_user="", remote_pass="", remote_path="/tmp/niu_bot"):
    """多機部署。用法：multi_deploy <deploy|status|rollback> [host] [user] [pass] [path]"""
    from bot import execute_multi_deploy
    print(execute_multi_deploy(action, remote_host, remote_user, remote_pass, remote_path))


def multi_perspective_tool(topic, lang="zh-tw"):
    """多角度分析。用法：multi_perspective <主題> [語言]"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import multi_perspective
    print(multi_perspective(topic, lang))


def network_config_tool(action, name="", ip="", dns1="", dns2="", domain="", duration=10):
    """網路設定。用法：network_config <list|set_ip|set_dns|get_dns|hosts_list|hosts_add|hosts_remove|traffic> [參數...]"""
    from bot import execute_network_config
    print(execute_network_config(action, name, ip, dns1, dns2, domain, int(duration)))


def network_diag_tool(action, host, ports="22,80,443,3306,3389,8080"):
    """網路診斷。用法：network_diag <ping|traceroute|portscan> <host> [ports]"""
    from bot import execute_network_diag
    print(execute_network_diag(action, host, ports))


def news_monitor_tool(action, keywords="", interval=300, duration=3600):
    """新聞監控。用法：news_monitor <start|stop|status> [關鍵字] [間隔秒] [持續秒]"""
    from bot import execute_news_monitor
    print(execute_news_monitor(action, keywords, int(interval), int(duration)))


def news_search_tool(query, lang="zh-TW", count=6):
    """新聞搜尋。用法：news_search <關鍵字> [語言] [數量]"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import search_news
    print(search_news(query, lang, int(count)))


def nlp_tool(action, text):
    """NLP 工具。用法：nlp <summarize|sentiment|keywords|ner> <文字>"""
    from bot import execute_nlp
    print(execute_nlp(action, text))


def osint_search_tool(action, query="", target="", limit=10):
    """OSINT 搜尋。用法：osint_search <web_search|news_search|ip_osint|domain_osint|top_news> [query] [target] [limit]"""
    from bot import execute_osint_search
    print(execute_osint_search(action, query, target, int(limit)))


def password_mgr_tool(action, site, master, username="", password=""):
    """密碼管理。用法：password_mgr <save|get|list|delete> <網站> <主密碼> [帳號] [密碼]"""
    from bot import execute_password_mgr
    print(execute_password_mgr(action, site, master, username, password))


def pdf_edit_tool(action, path="", output="", paths="", text=""):
    """PDF 編輯。用法：pdf_edit <merge|split|watermark|info> [路徑] [輸出] [paths] [text]"""
    from bot import execute_pdf_edit
    print(execute_pdf_edit(action, path, output, paths, text))


def pdf_image_tool(path, output_dir="", dpi=150):
    """PDF 轉圖片。用法：pdf_image <路徑> [輸出資料夾] [dpi]"""
    from bot import execute_pdf_to_image
    print(execute_pdf_to_image(path, output_dir, int(dpi)))


def pentest_tool(action, target="", port_range="1-1000", timeout=2):
    """滲透測試。用法：pentest <scan|vuln_check|banner_grab|ssl_check> [target] [port_range] [timeout]"""
    from bot import execute_pentest
    print(execute_pentest(action, target, port_range, int(timeout)))


def portfolio_tool(action, chat_id=0, symbol="", shares=0, cost=0):
    """投資組合。用法：portfolio <add|remove|view|clear> [chat_id] [symbol] [shares] [cost]"""
    from bot import execute_portfolio
    print(execute_portfolio(action, int(chat_id), symbol, float(shares), float(cost)))


def power_control_tool(action):
    """電源控制。用法：power_control <sleep|restart|shutdown|hibernate|lock>"""
    from bot import execute_power
    print(execute_power(action))


def pptx_control_tool(action, path, slides=""):
    """PowerPoint 控制。用法：pptx_control <read|create|add_slide> <路徑> [slides_json]"""
    from bot import execute_pptx
    print(execute_pptx(action, path, slides))


def proactive_alert_tool(action, name="", condition="", threshold="", target="", interval=60):
    """主動警報。用法：proactive_alert <add|remove|list|status> [名稱] [條件] [閾值] [目標] [間隔秒]"""
    from bot import execute_proactive_alert
    print(execute_proactive_alert(action, name, condition, threshold, target, int(interval)))


def ptt_search_tool(keyword, board="Gossiping", count=5):
    """PTT 搜尋。用法：ptt_search <關鍵字> [看板] [數量]"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import ptt_search as _ptt_search
    print(_ptt_search(keyword, board, int(count)))


def push_notify_tool(platform, message, webhook_or_token):
    """推播通知。用法：push_notify <discord|line|slack> <訊息> <webhook_or_token>"""
    from bot import execute_push_notify
    print(execute_push_notify(platform, message, webhook_or_token))


def qr_code_tool(action, content="", path="", duration=30.0):
    """QR Code。用法：qr_code <generate|scan|watch> [content] [path] [duration]"""
    from bot import execute_qr_code
    print(execute_qr_code(action, content, path, float(duration)))


def read_screen_tool(question="描述螢幕上有什麼", monitor=1):
    """讀取螢幕。用法：read_screen [問題] [螢幕號]"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import fetch_read_screen
    print(fetch_read_screen(question, int(monitor)))


def read_webpage_tool(url, max_chars=3000):
    """讀取網頁。用法：read_webpage <URL> [最大字數]"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import read_webpage
    print(read_webpage(url, int(max_chars)))


def registry_tool(action, key, value_name="", value=""):
    """登錄檔操作。用法：registry <read|write|delete|list> <key> [value_name] [value]"""
    from bot import execute_registry
    print(execute_registry(action, key, value_name, value))


def reminder_tool(time_str, message):
    """提醒。用法：reminder <HH:MM或秒數> <訊息>"""
    from bot import execute_reminder
    print(execute_reminder(time_str, message))


def report_tool(title, data_json, output=""):
    """報告生成。用法：report <標題> <資料JSON> [輸出路徑]"""
    from bot import execute_report
    print(execute_report(title, data_json, output))


def restore_point_tool(action, description=""):
    """系統還原點。用法：restore_point <create|list> [說明]"""
    from bot import execute_restore_point
    print(execute_restore_point(action, description))


def run_code_tool(type_, code):
    """執行程式碼。用法：run_code <python|powershell|cmd|javascript> <code>"""
    from bot import execute_run_code
    print(execute_run_code(type_, code))


def screen_vision_tool(question="請描述這個畫面上有什麼，以及目前電腦在做什麼事。"):
    """螢幕視覺分析。用法：screen_vision [問題]"""
    from bot import execute_screen_vision
    analysis, img_bytes = execute_screen_vision(question)
    if img_bytes:
        out = str(Path.home() / "Desktop" / f"vision_{int(time.time())}.png")
        with open(out, "wb") as f:
            f.write(img_bytes)
        print(f"{analysis}\n\n截圖已儲存：{out}")
    else:
        print(analysis)


def scroll_at_tool(direction="down", amount=3, x=None, y=None, monitor=1, description=""):
    """指定位置滾動。用法：scroll_at [方向] [格數] [x] [y] [螢幕號] [描述]"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import fetch_scroll_at
    print(fetch_scroll_at(direction, int(amount),
                          int(x) if x and x != "None" else None,
                          int(y) if y and y != "None" else None,
                          int(monitor), description))


def self_benchmark_tool(action):
    """自我評測。用法：self_benchmark <run|report|compare>"""
    from bot import execute_self_benchmark
    print(execute_self_benchmark(action))


def send_email_tool(to, subject, body):
    """發送 Email。用法：send_email <收件人> <主旨> <內容>"""
    from bot import execute_send_email
    print(execute_send_email(to, subject, body))


def send_voice_tool(text, voice="zh-CN-YunxiNeural"):
    """生成語音。用法：send_voice <文字> [語音]"""
    try:
        sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
        from bot import generate_voice_ogg
        ogg_data = generate_voice_ogg(text, voice)
        out = str(Path.home() / "Desktop" / f"voice_{int(time.time())}.ogg")
        with open(out, "wb") as f:
            f.write(ogg_data)
        print(f"✅ 語音已生成：{out}")
    except Exception as e:
        print(f"語音生成失敗：{e}")


def smart_home_tool(action, device="", value="", host="", token=""):
    """智慧家居。用法：smart_home <list|control|status|scene> [device] [value] [host] [token]"""
    from bot import execute_smart_home
    print(execute_smart_home(action, device, value, host, token))


def software_tool(action, name="", keyword=""):
    """軟體管理。用法：software <list|install|uninstall> [名稱] [關鍵字]"""
    from bot import execute_software
    print(execute_software(action, name, keyword))


def ssh_sftp_tool(action, host, user, password, command="", local="", remote="", port=22):
    """SSH/SFTP。用法：ssh_sftp <run|upload|download> <host> <user> <pass> [command] [local] [remote] [port]"""
    from bot import execute_ssh_sftp
    print(execute_ssh_sftp(action, host, user, password, command, local, remote, int(port)))


def startup_tool(action, name="", command=""):
    """開機自啟。用法：startup <list|add|remove> [名稱] [指令]"""
    from bot import execute_startup
    print(execute_startup(action, name, command))


def system_monitor_tool(action, target=""):
    """系統監控。用法：system_monitor <sysinfo|process_list|process_kill|disk_usage|network|battery> [target]"""
    from bot import execute_system_monitor
    print(execute_system_monitor(action, target))


def system_tools_tool(action, **kwargs):
    """系統工具。用法：system_tools <event_log|usb_list|firewall_list|printer_list|...> [參數...]"""
    from bot import execute_system_tools
    print(execute_system_tools(action, **kwargs))


def think_as_tool(person, question, list_available="false"):
    """角色思考。用法：think_as <人物> <問題> [list]"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import execute_think_as
    print(execute_think_as(person, question, list_available.lower() == "true"))


def threat_intel_tool(action, target="", api_key=""):
    """威脅情報。用法：threat_intel <ip_check|domain_check|hash_check|cve_search> [target] [api_key]"""
    from bot import execute_threat_intel
    print(execute_threat_intel(action, target, api_key))


def todo_list_tool(action, task="", todo_id=0):
    """任務清單。用法：todo_list <add|list|done|delete|clear> [task] [id]"""
    from bot import execute_todo
    print(execute_todo(action, task, int(todo_id)))


def tts_advanced_tool(action, text="", voice="zh-CN-YunxiNeural"):
    """進階 TTS。用法：tts_advanced <speak|list_voices> [文字] [語音]"""
    from bot import execute_tts_advanced
    print(execute_tts_advanced(action, text, voice))


def user_account_tool(action, username="", password=""):
    """使用者帳戶。用法：user_account <list|create|delete> [帳號] [密碼]"""
    from bot import execute_user_account
    print(execute_user_account(action, username, password))


def video_gif_tool(path, start=0, duration=5.0, output="", fps=10):
    """影片轉 GIF。用法：video_gif <路徑> [起始秒] [持續秒] [輸出] [fps]"""
    from bot import execute_video_gif
    print(execute_video_gif(path, float(start), float(duration), output, int(fps)))


def video_process_tool(action, path, second=0, start=0, end=0, output=""):
    """影片處理。用法：video_process <screenshot|trim|info|to_gif> <路徑> [秒數] [起始] [結束] [輸出]"""
    from bot import execute_video_process
    print(execute_video_process(action, path, float(second), float(start), float(end), output))


def virtual_desktop_tool(action):
    """虛擬桌面。用法：virtual_desktop <left|right|new>"""
    from bot import execute_vdesktop
    print(execute_vdesktop(action))


def voice_cmd_tool(action, duration=300.0, language="zh-TW"):
    """語音命令。用法：voice_cmd <start|stop|status> [持續秒] [語言]"""
    from bot import execute_voice_cmd
    print(execute_voice_cmd(action, float(duration), language))


def voice_id_tool(action, name="", audio_path="", duration=5):
    """聲紋辨識。用法：voice_id <register|identify|list|delete> [名稱] [音檔] [秒數]"""
    from bot import execute_voice_id
    print(execute_voice_id(action, name, audio_path, int(duration)))


def volume_tool(action, level=None):
    """音量控制。用法：volume <get|set|mute|unmute> [音量]"""
    from bot import execute_volume
    print(execute_volume(action, int(level) if level else None))


def vpn_tool(action, name="", user="", password=""):
    """VPN 控制。用法：vpn <list|connect|disconnect> [名稱] [帳號] [密碼]"""
    from bot import execute_vpn
    print(execute_vpn(action, name, user, password))


def wait_seconds_tool(seconds):
    """等待。用法：wait_seconds <秒數>"""
    from bot import execute_wait_seconds
    print(execute_wait_seconds(float(seconds)))


def web_scrape_tool(action, url="", selector="body", interval=2.0, region="full"):
    """網頁爬取。用法：web_scrape <scrape|monitor|screenshot> [url] [selector] [interval] [region]"""
    from bot import execute_web_scrape
    print(execute_web_scrape(action, url, selector, float(interval), region))


def webpage_shot_tool(action, url, selector="body", interval=60.0, duration=3600.0):
    """網頁截圖。用法：webpage_shot <screenshot|monitor> <url> [selector] [interval] [duration]"""
    from bot import execute_webpage_shot
    print(execute_webpage_shot(action, url, selector, float(interval), float(duration)))


def wikipedia_search_tool(query, lang="zh"):
    """Wikipedia 搜尋。用法：wikipedia_search <關鍵字> [語言]"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import wikipedia_search as _wikipedia_search
    print(_wikipedia_search(query, lang))


def win_notify_relay_tool(action, duration=3600.0, filter_app=""):
    """Windows 通知轉發。用法：win_notify_relay <start|stop|status> [持續秒] [filter_app]"""
    from bot import execute_win_notify_relay
    print(execute_win_notify_relay(action, float(duration), filter_app))


def window_control_tool(action, keyword=""):
    """視窗控制。用法：window_control <list|focus|maximize|minimize|close|restore> [關鍵字]"""
    from bot import execute_window_control
    print(execute_window_control(action, keyword))


def window_manager_tool(action="list", window_name=""):
    """視窗管理員。用法：window_manager <list|focus|maximize|minimize|close> [視窗名]"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import fetch_window_manager
    print(fetch_window_manager(action, window_name))


def windows_update_tool(action):
    """Windows 更新。用法：windows_update <list|install|check>"""
    from bot import execute_win_update
    print(execute_win_update(action))


def workflow_tool(action, name="", steps=""):
    """工作流程。用法：workflow <run|save|list|delete> [名稱] [steps_json]"""
    from bot import execute_workflow
    print(execute_workflow(action, name, steps))


def youtube_summary_tool(url):
    """YouTube 摘要。用法：youtube_summary <URL>"""
    sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    from bot import youtube_summary as _youtube_summary
    print(_youtube_summary(url))


def tg_auto_reply_tool(action="start", stop_time="", duration="30"):
    """Telegram 自動回覆。用法：tg_auto_reply [start|stop] [stop_time] [duration_minutes]"""
    import subprocess
    if action == "stop":
        subprocess.run(["powershell.exe", "-Command", "Get-Process python3.12 -ErrorAction SilentlyContinue | Stop-Process -Force"], capture_output=True)
        print("自動回覆已停止")
        return
    # 修改 tg_auto_reply.py 的 STOP_TIME 並執行
    script = "C:/Users/blue_/Desktop/測試檔案/tg_auto_reply.py"
    import re
    with open(script, "r", encoding="utf-8") as f:
        content = f.read()
    if stop_time:
        content = re.sub(r'STOP_TIME = ".*?"', f'STOP_TIME = "{stop_time}"', content)
    else:
        from datetime import datetime, timedelta
        end = (datetime.now() + timedelta(minutes=int(duration))).strftime("%H:%M")
        content = re.sub(r'STOP_TIME = ".*?"', f'STOP_TIME = "{end}"', content)
    with open(script, "w", encoding="utf-8") as f:
        f.write(content)
    subprocess.Popen(["C:/Users/blue_/AppData/Local/Microsoft/WindowsApps/python3.12.exe", "-u", script])
    print(f"自動回覆已開啟，監控到 {stop_time or end}")


# ── 主程式 ──────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)

    tool = sys.argv[1]
    args = sys.argv[2:]

    tools = {
        "weather":       lambda: get_weather(args[0]),
        "stock":         lambda: get_stock(args[0], args[1] if len(args) > 1 else "1mo"),
        "image":         lambda: generate_image(args[0], args[1] if len(args) > 1 else ""),
        "screenshot":    lambda: screenshot(),
        "click":         lambda: click(*args),
        "double_click":  lambda: double_click(*args),
        "right_click":   lambda: right_click(*args),
        "move":          lambda: move(*args),
        "type":          lambda: type_text(" ".join(args)),
        "press":         lambda: press_key(args[0]),
        "open":          lambda: open_app(" ".join(args)),
        "scroll":        lambda: scroll(*args),
        "pos":           lambda: pos(),
        "bot_status":    lambda: bot_status(),
        "bot_restart":   lambda: bot_restart(),
        "schedule_list": lambda: schedule_list(),
        "schedule_add":  lambda: schedule_add(args[0], args[1], args[2]),
        "schedule_del":  lambda: schedule_del(args[0]),
        "memory_save":   lambda: memory_save(int(args[0]), " ".join(args[1:])),
        "memory_list":   lambda: memory_list(int(args[0])),
        "memory_del":    lambda: memory_del(int(args[0]), int(args[1])),
        "vision":        lambda: vision(" ".join(args) if args else "請描述這個畫面上有什麼，以及目前電腦在做什麼事。"),
        "find_image":    lambda: find_image(args[0], float(args[1]) if len(args) > 1 else 0.8),
        "browser":       lambda: browser(args[0], *args[1:]),
        "window_list":   lambda: window_list(),
        "window_focus":  lambda: window_focus(" ".join(args)),
        "window_close":  lambda: window_close(" ".join(args)),
        "window_min":    lambda: window_min(" ".join(args)),
        "window_max":    lambda: window_max(" ".join(args)),
        "hotkey":        lambda: hotkey(*args),
        "clipboard_get": lambda: clipboard_get(),
        "clipboard_set": lambda: clipboard_set(" ".join(args)),
        "file_list":     lambda: file_list(args[0] if args else "."),
        "file_read":     lambda: file_read(args[0]),
        "file_write":    lambda: file_write(args[0], " ".join(args[1:])),
        "file_delete":   lambda: file_delete(args[0]),
        "file_copy":     lambda: file_copy(args[0], args[1]),
        "file_move":     lambda: file_move(args[0], args[1]),
        "file_search":   lambda: file_search(args[0], args[1]),
        "sysinfo":       lambda: sysinfo(),
        "process_list":  lambda: process_list(),
        "process_kill":  lambda: process_kill(args[0]),
        "notify":        lambda: notify(args[0], " ".join(args[1:])),
        "tts":           lambda: tts(" ".join(args)),
        "record_start":  lambda: record_start(),
        "record_stop":   lambda: record_stop(),
        "record_play":   lambda: record_play(args[0]),
        "email":         lambda: send_email(args[0], args[1], " ".join(args[2:])),
        "stt":           lambda: stt(args[0] if args and args[0].isdigit() else 5, args[0] if args and not args[0].isdigit() else "", args[1] if len(args)>1 else "zh-TW"),
        "ocr":           lambda: ocr(args[0] if args else ""),
        "workflow_run":  lambda: workflow_run(args[0]),
        "workflow_save": lambda: workflow_save(args[0], " ".join(args[1:])),
        "screen_watch":  lambda: screen_watch(args[0], args[1], float(args[2]) if len(args)>2 else 2.0),
        "monitors":      lambda: monitors(),
        "zip":           lambda: zip_files(args[0], args[1]),
        "unzip":         lambda: unzip(args[0], args[1]),
        "download":      lambda: download(args[0], args[1] if len(args)>1 else ""),
        "print_file":    lambda: print_file(args[0]),
        "wifi_list":       lambda: wifi_list(),
        "wifi_connect":    lambda: wifi_connect(args[0], args[1]),
        "screen_stream":   lambda: screen_stream(int(args[0]) if args else 10),
        "wake_listen":     lambda: wake_listen(args[0] if args else "小牛馬"),
        "drag":            lambda: drag(*args),
        "right_menu":      lambda: right_menu(args[0], args[1], args[2] if len(args)>2 else ""),
        "ai_plan":         lambda: ai_plan(" ".join(args)),
        "clipboard_history": lambda: clipboard_history(),
        "vdesktop":        lambda: vdesktop(args[0]),
        "power":           lambda: power(args[0]),
        "bt_scan":         lambda: bt_scan(),
        "bt_connect":      lambda: bt_connect(args[0]),
        "run_python":      lambda: run_python(" ".join(args)),
        "run_shell":       lambda: run_shell(" ".join(args)),
        "word_read":       lambda: word_read(args[0]),
        "word_write":      lambda: word_write(args[0], " ".join(args[1:])),
        "excel_read":      lambda: excel_read(args[0], args[1] if len(args)>1 else None),
        "excel_write":     lambda: excel_write(args[0], args[1], " ".join(args[2:])),
        "pdf_read":        lambda: pdf_read(args[0]),
        "screen_diff":     lambda: screen_diff(float(args[0]) if args else 2.0, args[1] if len(args)>1 else "full"),
        "scrape":          lambda: scrape(args[0], args[1] if len(args)>1 else "body"),
        "img_edit":        lambda: img_edit(args[0], args[1], *args[2:]),
        "gdrive_upload":   lambda: gdrive_upload(args[0], args[1] if len(args)>1 else "root"),
        "gdrive_download": lambda: gdrive_download(args[0], args[1]),
        "db_query":        lambda: db_query(args[0], " ".join(args[1:])),
        "db_mysql":        lambda: db_mysql(args[0], args[1], " ".join(args[2:])),
        "encrypt":         lambda: encrypt(args[0], args[1]),
        "decrypt":         lambda: decrypt(args[0], args[1]),
        "clipboard_watch": lambda: clipboard_watch(float(args[0]) if args else 30.0),
        "qr_gen":          lambda: qr_gen(args[0], args[1] if len(args)>1 else ""),
        "qr_scan":         lambda: qr_scan(args[0] if args else ""),
        "screen_record":   lambda: screen_record(float(args[0]) if args else 10.0, args[1] if len(args)>1 else ""),
        "webcam":          lambda: webcam_capture(args[0] if args else ""),
        "translate":       lambda: translate(" ".join(args[:1] if len(args)==1 else [args[0]]), args[1] if len(args)>1 else "zh-TW", args[2] if len(args)>2 else "auto"),
        "chart":           lambda: chart(args[0], args[1], args[2] if len(args)>2 else "", args[3] if len(args)>3 else ""),
        "pptx_read":       lambda: pptx_read(args[0]),
        "pptx_create":     lambda: pptx_create(args[0], " ".join(args[1:])),
        "api_call":        lambda: api_call(args[0], args[1], args[2] if len(args)>2 else "{}", args[3] if len(args)>3 else "{}"),
        "watchdog":        lambda: watchdog(args[0], args[1], float(args[2]) if len(args)>2 else 60.0),
        "ssh_run":         lambda: ssh_run(args[0], args[1], args[2], " ".join(args[3:])),
        "sftp_upload":     lambda: sftp_upload(args[0], args[1], args[2], args[3], args[4]),
        "sftp_download":   lambda: sftp_download(args[0], args[1], args[2], args[3], args[4]),
        "net_ping":        lambda: net_ping(args[0], int(args[1]) if len(args)>1 else 4),
        "net_traceroute":  lambda: net_traceroute(args[0]),
        "net_portscan":    lambda: net_portscan(args[0], args[1] if len(args)>1 else "22,80,443,3306,3389,8080"),
        "win_service":     lambda: win_service(args[0], args[1] if len(args)>1 else ""),
        "pdf_merge":       lambda: pdf_merge(args[0], args[1]),
        "pdf_split":       lambda: pdf_split(args[0], args[1]),
        "pdf_watermark":   lambda: pdf_watermark(args[0], args[1], args[2] if len(args)>2 else ""),
        "audio_convert":   lambda: audio_convert(args[0], args[1]),
        "audio_trim":      lambda: audio_trim(args[0], int(args[1]), int(args[2]), args[3] if len(args)>3 else ""),
        "discord_notify":  lambda: discord_notify(args[0], " ".join(args[1:])),
        "line_notify":     lambda: line_notify(args[0], " ".join(args[1:])),
        "disk_clean":      lambda: disk_clean(args[0] if args else "list"),
        "backup":          lambda: backup(args[0], args[1]),
        "registry_read":   lambda: registry_read(args[0], args[1] if len(args)>1 else ""),
        "registry_write":  lambda: registry_write(args[0], args[1], " ".join(args[2:])),
        "video_screenshot":lambda: video_screenshot(args[0], float(args[1]) if len(args)>1 else 0, args[2] if len(args)>2 else ""),
        "video_trim":      lambda: video_trim(args[0], float(args[1]), float(args[2]), args[3] if len(args)>3 else ""),
        "monitor_list":    lambda: monitor_list(),
        "email_read":      lambda: email_read(args[0], args[1], args[2], args[3] if len(args)>3 else "INBOX", int(args[4]) if len(args)>4 else 5),
        "gcal_list":       lambda: gcal_list(int(args[0]) if args else 7),
        "gcal_add":        lambda: gcal_add(args[0], args[1], args[2], args[3] if len(args)>3 else ""),
        "global_hotkey":   lambda: global_hotkey_listen(args[0], args[1], float(args[2]) if len(args)>2 else 60.0),
        "git_op":          lambda: git_op(args[0], args[1] if len(args)>1 else ".", args[2] if len(args)>2 else "", args[3] if len(args)>3 else "origin", args[4] if len(args)>4 else "master"),
        "hw_monitor":      lambda: hw_monitor(),
        "report_gen":      lambda: report_gen(args[0], args[1], args[2] if len(args)>2 else ""),
        "dropbox_upload":  lambda: dropbox_upload(args[0], args[1], args[2] if len(args)>2 else ""),
        "dropbox_download":lambda: dropbox_download(args[0], args[1], args[2] if len(args)>2 else ""),
        "docker_op":       lambda: docker_op(args[0], args[1] if len(args)>1 else ""),
        "pdf_to_images":   lambda: pdf_to_images(args[0], args[1] if len(args)>1 else "", int(args[2]) if len(args)>2 else 150),
        "barcode_scan":    lambda: barcode_scan(args[0] if args else ""),
        "nlp_summarize":   lambda: nlp_summarize(" ".join(args)),
        "nlp_sentiment":   lambda: nlp_sentiment(" ".join(args)),
        "vpn_control":     lambda: vpn_control(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", args[3] if len(args)>3 else ""),
        "sys_restore":     lambda: sys_restore(args[0], " ".join(args[1:]) if len(args)>1 else ""),
        "disk_analyze":    lambda: disk_analyze(args[0] if args else "C:/", int(args[1]) if len(args)>1 else 10),
        "face_detect":     lambda: face_detect(args[0] if args else "", args[1] if len(args)>1 else ""),
        "video_to_gif":    lambda: video_to_gif(args[0], float(args[1]) if len(args)>1 else 0, float(args[2]) if len(args)>2 else 5.0, args[3] if len(args)>3 else "", int(args[4]) if len(args)>4 else 10),
        "excel_chart":     lambda: excel_chart(args[0], args[1], args[2] if len(args)>2 else "bar", args[3] if len(args)>3 else ""),
        "speedtest":       lambda: speedtest_run(),
        "screenshot_compare": lambda: screenshot_compare(args[0] if args else "", args[1] if len(args)>1 else "", args[2] if len(args)>2 else ""),
        "set_reminder":    lambda: set_reminder(args[0], " ".join(args[1:])),
        "webpage_screenshot": lambda: webpage_screenshot(args[0], args[1] if len(args)>1 else ""),
        "web_monitor":     lambda: web_monitor(args[0], args[1] if len(args)>1 else "body", float(args[2]) if len(args)>2 else 60.0, float(args[3]) if len(args)>3 else 3600.0),
        "batch_rename":    lambda: batch_rename(args[0], args[1], args[2], args[3] if len(args)>3 else ""),
        "img_compress":    lambda: img_compress(args[0], int(args[1]) if len(args)>1 else 75, args[2] if len(args)>2 else ""),
        "batch_img_process": lambda: batch_img_process(args[0], args[1], int(args[2]) if len(args)>2 else 0, int(args[3]) if len(args)>3 else 0, int(args[4]) if len(args)>4 else 75),
        "ocr_translate":   lambda: ocr_translate(args[0] if args else "", args[1] if len(args)>1 else "zh-TW"),
        "ip_info":         lambda: ip_info(args[0] if args else ""),
        "currency":        lambda: currency(float(args[0]), args[1], args[2]),
        "event_log":       lambda: event_log(args[0] if args else "System", args[1] if len(args)>1 else "Error", int(args[2]) if len(args)>2 else 10),
        "tts_edge":        lambda: tts_edge(" ".join(args[:-1]) if len(args)>1 else " ".join(args), args[-1] if len(args)>1 and args[-1].startswith("zh-") else "zh-TW-HsiaoChenNeural"),
        "tts_voices":      lambda: tts_voices(),
        "send_email_attach": lambda: send_email_attach(args[0], args[1], args[2], args[3] if len(args)>3 else ""),
        "clipboard_img_get": lambda: clipboard_img_get(args[0] if args else ""),
        "clipboard_img_set": lambda: clipboard_img_set(args[0]),
        "usb_list":        lambda: usb_list(),
        "firewall":        lambda: firewall(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "in", int(args[3]) if len(args)>3 else 0, args[4] if len(args)>4 else "TCP"),
        "todo":            lambda: todo(args[0], " ".join(args[1:]) if args[0]=="add" else "", int(args[1]) if len(args)>1 and args[0] in ("done","delete") else 0),
        "file_sync":       lambda: file_sync(args[0], args[1], args[2].lower()=="true" if len(args)>2 else False),
        "sysres_chart":    lambda: sysres_chart(int(args[0]) if args else 10, args[1] if len(args)>1 else ""),
        "password_save":   lambda: password_save(args[0], args[1], args[2], args[3]),
        "password_get":    lambda: password_get(args[0], args[1]),
        "rdp_connect":     lambda: rdp_connect(args[0], args[1] if len(args)>1 else "", int(args[2]) if len(args)>2 else 1280, int(args[3]) if len(args)>3 else 720),
        "chrome_bookmarks": lambda: chrome_bookmarks(),
        "printer_list":    lambda: printer_list(),
        "printer_jobs":    lambda: printer_jobs(),
        "net_share":       lambda: net_share(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "Z:", args[3] if len(args)>3 else "", args[4] if len(args)>4 else ""),
        "font_list":       lambda: font_list(args[0] if args else ""),
        "volume_get":      lambda: volume_get(),
        "volume_set":      lambda: volume_set(int(args[0])),
        "volume_mute":     lambda: volume_mute(args[0].lower() != "false" if args else True),
        "brightness_get":  lambda: brightness_get(),
        "brightness_set":  lambda: brightness_set(int(args[0])),
        "resolution_list": lambda: resolution_list(),
        "media_control":   lambda: media_control(args[0]),
        "audio_devices":   lambda: audio_devices(),
        "audio_switch":    lambda: audio_switch(" ".join(args)),
        "software_list":   lambda: software_list(args[0] if args else ""),
        "software_install":lambda: software_install(" ".join(args)),
        "software_uninstall": lambda: software_uninstall(" ".join(args)),
        "startup_list":    lambda: startup_list(),
        "startup_add":     lambda: startup_add(args[0], " ".join(args[1:])),
        "startup_remove":  lambda: startup_remove(args[0]),
        "env_get":         lambda: env_get(args[0] if args else ""),
        "env_set":         lambda: env_set(args[0], args[1], len(args)>2 and args[2].lower()=="true"),
        "user_list":       lambda: user_list(),
        "user_create":     lambda: user_create(args[0], args[1]),
        "user_delete":     lambda: user_delete(args[0]),
        "win_update":      lambda: win_update(args[0] if args else "list"),
        "device_list":     lambda: device_list(args[0] if args else ""),
        "device_toggle":   lambda: device_toggle(args[0], len(args)<2 or args[1].lower()!="false"),
        "netadapter_list": lambda: netadapter_list(),
        "netadapter_toggle": lambda: netadapter_toggle(args[0], len(args)<2 or args[1].lower()!="false"),
        "dns_get":         lambda: dns_get(),
        "dns_set":         lambda: dns_set(args[0], args[1], args[2] if len(args)>2 else ""),
        "ip_config":       lambda: ip_config(args[0], args[1], args[2] if len(args)>2 else "255.255.255.0", args[3] if len(args)>3 else ""),
        "hosts_list":      lambda: hosts_list(),
        "hosts_add":       lambda: hosts_add(args[0], args[1]),
        "hosts_remove":    lambda: hosts_remove(args[0]),
        "net_traffic":     lambda: net_traffic(int(args[0]) if args else 5),
        "if_then":         lambda: if_then(args[0], args[1], " ".join(args[2:-1]) if len(args)>3 else args[2], float(args[-1]) if len(args)>3 else 300.0),
        "window_arrange":  lambda: window_arrange(args[0] if args else "side_by_side"),
        "region_ocr":      lambda: region_ocr(int(args[0]), int(args[1]), int(args[2]), int(args[3]), args[4] if len(args)>4 else "ch_tra"),
        "window_screenshot": lambda: window_screenshot(args[0], args[1] if len(args)>1 else ""),
        "firewall":          lambda: firewall(args[0], args[1] if len(args)>1 else "", int(args[2]) if len(args)>2 else None, args[3] if len(args)>3 else "TCP", args[4] if len(args)>4 else "Inbound"),
        "process_mgr":       lambda: process_mgr(args[0], args[1] if len(args)>1 else "", int(args[2]) if len(args)>2 and args[2].isdigit() else None, args[3] if len(args)>3 else "normal"),
        "power_plan":        lambda: power_plan(args[0], args[1] if len(args)>1 else "balanced"),
        "event_log":         lambda: event_log(args[0] if args else "System", args[1] if len(args)>1 else "Error", int(args[2]) if len(args)>2 else 10),
        "datetime_config":   lambda: datetime_config(args[0], args[1] if len(args)>1 else "", " ".join(args[2:]) if len(args)>2 else ""),
        "ui_auto":           lambda: ui_auto(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", " ".join(args[3:]) if len(args)>3 else ""),
        "macro":             lambda: macro(args[0], args[1] if len(args)>1 else "", int(args[2]) if len(args)>2 else 1, float(args[3]) if len(args)>3 else 10.0),
        "color_pick":        lambda: color_pick(int(args[0]), int(args[1]), args[2] if len(args)>2 else "pick", int(args[3]) if len(args)>3 else 100, int(args[4]) if len(args)>4 else 100),
        "webcam":            lambda: webcam(args[0], int(args[1]) if len(args)>1 and args[1].isdigit() else 0, float(args[2]) if len(args)>2 else 5.0, args[3] if len(args)>3 else ""),
        "multi_monitor":     lambda: multi_monitor(args[0], int(args[1]) if len(args)>1 else 1, args[2] if len(args)>2 else ""),
        "printer":           lambda: printer(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else ""),
        "wifi":              lambda: wifi(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else ""),
        "proxy":             lambda: proxy(args[0], args[1] if len(args)>1 else ""),
        "lock_screen":       lambda: lock_screen(args[0] if args else "lock"),
        "defender":          lambda: defender(args[0], args[1] if len(args)>1 else ""),
        "vision_loop":       lambda: vision_loop(" ".join(args), 20, 3.0, 120.0),
        "alert_monitor":     lambda: alert_monitor(args[0], args[1] if len(args)>1 else "80", args[2] if len(args)>2 else "", int(args[3]) if len(args)>3 else 30, int(args[4]) if len(args)>4 else 3600),
        "interval_schedule": lambda: interval_schedule(" ".join(args[:-2]) if len(args)>2 else args[0], float(args[-2]) if len(args)>1 else 60.0, int(args[-1]) if len(args)>2 else 0),
        "wait_for_text":     lambda: wait_for_text(" ".join(args[:-1]) if len(args)>1 else args[0], float(args[-1]) if len(args)>1 and args[-1].replace(".","").isdigit() else 60.0),
        "data_process":      lambda: data_process(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", " ".join(args[3:]) if len(args)>3 else ""),
        "wake_on_lan":       lambda: wake_on_lan(args[0], args[1] if len(args)>1 else "255.255.255.255", int(args[2]) if len(args)>2 else 9),
        "clipboard_history": lambda: clipboard_history(args[0] if args else "list", int(args[1]) if len(args)>1 else 0),
        "file_watcher":      lambda: file_watcher(args[0], args[1] if len(args)>1 else "all", args[2] if len(args)>2 else "", float(args[3]) if len(args)>3 else 3600.0),
        "pixel_watch":       lambda: pixel_watch(int(args[0]), int(args[1]), int(args[2]) if len(args)>2 else 10, float(args[3]) if len(args)>3 else 1.0, float(args[4]) if len(args)>4 else 60.0, args[5] if len(args)>5 else ""),
        "object_detect":     lambda: object_detect(" ".join(args[:-1]) if len(args)>1 else args[0], args[-1] if len(args)>1 and args[-1] in ("click","double_click","find") else "find"),
        "mouse_record":      lambda: mouse_record(args[0], args[1] if len(args)>1 else "", float(args[2]) if len(args)>2 else 10.0, int(args[3]) if len(args)>3 else 1, float(args[4]) if len(args)>4 else 1.0),
        "adb":               lambda: adb(args[0], int(args[1]) if len(args)>1 and args[1].isdigit() else 0, int(args[2]) if len(args)>2 and args[2].isdigit() else 0, 0, 0, " ".join(args[3:]) if len(args)>3 else ""),
        "wifi_hotspot":      lambda: wifi_hotspot(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else ""),
        "onedrive":          lambda: onedrive(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else ""),
        "ftp":               lambda: ftp(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", args[3] if len(args)>3 else "", args[4] if len(args)>4 else "", args[5] if len(args)>5 else ""),
        "wsl":               lambda: wsl(args[0], args[1] if len(args)>1 else "", " ".join(args[2:]) if len(args)>2 else ""),
        "hyperv":            lambda: hyperv(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else ""),
        "file_diff":         lambda: file_diff(args[0], args[1], args[2] if len(args)>2 else "", args[3] if len(args)>3 else "unified"),
        "screen_live":       lambda: screen_live(float(args[0]) if args else 0.5, float(args[1]) if len(args)>1 else 60.0, int(args[2]) if len(args)>2 else 50),
        # ── 影片生成 ──────────────────────────────────────────────────
        # 用法：video_gen <mode> [output] [key=value ...]
        # mode: slideshow / text_video / tts_video / screen_record
        # 例：video_gen text_video "" text="你好世界" duration=5
        "ai_video":          lambda: ai_video(
            " ".join(args[:1]) if args else "",
            args[1] if len(args) > 1 else "replicate",
            args[2] if len(args) > 2 else "",
            args[3] if len(args) > 3 else "",
            float(args[4]) if len(args) > 4 else 5,
            args[5] if len(args) > 5 else ""
        ),
        "video_gen":         lambda: video_gen(
            args[0] if args else "text_video",
            args[1] if len(args) > 1 else "",
            **dict(
                kv.split("=", 1) for kv in args[2:]
                if "=" in kv
            )
        ),
        # 缺口1
        "email_trigger":     lambda: email_trigger(args[0] if args else "check", *args[1:]),
        "file_trigger":      lambda: file_trigger(args[0] if args else "", args[1] if len(args)>1 else "any", args[2] if len(args)>2 else "notify"),
        "webhook_server":    lambda: webhook_server(args[0] if args else "status", *args[1:]),
        # 缺口2
        "com_auto":          lambda: com_auto(args[0] if args else "excel", args[1] if len(args)>1 else "list_sheets", *args[2:]),
        "dialog_auto":       lambda: dialog_auto(args[0] if args else "list_dialogs", *args[1:]),
        "ime_switch":        lambda: ime_switch(args[0] if args else "status"),
        # 缺口3
        "wake_word":         lambda: wake_word(args[0] if args else "listen_once", *args[1:]),
        "sound_detect":      lambda: sound_detect(args[0] if args else "volume_level", *args[1:]),
        "face_recognize":    lambda: face_recognize(args[0] if args else "detect", *args[1:]),
        # 缺口4
        "http_server":       lambda: http_server(args[0] if args else "status", *args[1:]),
        "lan_scan":          lambda: lan_scan(args[0] if args else "get_local_ip", *args[1:]),
        "serial_port":       lambda: serial_port(args[0] if args else "list", *args[1:]),
        "mqtt":              lambda: mqtt(args[0] if args else "test_connect", args[1] if len(args)>1 else "localhost", *args[2:]),
        # 缺口5
        "doc_ai":            lambda: doc_ai(args[0] if args else "summarize", *args[1:]),
        "web_monitor":       lambda: web_monitor(args[0] if args else "check_once", args[1] if len(args)>1 else "", *args[2:]),
        "audio_transcribe":  lambda: audio_transcribe(args[0] if args else "transcribe_mic", *args[1:]),
        # 投資技能
        "get_institutional":     lambda: get_institutional(args[0] if args else "", args[1] if len(args)>1 else ""),
        "get_sector":            lambda: get_sector(args[0] if args else "us"),
        "get_commodity":         lambda: get_commodity(args[0] if args else "all"),
        "get_bond_yield":        lambda: get_bond_yield(),
        "get_dividend_calendar": lambda: get_dividend_calendar(args[0]),
        "stock_screener":        lambda: stock_screener(args[0], args[1] if len(args)>1 else "us"),
        "get_margin_trading":    lambda: get_margin_trading(args[0], args[1] if len(args)>1 else ""),
        "get_options":           lambda: get_options(args[0], args[1] if len(args)>1 else ""),
        "get_futures":           lambda: get_futures(args[0] if args else "all"),
        "get_ipo":               lambda: get_ipo(args[0] if args else 10),
        "backtest":              lambda: backtest(args[0], args[1] if len(args)>1 else "ma_cross", args[2] if len(args)>2 else "2y"),
        # 中國大陸技能
        "get_ashare":            lambda: get_ashare(args[0], args[1] if len(args)>1 else "1mo"),
        "get_cn_news":           lambda: get_cn_news(args[0] if args else "all", args[1] if len(args)>1 else 5),
        "china_search":          lambda: china_search(args[0], args[1] if len(args)>1 else "其他", args[2] if len(args)>2 else 6),
        # 進階投資技能
        "get_global_market":     lambda: get_global_market(),
        "get_economic_calendar": lambda: get_economic_calendar(args[0] if args else 10),
        "get_earnings_calendar": lambda: get_earnings_calendar(args[0] if args else 7),
        "get_analyst_ratings":   lambda: get_analyst_ratings(args[0]),
        "get_short_interest":    lambda: get_short_interest(args[0]),
        "get_correlation":       lambda: get_correlation(args, args[-1] if args[-1] in ["3mo","6mo","1y","2y"] else "1y"),
        "get_risk_metrics":      lambda: get_risk_metrics(args[0], args[1] if len(args)>1 else "1y"),
        "get_money_flow":        lambda: get_money_flow(args[0]),
        "get_concept_stocks":    lambda: get_concept_stocks(args[0]),
        "get_crypto_depth":      lambda: get_crypto_depth(args[0] if args else "bitcoin"),
        "drip_calculator":       lambda: drip_calculator(args[0], float(args[1]) if len(args)>1 else 100, int(args[2]) if len(args)>2 else 10, float(args[3]) if len(args)>3 else 0),
        "get_forex_chart":       lambda: get_forex_chart(args[0], args[1] if len(args)>1 else "3mo"),
        "get_warrant":           lambda: get_warrant(args[0]),
        "get_portfolio_risk":    lambda: get_portfolio_risk([{"symbol": s, "weight": 1/len(args)} for s in args], "1y"),
        # 理財規劃技能
        "retirement_calculator": lambda: retirement_calculator(*args),
        "loan_calculator":       lambda: loan_calculator(*args),
        "compound_calculator":   lambda: compound_calculator(*args),
        "asset_allocation":      lambda: asset_allocation(*args),
        "tw_tax_calculator":     lambda: tw_tax_calculator(*args),
        "currency_converter":    lambda: currency_converter(*args),
        "get_fund":              lambda: get_fund(args[0]),
        "get_reits":             lambda: get_reits(args[0]),
        "inflation_adjusted":    lambda: inflation_adjusted(*args),
        "defi_calculator":       lambda: defi_calculator(*args),
        "gold_calculator":       lambda: gold_calculator(*args),
        "forex_deposit":         lambda: forex_deposit(*args),
        "financial_health":      lambda: financial_health(*args),
        # 研究統整與觀點技能
        "deep_research":         lambda: deep_research(args[0], args[1] if len(args)>1 else "zh-tw", args[2] if len(args)>2 else 5),
        "fact_check":            lambda: fact_check(" ".join(args)),
        "timeline_events":       lambda: timeline_events(args[0], args[1] if len(args)>1 else "zh-tw"),
        "sentiment_scan":        lambda: sentiment_scan(args[0], args[1] if len(args)>1 else "zh-tw"),
        "compare_analysis":      lambda: compare_analysis(args[0], args[1] if len(args)>1 else ""),
        "pros_cons_analysis":    lambda: pros_cons_analysis(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "zh-tw"),
        "research_report":       lambda: research_report(args[0], args[1] if len(args)>1 else "一般研究", args[2] if len(args)>2 else "zh-tw"),
        "opinion_writer":        lambda: opinion_writer(args[0], args[1] if len(args)>1 else "中立", args[2] if len(args)>2 else "正式"),
        "trend_forecast":        lambda: trend_forecast(args[0], args[1] if len(args)>1 else "全部", args[2] if len(args)>2 else "zh-tw"),
        "debate_simulator":      lambda: debate_simulator(" ".join(args)),
        # 全域研究技能
        "academic_search":       lambda: academic_search(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "en"),
        "health_research":       lambda: health_research(" ".join(args)),
        "law_research":          lambda: law_research(args[0], args[1] if len(args)>1 else "台灣", args[2] if len(args)>2 else "zh-tw"),
        "person_research":       lambda: person_research(args[0], args[1] if len(args)>1 else ""),
        "company_research":      lambda: company_research(args[0]),
        "product_review":        lambda: product_review(args[0], args[1] if len(args)>1 else ""),
        "travel_research":       lambda: travel_research(args[0], args[1] if len(args)>1 else None, args[2] if len(args)>2 else ""),
        "job_market":            lambda: job_market(args[0], args[1] if len(args)>1 else "台灣"),
        "impact_analysis":       lambda: impact_analysis(args[0], args[1] if len(args)>1 else ""),
        "scenario_planning":     lambda: scenario_planning(args[0], args[1] if len(args)>1 else ""),
        "decision_helper":       lambda: decision_helper(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else ""),
        "devil_advocate":        lambda: devil_advocate(" ".join(args)),
        # 統整與觀點強化技能
        "summary_writer":        lambda: summary_writer(args[0], args[1] if len(args)>1 else 7),
        "key_insights":          lambda: key_insights(args[0], args[1] if len(args)>1 else 5),
        "bias_detector":         lambda: bias_detector(" ".join(args)),
        "second_opinion":        lambda: second_opinion(args[0], args[1] if len(args)>1 else ""),
        "brainstorm":            lambda: brainstorm(args[0], args[1] if len(args)>1 else 8, args[2] if len(args)>2 else "實用"),
        "benchmark_analysis":    lambda: benchmark_analysis(args[0], args[1] if len(args)>1 else ""),
        # 觀點表達技能
        "steel_man":             lambda: steel_man(args[0], args[1] if len(args)>1 else ""),
        "socratic_questioning":  lambda: socratic_questioning(args[0], args[1] if len(args)>1 else 5),
        "analogy_maker":         lambda: analogy_maker(args[0], args[1] if len(args)>1 else "一般大眾", args[2] if len(args)>2 else 3),
        "narrative_builder":     lambda: narrative_builder(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else ""),
        "critique_writer":       lambda: critique_writer(args[0], args[1] if len(args)>1 else "觀點"),
        "position_statement":    lambda: position_statement(args[0], args[1] if len(args)>1 else "支持"),
        # 桌面視覺控制技能
        "ocr_click":             lambda: ocr_click(args[0], args[1] if len(args)>1 else 1, args[2] if len(args)>2 else "click"),
        "vision_locate":         lambda: vision_locate(args[0], args[1] if len(args)>1 else 1, args[2] if len(args)>2 else "click"),
        "screen_workflow":       lambda: screen_workflow(args[0]),
        "app_navigator":         lambda: app_navigator(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", args[3] if len(args)>3 else 1),
        "wait_and_click":        lambda: wait_and_click(args[0], args[1] if len(args)>1 else 15, args[2] if len(args)>2 else 1),
        "drag_drop":             lambda: drag_drop(args[0] if args else None, args[1] if len(args)>1 else None, args[2] if len(args)>2 else None, args[3] if len(args)>3 else None),
        # ── 新增 116 個工具 ──────────────────────────────────
        "analyze_pdf":           lambda: analyze_pdf_tool(args[0], args[1] if len(args)>1 else 4000),
        "audio_process":         lambda: audio_process_tool(args[0], args[1], args[2] if len(args)>2 else "", args[3] if len(args)>3 else 0, args[4] if len(args)>4 else 0),
        "auto_skill":            lambda: auto_skill_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", " ".join(args[3:]) if len(args)>3 else ""),
        "auto_trade":            lambda: auto_trade_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else 0, args[3] if len(args)>3 else 0, args[4] if len(args)>4 else "market"),
        "automation":            lambda: automation_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", " ".join(args[3:]) if len(args)>3 else ""),
        "barcode":               lambda: barcode_tool(args[0] if args else ""),
        "bluetooth":             lambda: bluetooth_tool(args[0], args[1] if len(args)>1 else ""),
        "browser_advanced":      lambda: browser_advanced_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", args[3] if len(args)>3 else "", args[4] if len(args)>4 else 0, args[5] if len(args)>5 else 30.0, args[6] if len(args)>6 else ""),
        "browser_control":       lambda: browser_control_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", " ".join(args[3:]) if len(args)>3 else ""),
        "calendar":              lambda: calendar_tool(args[0], args[1] if len(args)>1 else 7, args[2] if len(args)>2 else "", args[3] if len(args)>3 else "", args[4] if len(args)>4 else "", " ".join(args[5:]) if len(args)>5 else ""),
        "clipboard":             lambda: clipboard_tool(args[0], " ".join(args[1:]) if len(args)>1 else ""),
        "clipboard_image":       lambda: clipboard_image_tool(args[0], args[1] if len(args)>1 else ""),
        "cloud_storage":         lambda: cloud_storage_tool(args[0], args[1], args[2] if len(args)>2 else "root"),
        "compare_stocks":        lambda: compare_stocks_tool(args[0], args[1] if len(args)>1 else "all"),
        "database":              lambda: database_tool(args[0], args[1], " ".join(args[2:]) if len(args)>2 else ""),
        "ddg_search":            lambda: ddg_search_tool(" ".join(args[:1] if len(args)==1 else [args[0]]), args[1] if len(args)>1 else "zh-tw", args[2] if len(args)>2 else 5),
        "desktop_control":       lambda: desktop_control_tool(args[0], args[1] if len(args)>1 else None, args[2] if len(args)>2 else None, args[3] if len(args)>3 else None, args[4] if len(args)>4 else None),
        "device_manager":        lambda: device_manager_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else ""),
        "disk_backup":           lambda: disk_backup_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else ""),
        "display":               lambda: display_tool(args[0], args[1] if len(args)>1 else None),
        "docker":                lambda: docker_tool(args[0], args[1] if len(args)>1 else ""),
        "document_control":      lambda: document_control_tool(args[0], args[1], " ".join(args[2:]) if len(args)>2 else ""),
        "download_file":         lambda: download_file_tool(args[0], args[1] if len(args)>1 else ""),
        "dropbox":               lambda: dropbox_tool(args[0], args[1], args[2], args[3] if len(args)>3 else ""),
        "email_control":         lambda: email_control_tool(args[0], args[1], args[2], args[3] if len(args)>3 else "INBOX", args[4] if len(args)>4 else 5),
        "emotion_detect":        lambda: emotion_detect_tool(args[0], " ".join(args[1:]) if len(args)>1 else "", ""),
        "encrypt_file":          lambda: encrypt_file_tool(args[0], args[1], args[2]),
        "env_var":               lambda: env_var_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", args[3] if len(args)>3 else "false"),
        "file_system":           lambda: file_system_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", " ".join(args[3:]) if len(args)>3 else ""),
        "file_tools":            lambda: file_tools_tool(args[0], args[1], args[2] if len(args)>2 else "", args[3] if len(args)>3 else "", args[4] if len(args)>4 else "", args[5] if len(args)>5 else ""),
        "file_transfer":         lambda: file_transfer_tool(args[0], args[1], args[2] if len(args)>2 else ""),
        "find_image_on_screen":  lambda: find_image_on_screen_tool(args[0], args[1] if len(args)>1 else 0.8),
        "generate_image":        lambda: generate_image_tool(args[0], args[1] if len(args)>1 else 512, args[2] if len(args)>2 else 512, " ".join(args[3:]) if len(args)>3 else ""),
        "get_candlestick_chart": lambda: get_candlestick_chart_tool(args[0], args[1] if len(args)>1 else "3mo"),
        "get_crypto":            lambda: get_crypto_tool(args[0], args[1] if len(args)>1 else "usd"),
        "get_earnings":          lambda: get_earnings_tool(args[0]),
        "get_etf":               lambda: get_etf_tool(args[0]),
        "get_finance_news":      lambda: get_finance_news_tool(args[0] if args else "all", args[1] if len(args)>1 else 5),
        "get_forex":             lambda: get_forex_tool(args[0]),
        "get_fundamentals":      lambda: get_fundamentals_tool(args[0]),
        "get_macro":             lambda: get_macro_tool(args[0]),
        "get_market_sentiment":  lambda: get_market_sentiment_tool(),
        "get_stock":             lambda: get_stock_tool(args[0], args[1] if len(args)>1 else "1mo"),
        "get_stock_advanced":    lambda: get_stock_advanced_tool(args[0], args[1] if len(args)>1 else "macd,bb,kd"),
        "get_weather":           lambda: get_weather_tool(args[0]),
        "git":                   lambda: git_tool(args[0], args[1] if len(args)>1 else ".", args[2] if len(args)>2 else "", args[3] if len(args)>3 else "master"),
        "goal_manager":          lambda: goal_manager_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", args[3] if len(args)>3 else "", args[4] if len(args)>4 else "normal"),
        "google_trends":         lambda: google_trends_tool(args[0], args[1] if len(args)>1 else "today 3-m", args[2] if len(args)>2 else "TW"),
        "hardware":              lambda: hardware_tool(),
        "image_edit":            lambda: image_edit_tool(args[0], args[1], *args[2:]),
        "image_tools":           lambda: image_tools_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else 75, args[3] if len(args)>3 else 0, args[4] if len(args)>4 else 0),
        "knowledge_base":        lambda: knowledge_base_tool(args[0], " ".join(args[1:]) if len(args)>1 else ""),
        "long_term_memory":      lambda: long_term_memory_tool(args[0], args[1] if len(args)>1 else "0", " ".join(args[2:]) if len(args)>2 else "", args[3] if len(args)>3 else ""),
        "lookup":                lambda: lookup_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else 1.0, args[3] if len(args)>3 else "USD", args[4] if len(args)>4 else "TWD"),
        "manage_schedule":       lambda: manage_schedule_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", args[3] if len(args)>3 else ""),
        "media":                 lambda: media_tool(args[0], args[1] if len(args)>1 else ""),
        "monitor_config":        lambda: monitor_config_tool(),
        "multi_deploy":          lambda: multi_deploy_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", args[3] if len(args)>3 else "", args[4] if len(args)>4 else "/tmp/niu_bot"),
        "multi_perspective":     lambda: multi_perspective_tool(" ".join(args), "zh-tw"),
        "network_config":        lambda: network_config_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", args[3] if len(args)>3 else "", args[4] if len(args)>4 else ""),
        "network_diag":          lambda: network_diag_tool(args[0], args[1], args[2] if len(args)>2 else "22,80,443,3306,3389,8080"),
        "news_monitor":          lambda: news_monitor_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else 300, args[3] if len(args)>3 else 3600),
        "news_search":           lambda: news_search_tool(" ".join(args[:1] if len(args)==1 else [args[0]]), args[1] if len(args)>1 else "zh-TW", args[2] if len(args)>2 else 6),
        "nlp":                   lambda: nlp_tool(args[0], " ".join(args[1:])),
        "osint_search":          lambda: osint_search_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", args[3] if len(args)>3 else 10),
        "password_mgr":          lambda: password_mgr_tool(args[0], args[1], args[2], args[3] if len(args)>3 else "", args[4] if len(args)>4 else ""),
        "pdf_edit":              lambda: pdf_edit_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", args[3] if len(args)>3 else "", " ".join(args[4:]) if len(args)>4 else ""),
        "pdf_image":             lambda: pdf_image_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else 150),
        "pentest":               lambda: pentest_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "1-1000", args[3] if len(args)>3 else 2),
        "portfolio":             lambda: portfolio_tool(args[0], args[1] if len(args)>1 else 0, args[2] if len(args)>2 else "", args[3] if len(args)>3 else 0, args[4] if len(args)>4 else 0),
        "power_control":         lambda: power_control_tool(args[0]),
        "pptx_control":          lambda: pptx_control_tool(args[0], args[1], " ".join(args[2:]) if len(args)>2 else ""),
        "proactive_alert":       lambda: proactive_alert_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", args[3] if len(args)>3 else "", args[4] if len(args)>4 else "", args[5] if len(args)>5 else 60),
        "ptt_search":            lambda: ptt_search_tool(args[0], args[1] if len(args)>1 else "Gossiping", args[2] if len(args)>2 else 5),
        "push_notify":           lambda: push_notify_tool(args[0], " ".join(args[1:-1]), args[-1]),
        "qr_code":               lambda: qr_code_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", args[3] if len(args)>3 else 30.0),
        "read_screen":           lambda: read_screen_tool(" ".join(args) if args else "描述螢幕上有什麼", args[-1] if args and args[-1].isdigit() else 1),
        "read_webpage":          lambda: read_webpage_tool(args[0], args[1] if len(args)>1 else 3000),
        "registry":              lambda: registry_tool(args[0], args[1], args[2] if len(args)>2 else "", " ".join(args[3:]) if len(args)>3 else ""),
        "reminder":              lambda: reminder_tool(args[0], " ".join(args[1:])),
        "report":                lambda: report_tool(args[0], args[1], args[2] if len(args)>2 else ""),
        "restore_point":         lambda: restore_point_tool(args[0], " ".join(args[1:]) if len(args)>1 else ""),
        "run_code":              lambda: run_code_tool(args[0], " ".join(args[1:])),
        "screen_vision":         lambda: screen_vision_tool(" ".join(args) if args else "請描述這個畫面上有什麼"),
        "scroll_at":             lambda: scroll_at_tool(args[0] if args else "down", args[1] if len(args)>1 else 3, args[2] if len(args)>2 else None, args[3] if len(args)>3 else None, args[4] if len(args)>4 else 1, " ".join(args[5:]) if len(args)>5 else ""),
        "self_benchmark":        lambda: self_benchmark_tool(args[0]),
        "send_email":            lambda: send_email_tool(args[0], args[1], " ".join(args[2:])),
        "send_voice":            lambda: send_voice_tool(" ".join(args[:-1]) if len(args)>1 else " ".join(args), args[-1] if len(args)>1 and "Neural" in args[-1] else "zh-CN-YunxiNeural"),
        "smart_home":            lambda: smart_home_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", args[3] if len(args)>3 else "", args[4] if len(args)>4 else ""),
        "software":              lambda: software_tool(args[0], " ".join(args[1:]) if len(args)>1 else ""),
        "ssh_sftp":              lambda: ssh_sftp_tool(args[0], args[1], args[2], args[3], " ".join(args[4:]) if len(args)>4 else ""),
        "startup":               lambda: startup_tool(args[0], args[1] if len(args)>1 else "", " ".join(args[2:]) if len(args)>2 else ""),
        "system_monitor":        lambda: system_monitor_tool(args[0], args[1] if len(args)>1 else ""),
        "system_tools":          lambda: system_tools_tool(args[0]),
        "think_as":              lambda: think_as_tool(args[0], " ".join(args[1:]) if len(args)>1 else ""),
        "threat_intel":          lambda: threat_intel_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else ""),
        "todo_list":             lambda: todo_list_tool(args[0], " ".join(args[1:]) if len(args)>1 else "", 0),
        "tts_advanced":          lambda: tts_advanced_tool(args[0], " ".join(args[1:]) if len(args)>1 else "", args[-1] if len(args)>1 and "Neural" in args[-1] else "zh-CN-YunxiNeural"),
        "user_account":          lambda: user_account_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else ""),
        "video_gif":             lambda: video_gif_tool(args[0], args[1] if len(args)>1 else 0, args[2] if len(args)>2 else 5.0, args[3] if len(args)>3 else "", args[4] if len(args)>4 else 10),
        "video_process":         lambda: video_process_tool(args[0], args[1], args[2] if len(args)>2 else 0, args[3] if len(args)>3 else 0, args[4] if len(args)>4 else 0, args[5] if len(args)>5 else ""),
        "virtual_desktop":       lambda: virtual_desktop_tool(args[0]),
        "voice_cmd":             lambda: voice_cmd_tool(args[0], args[1] if len(args)>1 else 300.0, args[2] if len(args)>2 else "zh-TW"),
        "voice_id":              lambda: voice_id_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", args[3] if len(args)>3 else 5),
        "volume":                lambda: volume_tool(args[0], args[1] if len(args)>1 else None),
        "vpn":                   lambda: vpn_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", args[3] if len(args)>3 else ""),
        "wait_seconds":          lambda: wait_seconds_tool(args[0]),
        "web_scrape":            lambda: web_scrape_tool(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "body", args[3] if len(args)>3 else 2.0, args[4] if len(args)>4 else "full"),
        "webpage_shot":          lambda: webpage_shot_tool(args[0], args[1], args[2] if len(args)>2 else "body", args[3] if len(args)>3 else 60.0, args[4] if len(args)>4 else 3600.0),
        "wikipedia_search":      lambda: wikipedia_search_tool(args[0], args[1] if len(args)>1 else "zh"),
        "win_notify_relay":      lambda: win_notify_relay_tool(args[0], args[1] if len(args)>1 else 3600.0, args[2] if len(args)>2 else ""),
        "window_control":        lambda: window_control_tool(args[0], " ".join(args[1:]) if len(args)>1 else ""),
        "window_manager":        lambda: window_manager_tool(args[0] if args else "list", " ".join(args[1:]) if len(args)>1 else ""),
        "windows_update":        lambda: windows_update_tool(args[0]),
        "workflow":              lambda: workflow_tool(args[0], args[1] if len(args)>1 else "", " ".join(args[2:]) if len(args)>2 else ""),
        "youtube_summary":       lambda: youtube_summary_tool(args[0]),
        "tg_auto_reply":         lambda: tg_auto_reply_tool(args[0] if args else "start", args[1] if len(args)>1 else "", args[2] if len(args)>2 else "30"),
    }

    if tool not in tools:
        print(f"未知工具：{tool}")
        print(__doc__)
        sys.exit(1)

    try:
        tools[tool]()
    except Exception as e:
        print(f"執行失敗：{e}")
        sys.exit(1)
