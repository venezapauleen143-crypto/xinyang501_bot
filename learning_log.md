

## 2026-04-11 — Python 排程任務：APScheduler 進階用法

# 📓 小牛馬的學習筆記
**日期：** 今晚　**主題：** Python 排程任務：APScheduler 進階用法

---

## 🧠 重點知識

### 1️⃣ APScheduler 四大核心元件

| 元件 | 說明 |
|------|------|
| **Scheduler（排程器）** | 整個系統的核心，負責管理任務 |
| **Trigger（觸發器）** | 決定任務「何時」執行 |
| **Job Store（任務儲存）** | 決定任務「存在哪裡」（記憶體 / 資料庫） |
| **Executor（執行器）** | 決定任務「如何執行」（執行緒 / 程序） |

---

### 2️⃣ 三種 Trigger 類型比較

| 類型 | 適用情境 | 範例 |
|------|----------|------|
| `date` | 只執行一次的固定時間點 | 明天早上九點發通知 |
| `interval` | 固定間隔重複執行 | 每 5 分鐘檢查一次 |
| `cron` | 複雜的週期規則（類 Linux cron） | 每週一早上八點執行 |

---

### 3️⃣ 常用 Scheduler 類型

- **`BlockingScheduler`** — 主程式的唯一任務就是跑排程（會阻塞）
- **`BackgroundScheduler`** — 排程在背景執行，主程式可同時做其他事
- **`AsyncIOScheduler`** — 適合 `asyncio` 非同步程式架構
- **`GeventScheduler`** — 適合 Gevent 協程環境

---

### 4️⃣ 進階設定重點

- **`misfire_grace_time`**：若任務錯過執行時間，寬限幾秒內仍可補執行
- **`coalesce`**：若任務積壓多次，是否合併成一次執行（`True` = 合併）
- **`max_instances`**：同一任務最多允許幾個實例同時執行
- **`replace_existing`**：重新加入同 ID 任務時是否覆蓋

---

### 5️⃣ Job Store 持久化（避免重啟後任務消失）

- 預設存在 **記憶體**，程式重啟後任務消失
- 可改用 **SQLAlchemy**（SQLite / PostgreSQL）或 **MongoDB** 持久化
- 持久化後任務重啟仍存在，適合正式環境

---

## 💻 實用範例程式碼

### 範例一：基本三種 Trigger 示範

```python
from apscheduler.schedulers.blocking import BlockingScheduler
from datetime import datetime, timedelta

scheduler = BlockingScheduler(timezone="Asia/Taipei")

def say_hello(name):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] 哈囉，{name}！")

# ✅ date：只執行一次
scheduler.add_job(
    say_hello,
    trigger="date",
    run_date=datetime.now() + timedelta(seconds=5),
    args=["小牛馬"],
    id="job_date"
)

# ✅ interval：每 10 秒執行一次
scheduler.add_job(
    say_hello,
    trigger="interval",
    seconds=10,
    args=["定時任務"],
    id="job_interval"
)

# ✅ cron：每天早上 8:30 執行
scheduler.add_job(
    say_hello,
    trigger="cron",
    hour=8,
    minute=30,
    args=["早安任務"],
    id="job_cron"
)

print("排程啟動中...")
scheduler.start()
```

---

### 範例二：BackgroundScheduler + 動態新增／移除任務

```python
from apscheduler.schedulers.background import BackgroundScheduler
import time

scheduler = BackgroundScheduler(timezone="Asia/Taipei")

def fetch_data(source):
    print(f"📡 從 {source} 抓取資料...")

# 背景啟動
scheduler.start()

# 動態新增任務
scheduler.add_job(fetch_data, "interval", seconds=5, args=["API_A"], id="job_api_a")
print("✅ job_api_a 已加入")

time.sleep(12)  # 主程式繼續做其他事

# 動態暫停任務
scheduler.pause_job("job_api_a")
print("⏸️ job_api_a 已暫停")

time.sleep(5)

# 恢復任務
scheduler.resume_job("job_api_a")
print("▶️ job_api_a 已恢復")

time.sleep(5)

# 移除任務
scheduler.remove_job("job_api_a")
print("🗑️ job_api_a 已移除")

# 查詢所有任務
jobs = scheduler.get_jobs()
print(f"目前任務數：{len(jobs)}")

scheduler.shutdown()
```

---

### 範例三：進階設定（misfire / coalesce / max_instances）

```python
from apscheduler.

---

## 2026-04-11 — Python 排程任務：APScheduler 進階用法

# 📓 小牛馬的學習筆記
**日期：** 今晚
**主題：** Python 排程任務：APScheduler 進階用法

---

## 一、重點知識

### 🔧 APScheduler 核心架構
- **APScheduler**（Advanced Python Scheduler）是 Python 最成熟的排程框架之一
- 四大核心元件：
  - `Scheduler`（排程器）：整個系統的控制中心
  - `JobStore`（工作儲存）：儲存已排程的工作（預設為記憶體）
  - `Executor`（執行器）：決定工作如何被執行（thread / process）
  - `Trigger`（觸發器）：決定工作何時被觸發

---

### ⏰ 三種觸發器（Trigger）比較

| 觸發器 | 說明 | 適用場景 |
|--------|------|----------|
| `date` | 在指定時間點執行一次 | 一次性任務 |
| `interval` | 每隔固定時間執行 | 定期輪詢、心跳檢查 |
| `cron` | 類似 Linux cron 的彈性排程 | 每天固定時間、複雜週期 |

---

### 💾 JobStore 工作儲存後端

- `MemoryJobStore`（預設）：程式重啟後工作會消失
- `SQLAlchemyJobStore`：支援 SQLite / PostgreSQL / MySQL，可持久化
- `RedisJobStore`：適合高效能、分散式場景
- `MongoDBJobStore`：文件型資料庫儲存

---

### ⚙️ Executor 執行器類型

- `ThreadPoolExecutor`：適合 I/O 密集型任務（預設，max_workers 可設定）
- `ProcessPoolExecutor`：適合 CPU 密集型任務，使用多進程避免 GIL
- `AsyncIOExecutor`：配合 `AsyncIOScheduler` 使用非同步任務

---

### 🎛️ 排程器種類

| 排程器 | 說明 |
|--------|------|
| `BlockingScheduler` | 主執行緒阻塞，程式只做排程 |
| `BackgroundScheduler` | 背景執行緒，不影響主程式 |
| `AsyncIOScheduler` | 搭配 asyncio 事件迴圈 |
| `GeventScheduler` | 搭配 Gevent 協程 |

---

### 🛡️ 進階設定重點

- **misfire_grace_time**：工作錯過執行時間後，允許的補跑寬限秒數
- **coalesce**：若多次錯過，設為 `True` 只補跑一次（避免爆量執行）
- **max_instances**：同一工作最多同時執行幾個實例（防止重疊）
- **replace_existing**：加入同 ID 工作時是否覆蓋舊工作

---

## 二、實用範例程式碼

### 📦 安裝

```bash
pip install apscheduler
pip install sqlalchemy  # 若需要持久化
```

---

### 範例 1：基本三種觸發器示範

```python
from apscheduler.schedulers.blocking import BlockingScheduler
from datetime import datetime, timedelta

scheduler = BlockingScheduler()

# ── 1. date：指定時間執行一次 ──────────────────
def one_time_task():
    print(f"[date] 只執行這一次！{datetime.now()}")

scheduler.add_job(
    one_time_task,
    trigger='date',
    run_date=datetime.now() + timedelta(seconds=5),
    id='one_time'
)

# ── 2. interval：每 10 秒執行一次 ──────────────
def interval_task():
    print(f"[interval] 定期執行：{datetime.now()}")

scheduler.add_job(
    interval_task,
    trigger='interval',
    seconds=10,
    id='interval_job',
    max_instances=1,          # 最多 1 個同時執行
    coalesce=True             # 錯過多次只補跑一次
)

# ── 3. cron：每天早上 9:30 執行 ─────────────────
def daily_report():
    print(f"[cron] 早安！產生每日報表：{datetime.now()}")

scheduler.add_job(
    daily_report,
    trigger='cron',
    hour=9,
    minute=30,
    id='daily_report',
    misfire_grace_time=60     # 60 秒內錯過仍可補跑
)

print("排程器啟動，按 Ctrl+C 停止")
try:
    scheduler.start()
except KeyboardInterrupt:
    print("排程器已停止")
```

---

### 範例 2：BackgroundScheduler + 動態新增/移除工作

```python
from apscheduler.schedulers.background import BackgroundScheduler
import time

scheduler = BackgroundScheduler()
scheduler.start()

def greet(name: str):
    print(f"👋 Hello, {name}！時間：{time.strftime('%H:%M:%S')}")

# 動態新增工作
job = scheduler.add_job(

---

## 2026-04-11 — Python 排程任務：APScheduler 進階用法

# 📓 小牛馬的學習筆記
**日期：** 今晚
**主題：** Python 排程任務：APScheduler 進階用法

---

## 🔑 重點知識

### 1. APScheduler 核心架構
- **Scheduler（排程器）**：整個系統的控制中心，常用的有：
  - `BlockingScheduler`：阻塞主程式執行，適合排程是唯一任務的情境
  - `BackgroundScheduler`：背景執行，不阻塞主程式，最常用
  - `AsyncIOScheduler`：搭配 `asyncio` 使用，適合非同步應用
- **Trigger（觸發器）**：決定任務「何時」執行，三種類型：
  - `date`：只執行一次，指定特定時間
  - `interval`：固定間隔重複執行
  - `cron`：仿 Linux crontab，彈性排程
- **JobStore（任務儲存）**：任務持久化的地方，預設存於記憶體，可改用 SQLite / Redis 等
- **Executor（執行器）**：負責實際執行任務的元件，預設是 `ThreadPoolExecutor`

---

### 2. 三種 Trigger 詳解

| Trigger | 說明 | 適用場景 |
|--------|------|---------|
| `date` | 指定單一時間點執行 | 一次性延遲任務 |
| `interval` | 每隔固定時間執行 | 心跳檢測、定期同步 |
| `cron` | 彈性時間規則（星期、月份等） | 每日報表、每週備份 |

---

### 3. 進階功能重點
- **`misfire_grace_time`**：任務錯過預定時間後，允許補執行的寬限秒數
- **`max_instances`**：同一任務允許同時執行的最大實例數，防止任務堆疊
- **`coalesce`**：若任務積壓多次未執行，設為 `True` 則只補執行一次
- **任務持久化**：搭配 `SQLAlchemyJobStore` 可讓任務在程式重啟後依然存在
- **事件監聽（Event Listener）**：可監聽任務成功、失敗、錯過等事件，實作告警機制

---

## 💻 實用範例程式碼

### 範例一：BackgroundScheduler 基本使用
```python
from apscheduler.schedulers.background import BackgroundScheduler
import time

def send_report():
    print("📊 產出每日報表中...")

scheduler = BackgroundScheduler()

# interval：每 30 秒執行一次
scheduler.add_job(send_report, 'interval', seconds=30)

scheduler.start()

try:
    while True:
        time.sleep(1)  # 主程式繼續跑，排程在背景執行
except KeyboardInterrupt:
    scheduler.shutdown()
    print("排程已關閉")
```

---

### 範例二：三種 Trigger 綜合示範
```python
from apscheduler.schedulers.blocking import BlockingScheduler
from datetime import datetime

scheduler = BlockingScheduler()

def job_once():
    print("🎯 這個任務只執行一次")

def job_interval():
    print(f"⏱️ 固定間隔執行：{datetime.now().strftime('%H:%M:%S')}")

def job_cron():
    print("🗓️ 每天早上 9 點執行晨間摘要")

# date：10 秒後執行一次
scheduler.add_job(job_once, 'date',
                  run_date=datetime.now().replace(second=datetime.now().second + 10))

# interval：每 5 秒執行一次
scheduler.add_job(job_interval, 'interval', seconds=5)

# cron：每天 09:00 執行（週一到週五）
scheduler.add_job(job_cron, 'cron',
                  day_of_week='mon-fri',
                  hour=9,
                  minute=0)

scheduler.start()
```

---

### 範例三：misfire_grace_time + max_instances 防護
```python
from apscheduler.schedulers.background import BackgroundScheduler
import time

def heavy_task():
    print("🔄 執行耗時任務...")
    time.sleep(10)  # 模擬耗時作業

scheduler = BackgroundScheduler()

scheduler.add_job(
    heavy_task,
    'interval',
    seconds=5,
    max_instances=1,        # 同時只允許 1 個實例，避免任務堆疊
    coalesce=True,          # 積壓的多次觸發只補執行一次
    misfire_grace_time=30   # 錯過 30 秒內仍可補執行
)

scheduler.start()
```

---

### 範例四：SQLAlchemy 持久化任務儲存
```python
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.jobstores.sqlalchemy import SQLAlchemyJobStore
from apscheduler.executors.pool import ThreadPoolExecutor

# 設定 JobStore 與 Executor
jobstores = {
    'default': SQLAlchemyJobStore(url

---