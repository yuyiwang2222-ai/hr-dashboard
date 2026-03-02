# 部署指南：Streamlit Cloud + Cloudflare R2

## 架構說明

```
┌─────────────────┐    上傳PDF    ┌─────────────────┐
│   本機電腦      │ ────────────> │  Cloudflare R2  │
│  auto_report.py │               │   (物件儲存)    │
└─────────────────┘               └────────┬────────┘
                                           │
                                           │ 讀取PDF
                                           ▼
                                  ┌─────────────────┐
                                  │ Streamlit Cloud │
                                  │   (儀表板)      │
                                  └─────────────────┘
                                           │
                                           │ 瀏覽
                                           ▼
                                  ┌─────────────────┐
                                  │   使用者瀏覽器  │
                                  └─────────────────┘
```

---

## 步驟一：設定 Cloudflare R2

### 1.1 建立 R2 Bucket

1. 登入 [Cloudflare Dashboard](https://dash.cloudflare.com)
2. 左側選單 > **R2 Object Storage**
3. 點擊 **Create bucket**
4. 設定：
   - Bucket name: `hr-reports`
   - Location: 選擇最近的區域
5. 點擊 **Create bucket**

### 1.2 建立 API Token

1. 在 R2 頁面，點擊 **Manage R2 API Tokens**
2. 點擊 **Create API Token**
3. 設定：
   - Token name: `hr-reports-upload`
   - Permissions: **Object Read & Write**
   - Specify bucket(s): 選擇 `hr-reports`
4. 點擊 **Create API Token**
5. **⚠️ 重要**：複製並保存以下資訊：
   - Access Key ID
   - Secret Access Key
   - Account ID（在 Dashboard 首頁右側）

### 1.3 設定公開存取（選填）

如果要讓 PDF 可以直接透過網址存取：

1. 進入 `hr-reports` bucket
2. **Settings** > **Public Access**
3. 啟用 **R2.dev subdomain**
4. 記下公開網域：`hr-reports.xxxx.r2.dev`

---

## 步驟二：設定本機環境

### 2.1 建立 .env 檔案

複製範本並填入實際值：

```bash
cp .env.example .env
```

編輯 `.env`：

```env
R2_ACCOUNT_ID=你的帳號ID
R2_ACCESS_KEY_ID=你的Access Key
R2_SECRET_ACCESS_KEY=你的Secret Key
R2_BUCKET_NAME=hr-reports
R2_PUBLIC_DOMAIN=hr-reports.xxxx.r2.dev
```

### 2.2 安裝依賴

```bash
pip install boto3 python-dotenv
```

### 2.3 測試上傳

```bash
# 上傳最新報告
python upload_to_r2.py

# 列出已上傳的報告
python upload_to_r2.py --list
```

---

## 步驟三：整合到自動報告流程

修改 `auto_report.py`，在報告產生後自動上傳：

```python
# 在 auto_report.py 的最後加入：
from upload_to_r2 import upload_file, get_latest_report_path

# 步驟 4：上傳到 R2
print("[步驟 4/4] 上傳到 Cloudflare R2...")
report_path = get_latest_report_path()
if report_path:
    upload_file(report_path)
```

或使用批次檔：

```bat
@echo off
python auto_report.py
python upload_to_r2.py
pause
```

---

## 步驟四：部署 Streamlit Cloud

### 4.1 準備 GitHub Repository

1. 在 GitHub 建立新的 repository
2. 上傳以下檔案：
   - `app.py`
   - `requirements.txt`
   - `config.yaml`
   - `src/` 資料夾
   - `data/` 資料夾（不含敏感資料）

### 4.2 部署到 Streamlit Cloud

1. 前往 [Streamlit Cloud](https://share.streamlit.io)
2. 使用 GitHub 帳號登入
3. 點擊 **New app**
4. 選擇你的 repository
5. 設定：
   - Branch: `main`
   - Main file path: `app.py`
6. 點擊 **Deploy**

### 4.3 設定 Secrets

1. 進入已部署的 App
2. 右上角 **⋮** > **Settings**
3. 左側選單 **Secrets**
4. 填入：

```toml
[r2]
account_id = "你的帳號ID"
access_key_id = "你的Access Key"
secret_access_key = "你的Secret Key"
bucket_name = "hr-reports"
public_domain = "hr-reports.xxxx.r2.dev"
```

5. 點擊 **Save**

---

## 步驟五：修改 Streamlit 讀取 R2 報告

在 `app.py` 中新增報告下載功能：

```python
import streamlit as st
import boto3
from botocore.config import Config

def get_r2_client():
    """建立 R2 客戶端"""
    try:
        r2_config = st.secrets["r2"]
        endpoint_url = f"https://{r2_config['account_id']}.r2.cloudflarestorage.com"
        
        return boto3.client(
            's3',
            endpoint_url=endpoint_url,
            aws_access_key_id=r2_config['access_key_id'],
            aws_secret_access_key=r2_config['secret_access_key'],
            config=Config(signature_version='s3v4')
        )
    except Exception as e:
        st.error(f"R2 連線失敗：{e}")
        return None

def list_r2_reports():
    """列出 R2 中的報告"""
    client = get_r2_client()
    if not client:
        return []
    
    try:
        response = client.list_objects_v2(
            Bucket=st.secrets["r2"]["bucket_name"],
            Prefix="reports/"
        )
        return [obj['Key'] for obj in response.get('Contents', [])]
    except Exception as e:
        st.error(f"列出報告失敗：{e}")
        return []
```

---

## 完成！

部署完成後，您的工作流程將變成：

1. **本機執行** `auto_report.py` 產生 PDF
2. **自動上傳** 到 Cloudflare R2
3. **Streamlit Cloud** 儀表板可存取報告
4. **任何人** 都可以透過網址查看儀表板

---

## 常見問題

### Q: 上傳失敗，顯示 Access Denied
- 檢查 API Token 權限是否為 "Object Read & Write"
- 確認 Token 有綁定正確的 bucket

### Q: Streamlit Cloud 無法讀取 R2
- 確認 Secrets 設定正確
- 檢查 bucket 名稱是否一致

### Q: 想要自訂網域
1. 在 Cloudflare 設定 Custom Domain
2. 更新 `R2_PUBLIC_DOMAIN` 環境變數

---

## 相關資源

- [Cloudflare R2 文件](https://developers.cloudflare.com/r2/)
- [Streamlit Cloud 文件](https://docs.streamlit.io/streamlit-community-cloud)
- [boto3 S3 文件](https://boto3.amazonaws.com/v1/documentation/api/latest/reference/services/s3.html)
