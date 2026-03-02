# -*- coding: utf-8 -*-
"""
PDF 上傳到 Cloudflare R2 腳本
============================================
功能：
1. 將本機產生的 PDF 報告上傳到 Cloudflare R2
2. 可整合到 auto_report.py 自動執行
3. 支援手動執行

使用方式：
    python upload_to_r2.py                    # 上傳最新報告
    python upload_to_r2.py --file report.pdf  # 上傳指定檔案

環境變數設定（或在 .env 檔案中）：
    R2_ACCOUNT_ID=你的帳號ID
    R2_ACCESS_KEY_ID=你的 Access Key
    R2_SECRET_ACCESS_KEY=你的 Secret Key
    R2_BUCKET_NAME=你的 Bucket 名稱
"""

import os
import sys
import glob
import argparse
from datetime import datetime
from pathlib import Path

try:
    import boto3
    from botocore.config import Config
except ImportError:
    print("❌ 需要安裝 boto3：pip install boto3")
    sys.exit(1)

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # 如果沒有 python-dotenv，直接使用環境變數


def get_r2_client():
    """建立 R2 客戶端（使用 S3 相容 API）"""
    account_id = os.getenv('R2_ACCOUNT_ID')
    access_key = os.getenv('R2_ACCESS_KEY_ID')
    secret_key = os.getenv('R2_SECRET_ACCESS_KEY')
    
    if not all([account_id, access_key, secret_key]):
        print("❌ 缺少 R2 環境變數設定")
        print("   請設定：R2_ACCOUNT_ID, R2_ACCESS_KEY_ID, R2_SECRET_ACCESS_KEY")
        return None
    
    # R2 端點
    endpoint_url = f"https://{account_id}.r2.cloudflarestorage.com"
    
    # 建立 S3 客戶端（R2 相容 S3 API）
    client = boto3.client(
        's3',
        endpoint_url=endpoint_url,
        aws_access_key_id=access_key,
        aws_secret_access_key=secret_key,
        config=Config(
            signature_version='s3v4',
            retries={'max_attempts': 3}
        )
    )
    
    return client


def upload_file(file_path: str, bucket_name: str = None, object_key: str = None) -> bool:
    """
    上傳檔案到 R2
    
    Args:
        file_path: 本機檔案路徑
        bucket_name: R2 bucket 名稱（預設從環境變數讀取）
        object_key: R2 中的檔案路徑（預設使用原檔名）
    
    Returns:
        是否上傳成功
    """
    client = get_r2_client()
    if not client:
        return False
    
    bucket_name = bucket_name or os.getenv('R2_BUCKET_NAME', 'hr-reports')
    
    # 確認檔案存在
    if not os.path.exists(file_path):
        print(f"❌ 找不到檔案：{file_path}")
        return False
    
    # 設定 object key（R2 中的路徑）
    if not object_key:
        filename = os.path.basename(file_path)
        # 按年月分類儲存：reports/2026/02/人力分析報告_20260224.pdf
        now = datetime.now()
        object_key = f"reports/{now.year}/{now.month:02d}/{filename}"
    
    try:
        print(f"📤 上傳中：{file_path}")
        print(f"   目標：s3://{bucket_name}/{object_key}")
        
        # 設定 Content-Type
        content_type = 'application/pdf' if file_path.endswith('.pdf') else 'application/octet-stream'
        
        client.upload_file(
            file_path,
            bucket_name,
            object_key,
            ExtraArgs={
                'ContentType': content_type,
                'Metadata': {
                    'uploaded-at': datetime.now().isoformat(),
                    'source': 'auto-report'
                }
            }
        )
        
        print(f"✅ 上傳成功：{object_key}")
        
        # 產生公開 URL（如果 bucket 有設定公開存取）
        public_url = f"https://{os.getenv('R2_PUBLIC_DOMAIN', 'your-domain.r2.dev')}/{object_key}"
        print(f"   公開網址：{public_url}")
        
        return True
        
    except Exception as e:
        print(f"❌ 上傳失敗：{e}")
        return False


def list_reports(bucket_name: str = None, prefix: str = "reports/") -> list:
    """
    列出 R2 中的報告列表
    
    Args:
        bucket_name: R2 bucket 名稱
        prefix: 路徑前綴
    
    Returns:
        報告列表
    """
    client = get_r2_client()
    if not client:
        return []
    
    bucket_name = bucket_name or os.getenv('R2_BUCKET_NAME', 'hr-reports')
    
    try:
        response = client.list_objects_v2(Bucket=bucket_name, Prefix=prefix)
        
        reports = []
        for obj in response.get('Contents', []):
            reports.append({
                'key': obj['Key'],
                'size': obj['Size'],
                'last_modified': obj['LastModified']
            })
        
        return sorted(reports, key=lambda x: x['last_modified'], reverse=True)
        
    except Exception as e:
        print(f"❌ 列出報告失敗：{e}")
        return []


def download_file(object_key: str, local_path: str, bucket_name: str = None) -> bool:
    """
    從 R2 下載檔案
    
    Args:
        object_key: R2 中的檔案路徑
        local_path: 本機儲存路徑
        bucket_name: R2 bucket 名稱
    
    Returns:
        是否下載成功
    """
    client = get_r2_client()
    if not client:
        return False
    
    bucket_name = bucket_name or os.getenv('R2_BUCKET_NAME', 'hr-reports')
    
    try:
        print(f"📥 下載中：{object_key}")
        
        # 確保目錄存在
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        
        client.download_file(bucket_name, object_key, local_path)
        
        print(f"✅ 下載成功：{local_path}")
        return True
        
    except Exception as e:
        print(f"❌ 下載失敗：{e}")
        return False


def get_latest_report_path() -> str:
    """取得最新的 PDF 報告路徑"""
    script_dir = Path(__file__).parent
    report_dir = script_dir / "報告"
    
    # 尋找最新的 PDF 檔案
    pdf_files = list(report_dir.glob("人力分析報告_*.pdf"))
    
    if not pdf_files:
        print("❌ 找不到任何 PDF 報告")
        return None
    
    # 按檔名排序（日期格式：YYYYMMDD）
    latest = sorted(pdf_files, key=lambda x: x.name, reverse=True)[0]
    return str(latest)


def main():
    parser = argparse.ArgumentParser(description='上傳 PDF 報告到 Cloudflare R2')
    parser.add_argument('--file', '-f', help='指定要上傳的檔案路徑')
    parser.add_argument('--list', '-l', action='store_true', help='列出 R2 中的報告')
    parser.add_argument('--download', '-d', help='下載指定的報告（object key）')
    parser.add_argument('--output', '-o', help='下載時的本機路徑')
    
    args = parser.parse_args()
    
    # 列出報告
    if args.list:
        print("📋 R2 報告列表：")
        reports = list_reports()
        for r in reports:
            size_kb = r['size'] / 1024
            print(f"   {r['key']} ({size_kb:.1f} KB) - {r['last_modified']}")
        return
    
    # 下載報告
    if args.download:
        output_path = args.output or os.path.basename(args.download)
        download_file(args.download, output_path)
        return
    
    # 上傳報告
    file_path = args.file
    if not file_path:
        file_path = get_latest_report_path()
    
    if file_path:
        upload_file(file_path)
    else:
        print("使用方式：")
        print("  python upload_to_r2.py                    # 上傳最新報告")
        print("  python upload_to_r2.py --file report.pdf  # 上傳指定檔案")
        print("  python upload_to_r2.py --list             # 列出報告")


if __name__ == '__main__':
    main()
