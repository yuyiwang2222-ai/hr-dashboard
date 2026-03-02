"""
設定郵件密碼
============
此腳本用於安全儲存您的郵件密碼
密碼會存在 Windows 認證管理員中，不會以明文儲存
"""

import keyring
import getpass

SERVICE_NAME = "df_hr_report_email"
EMAIL = "chiehyi@df-recycle.com.tw"

def setup_password():
    print("=" * 50)
    print("大豐環保人力報告 - 郵件密碼設定")
    print("=" * 50)
    print()
    print(f"郵件帳號：{EMAIL}")
    print()
    print("注意：如果您的帳戶有啟用雙重驗證 (MFA)，")
    print("請使用「應用程式密碼」而非一般密碼。")
    print()
    print("建立應用程式密碼的方式：")
    print("1. 前往 https://account.microsoft.com/security")
    print("2. 點選「進階安全性選項」")
    print("3. 在「應用程式密碼」區塊建立新密碼")
    print()
    
    password = getpass.getpass("請輸入密碼（輸入時不會顯示）：")
    
    if password:
        keyring.set_password(SERVICE_NAME, EMAIL, password)
        print()
        print("✅ 密碼已安全儲存在 Windows 認證管理員中")
        print()
        
        # 測試是否能讀取
        stored = keyring.get_password(SERVICE_NAME, EMAIL)
        if stored:
            print("✅ 驗證成功：密碼可正常讀取")
        else:
            print("❌ 驗證失敗：無法讀取密碼")
    else:
        print("❌ 未輸入密碼，設定取消")

if __name__ == "__main__":
    setup_password()
    input("\n按 Enter 鍵關閉...")
