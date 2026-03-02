"""
大豐環保人力儀表板 - 初始化資料設定
執行此腳本將建立 Excel 範本檔案與測試資料
"""

import pandas as pd
from datetime import datetime, timedelta
import random
import os

# 確保 data 目錄存在
os.makedirs('data/snapshots', exist_ok=True)

# ============================================
# 1. 建立部門對照表
# ============================================
departments_data = {
    'code': ['GM', 'BD', 'OPS', 'REC', 'INT', 'EC', 'MKT', 'RD', 'FIN', 'HR', 'IT'],
    'name': ['總經理室', '業務開發部', '營運管理部', '再生處理部', '國際事業部', 
             '電子商務部', '行銷部', '研發部', '財務部', '人資部', '資訊部']
}
df_departments = pd.DataFrame(departments_data)
df_departments.to_excel('data/departments.xlsx', index=False, sheet_name='部門')
print("✅ 已建立 data/departments.xlsx")

# ============================================
# 2. 建立員工主檔（含測試資料）
# ============================================

# 定義職務清單（依部門）
positions_by_dept = {
    'GM': ['總經理', '特助', '秘書'],
    'BD': ['業務經理', '業務專員', '業務助理'],
    'OPS': ['營運經理', '營運專員', '調度員', '司機'],
    'REC': ['廠長', '課長', '組長', '作業員', '技術員'],
    'INT': ['國際業務經理', '國際業務專員', '貿易專員'],
    'EC': ['電商經理', '電商專員', '客服專員'],
    'MKT': ['行銷經理', '行銷專員', '設計師'],
    'RD': ['研發經理', '資深工程師', '工程師', '助理工程師'],
    'FIN': ['財務經理', '會計', '出納'],
    'HR': ['人資經理', '人資專員', '招募專員'],
    'IT': ['資訊經理', '系統工程師', '網管工程師', 'MIS專員']
}

# 定義僱用類型權重（正職居多）
employment_types = ['正職', '約聘', '實習']
employment_weights = [0.85, 0.1, 0.05]

# 定義直間接（依部門）
indirect_depts = ['GM', 'FIN', 'HR', 'IT', 'MKT']  # 間接部門
direct_depts = ['BD', 'OPS', 'REC', 'INT', 'EC', 'RD']  # 直接部門

# 名字範本
first_names = ['志明', '淑芬', '俊傑', '雅婷', '建宏', '怡君', '家豪', '佳穎', 
               '宗翰', '詩涵', '冠宇', '雅琪', '柏翰', '宜蓁', '彥廷', '欣怡',
               '承恩', '婉婷', '宇軒', '筱涵', '柏均', '雅雯', '冠廷', '詩婷']
last_names = ['陳', '林', '黃', '張', '李', '王', '吳', '劉', '蔡', '楊', 
              '許', '鄭', '謝', '郭', '洪', '邱', '曾', '廖', '賴', '周']

# 生成員工資料
employees = []
emp_id = 1

# 各部門人數配置
dept_headcount = {
    'GM': 5,
    'BD': 25,
    'OPS': 40,
    'REC': 120,
    'INT': 15,
    'EC': 20,
    'MKT': 15,
    'RD': 30,
    'FIN': 10,
    'HR': 8,
    'IT': 12
}

today = datetime.now().date()

for dept_code, count in dept_headcount.items():
    positions = positions_by_dept[dept_code]
    
    for i in range(count):
        # 員工編號
        employee_id = f"TF{emp_id:04d}"
        emp_id += 1
        
        # 姓名
        name = random.choice(last_names) + random.choice(first_names)
        
        # 職務（依部門人數分配）
        if i == 0:
            position = positions[0]  # 第一位是主管
        else:
            position = random.choice(positions[1:]) if len(positions) > 1 else positions[0]
        
        # 僱用類型
        employment_type = random.choices(employment_types, employment_weights)[0]
        
        # 直間接
        labor_type = '間接' if dept_code in indirect_depts else '直接'
        
        # 到職日（過去1-5年內）
        days_ago = random.randint(30, 1825)  # 30天~5年
        hire_date = today - timedelta(days=days_ago)
        
        # 離職日（大部分在職，約5%離職）
        if random.random() < 0.05:
            # 離職員工
            leave_days_ago = random.randint(1, 60)  # 最近60天內離職
            leave_date = today - timedelta(days=leave_days_ago)
            status = '離職'
        else:
            leave_date = None
            status = '在職'
        
        employees.append({
            'employee_id': employee_id,
            'name': name,
            'department': dept_code,
            'position': position,
            'employment_type': employment_type,
            'labor_type': labor_type,
            'hire_date': hire_date,
            'leave_date': leave_date,
            'status': status
        })

# 加入一些近期入職的員工（本週/本月）
for i in range(8):
    dept_code = random.choice(list(dept_headcount.keys()))
    positions = positions_by_dept[dept_code]
    
    employee_id = f"TF{emp_id:04d}"
    emp_id += 1
    
    employees.append({
        'employee_id': employee_id,
        'name': random.choice(last_names) + random.choice(first_names),
        'department': dept_code,
        'position': random.choice(positions[1:]) if len(positions) > 1 else positions[0],
        'employment_type': '正職',
        'labor_type': '間接' if dept_code in indirect_depts else '直接',
        'hire_date': today - timedelta(days=random.randint(0, 14)),  # 最近兩週入職
        'leave_date': None,
        'status': '在職'
    })

# 加入一些近期離職的員工
for i in range(3):
    dept_code = random.choice(list(dept_headcount.keys()))
    positions = positions_by_dept[dept_code]
    
    employee_id = f"TF{emp_id:04d}"
    emp_id += 1
    
    hire_date = today - timedelta(days=random.randint(180, 730))
    leave_date = today - timedelta(days=random.randint(0, 7))  # 本週離職
    
    employees.append({
        'employee_id': employee_id,
        'name': random.choice(last_names) + random.choice(first_names),
        'department': dept_code,
        'position': random.choice(positions[1:]) if len(positions) > 1 else positions[0],
        'employment_type': '正職',
        'labor_type': '間接' if dept_code in indirect_depts else '直接',
        'hire_date': hire_date,
        'leave_date': leave_date,
        'status': '離職'
    })

# 建立 DataFrame 並儲存
df_employees = pd.DataFrame(employees)
df_employees.to_excel('data/employees.xlsx', index=False, sheet_name='員工')
print("✅ 已建立 data/employees.xlsx")

# ============================================
# 3. 顯示統計摘要
# ============================================
print("\n📊 資料統計摘要:")
print(f"  - 總員工數：{len(df_employees)}")
print(f"  - 在職人數：{len(df_employees[df_employees['status'] == '在職'])}")
print(f"  - 離職人數：{len(df_employees[df_employees['status'] == '離職'])}")
print(f"\n  部門分布：")
for dept in df_departments['code']:
    count = len(df_employees[(df_employees['department'] == dept) & (df_employees['status'] == '在職')])
    name = df_departments[df_departments['code'] == dept]['name'].values[0]
    print(f"    {name}：{count} 人")

print("\n✅ 資料初始化完成！")
print("   請執行 'streamlit run app.py' 啟動儀表板")
