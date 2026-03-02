"""
驗證資料腳本 - 檢查員工人數和工讀生標記
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.data_loader import load_config, load_employees
import pandas as pd

config = load_config()
df = load_employees(config=config)

active = df[df['status'] == '在職']

print('=== 驗證人數 ===')
print(f'總在職人數: {len(active)}')

# 檢查工讀生標記
print()
print('=== 工讀生檢查 ===')
if 'is_part_time' in active.columns:
    part_time_true = active[active['is_part_time'] == True]
    part_time_false = active[active['is_part_time'] == False]
    print(f'is_part_time=True: {len(part_time_true)}')
    print(f'is_part_time=False: {len(part_time_false)}')
else:
    print('is_part_time 欄位不存在')

# 檢查原始員工狀態
print()
print('=== 原始員工狀態 ===')
if 'employee_status_original' in active.columns:
    status_counts = active.groupby('employee_status_original').size().sort_values(ascending=False)
    for status, count in status_counts.items():
        print(f'  {status}: {count}')

# 檢查事業部分布
print()
print('=== 事業部人數 ===')
dept_counts = active.groupby('department').size().sort_values(ascending=False)
for dept, count in dept_counts.items():
    print(f'  {dept}: {count}')
print(f'  總計: {dept_counts.sum()}')

# 檢查排除條件
print()
print('=== 排除檢查 ===')
# 負責人
responsible_count = 0
if 'department_original' in active.columns:
    responsible = active[active['department_original'].astype(str).str.contains('負責人', na=False)]
    responsible_count = len(responsible)
    print(f'含「負責人」: {responsible_count}')
    if responsible_count > 0:
        dept_list = responsible['department_original'].unique().tolist()
        print(f'  部門原名: {dept_list}')

# 空白/未分類
blank = active[active['department'].isna() | (active['department'].astype(str).str.strip() == '')]
unclassified = active[active['department'] == '未分類']
print(f'空白部門: {len(blank)}')
print(f'未分類: {len(unclassified)}')

# 再生紡織設計部
textile = active[active['department'] == '再生紡織設計部']
print(f'再生紡織設計部: {len(textile)}')

# 排除後應有人數
excluded_count = responsible_count + len(unclassified) + len(textile)
print()
print(f'應排除人數: 負責人{responsible_count} + 未分類{len(unclassified)} + 再生紡織{len(textile)} = {excluded_count}')
print(f'排除後人數: {len(active) - excluded_count}')
