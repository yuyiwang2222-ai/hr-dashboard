"""
檢查幕僚部門代主管
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.data_loader import load_config, load_employees
import pandas as pd

config = load_config()
df = load_employees(config=config)

# 找出五個幕僚部門的代主管
staff_depts = ['人資部', '資訊部', '電子商務部', '行銷部', '財務部']
active = df[df['status'] == '在職']
deputy_managers = active[
    (active['job_title_original'] == '代主管') & 
    (active['department'].isin(staff_depts))
]

print('=== 幕僚部門代主管檢查 ===')
print(f'找到 {len(deputy_managers)} 位')
if len(deputy_managers) > 0:
    print()
    for _, row in deputy_managers.iterrows():
        print(f"  部門: {row['department']}")
        print(f"  所屬職務: {row['job_title_original']}")
        print(f"  職務類別(position): {row['position']}")
        print()
