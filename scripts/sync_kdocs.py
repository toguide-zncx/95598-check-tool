#!/usr/bin/env python3
"""
从金山文档同步 95598 智能座席助手项目数据，生成 plan-data.json
用法: python3 sync_kdocs.py
需要: kdocs-cli 已安装且 KINGSOFT_DOCS_TOKEN 环境变量已设置
"""
import json
import subprocess
import sys
import os
import re
from datetime import datetime
from collections import defaultdict, OrderedDict

FILE_ID = "dCYHGzpaJrMUMcY8UtM71xqMXeDjLSUsv"

# 部门顺序（用于保持输出顺序一致）
DEPT_ORDER = [
    "北中心客服一部", "北中心客服二部", "北中心客服三部", "北中心客服五部",
    "南中心客服一部", "南中心客服二部", "南中心客服三部", "南中心客服五部",
]

def kdocs_get_sheet(worksheet_id, row_to, col_to):
    """调用 kdocs-cli 获取 sheet 数据，返回 grid[row][col]"""
    # 支持通过 KDOCS_TOKEN 或 KINGSOFT_DOCS_TOKEN 环境变量传递 token
    token = os.environ.get("KDOCS_TOKEN") or os.environ.get("KINGSOFT_DOCS_TOKEN", "")
    if token:
        os.environ["KINGSOFT_DOCS_TOKEN"] = token
    cmd = [
        "kdocs-cli", "sheet", "get-range-data",
        json.dumps({
            "file_id": FILE_ID,
            "worksheet_id": worksheet_id,
            "range": {"rowFrom": 0, "rowTo": row_to, "colFrom": 0, "colTo": col_to}
        })
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"kdocs-cli error: {result.stderr}", file=sys.stderr)
        sys.exit(1)
    
    data = json.loads(result.stdout)
    cells = data['data']['detail']['rangeData']
    
    grid = {}
    for c in cells:
        r = c['originRow']
        col = c['originCol']
        val = c['cellText']
        if r not in grid:
            grid[r] = {}
        grid[r][col] = val
    
    return grid

def parse_date(date_str):
    """解析日期字符串，返回 'YYYY-M-D' 格式"""
    if not date_str:
        return ""
    # "2026年6月24日"
    m = re.match(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?', date_str)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    # "2026/7/3日"
    m = re.match(r'(\d{4})\s*/\s*(\d{1,2})\s*/\s*(\d{1,2})\s*日?', date_str)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    return date_str.strip()

def parse_sheet1(grid):
    """解析 Sheet1（任务计划表），返回 tasks 列表"""
    tasks = []
    
    for r in sorted(grid.keys()):
        if r <= 1:  # 跳过标题(row 0)和表头(row 1)
            continue
        row = grid[r]
        seq = row.get(0, '').strip()
        
        if not seq:
            continue
        
        # 阶段标题行（序号列包含中文数字如"一、"）
        if re.match(r'[一二三四五六七八九十]+、', seq):
            tasks.append({
                "id": seq,
                "start": "", "end": "", "category": "", "name": "",
                "detail": "", "deliverable": "", "owner": "",
                "priority": "", "status": ""
            })
            continue
        
        # 跳过非数字序号行（如样例行）
        if not seq.isdigit():
            continue
        
        task = {
            "id": seq,
            "start": parse_date(row.get(1, '')),
            "end": parse_date(row.get(2, '')),
            "category": row.get(3, '').strip(),
            "name": row.get(4, '').strip(),
            "detail": row.get(5, '').strip(),
            "deliverable": row.get(6, '').strip(),
            "owner": row.get(7, '').strip(),
            "priority": row.get(8, '').strip(),
            "status": row.get(9, '').strip()
        }
        tasks.append(task)
    
    # 自动重新编号：检测重复ID并按出现顺序重新分配唯一ID
    seen_ids = set()
    has_dup = False
    for t in tasks:
        if t['id'] and t['id'].isdigit():
            if t['id'] in seen_ids:
                has_dup = True
                break
            seen_ids.add(t['id'])
    
    if has_dup:
        counter = 1
        for t in tasks:
            if t['id'] and t['id'].isdigit():
                t['id'] = str(counter)
                counter += 1
        print(f"  ⚠️ 检测到重复ID，已自动重新编号为 1-{counter-1}")
    
    return tasks

def summarize_notes(notes_list):
    """将多条备注精简为关键信息摘要"""
    if not notes_list:
        return ''
    if len(notes_list) <= 2:
        return '；'.join(notes_list)
    
    # 提取关键问题类型
    key_notes = []
    has_slow = any('卡顿' in n or '慢' in n for n in notes_list)
    has_no_content = any('管理端无通话内容' in n for n in notes_list)
    has_unbind = any('未解绑' in n or '没有解绑' in n for n in notes_list)
    has_normal = any('转写正常' in n or '已完成' in n for n in notes_list)
    has_no_sound = any('没声音' in n or '没有显示来电' in n for n in notes_list)
    
    if has_slow:
        key_notes.append('部分电脑卡顿转写慢')
    if has_no_content:
        key_notes.append('管理端无通话内容')
    if has_unbind or has_no_sound:
        key_notes.append('1台硬话机未解绑')
    if not key_notes and has_normal:
        key_notes.append('已完成，转写正常')
    
    return '，'.join(key_notes) if key_notes else ''

def parse_sheet4(grid):
    """解析 Sheet4（坐席绑定台账），返回 seats 列表"""
    dept_stats = OrderedDict()
    for dept in DEPT_ORDER:
        dept_stats[dept] = {'total': 0, 'bound': 0, 'notes': []}
    
    for r in sorted(grid.keys()):
        if r <= 1:  # 跳过标题和表头
            continue
        row = grid[r]
        dept = row.get(3, '').strip()
        status = row.get(11, '').strip()
        note = row.get(12, '').strip()
        
        # 跳过非正式部门（如样例行 "XXXX"）
        if not dept or dept not in DEPT_ORDER:
            continue
        
        dept_stats[dept]['total'] += 1
        if status == '已绑定':
            dept_stats[dept]['bound'] += 1
        if note:
            dept_stats[dept]['notes'].append(note)
    
    seats = []
    for dept in DEPT_ORDER:
        stats = dept_stats[dept]
        if stats['total'] == 0:
            continue
        note_str = summarize_notes(stats['notes'])
        seats.append({
            "dept": dept,
            "bound": stats['bound'],
            "total": stats['total'],
            "note": note_str
        })
    
    return seats

def calc_stats(tasks, seats):
    """计算统计数据"""
    real_tasks = [t for t in tasks if t['id'] and t['id'].isdigit()]
    total = len(real_tasks)
    done = sum(1 for t in real_tasks if t['status'] == '已完成')
    doing = sum(1 for t in real_tasks if t['status'] == '进行中')
    pending = sum(1 for t in real_tasks if t['status'] == '待开展')
    delayed = sum(1 for t in real_tasks if t['status'] == '延期')
    
    seats_bound = sum(s['bound'] for s in seats)
    seats_total = sum(s['total'] for s in seats)
    
    return {
        "total": total,
        "done": done,
        "doing": doing,
        "pending": pending,
        "delayed": delayed,
        "seatsBound": seats_bound,
        "seatsTotal": seats_total
    }

def main():
    print("正在从金山文档拉取数据...")
    
    # 拉取 Sheet1（任务计划）
    print("  拉取 Sheet1（任务计划表）...")
    grid1 = kdocs_get_sheet(1, 35, 9)
    tasks = parse_sheet1(grid1)
    print(f"  解析到 {len(tasks)} 条记录（含阶段标题）")
    
    # 拉取 Sheet4（坐席绑定台账）
    print("  拉取 Sheet4（坐席绑定台账）...")
    grid4 = kdocs_get_sheet(4, 82, 12)
    seats = parse_sheet4(grid4)
    print(f"  解析到 {len(seats)} 个部门")
    
    # 计算统计
    stats = calc_stats(tasks, seats)
    
    # 生成 JSON
    now = datetime.now().strftime('%Y-%m-%d %H:%M')
    data = {
        "syncTime": now,
        "source": "金山文档95598智能座席助手实施计划表",
        "tasks": tasks,
        "seats": seats,
        "stats": stats
    }
    
    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'plan-data.json')
    output_path = os.path.abspath(output_path)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"\n同步完成！时间: {now}")
    print(f"任务: {stats['total']}项 (完成{stats['done']}