#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从爱豆糖推送系统下载当日底表 Excel

依赖: pip install requests
环境变量:
    AIDOUTANG_COOKIE  浏览器导出的完整 Cookie 字符串
    PUSH_TASK_ID      推送任务 ID(URL 里的 id 参数)

Usage:
    python download.py [<output.xlsx>]
"""
import os
import sys
import requests


API_BASE = "https://dw.aidoutang.com"


def main():
    cookie = os.environ.get("AIDOUTANG_COOKIE", "").strip()
    task_id = os.environ.get("PUSH_TASK_ID", "").strip()
    output = sys.argv[1] if len(sys.argv) > 1 else "今日底表.xlsx"

    if not cookie:
        print("[download] ERROR: 环境变量 AIDOUTANG_COOKIE 未设置", file=sys.stderr)
        sys.exit(1)
    if not task_id:
        print("[download] ERROR: 环境变量 PUSH_TASK_ID 未设置", file=sys.stderr)
        sys.exit(1)

    session = requests.Session()
    session.headers.update({
        "Cookie": cookie,
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Referer": "https://xwv5.aidoutang.com/dwpush/",
    })

    # 1. 检查登录
    print("[download] 验证会话...")
    check = session.get(f"{API_BASE}/pushservice/api/checkLogin", timeout=30)
    check.raise_for_status()
    check_data = check.json()
    if check_data.get("code") != 0:
        print(f"[download] 会话已失效: {check_data}", file=sys.stderr)
        print("[download] 请重新登录爱豆糖并导出 cookie,更新 GitHub Secret AIDOUTANG_COOKIE", file=sys.stderr)
        sys.exit(2)

    print(f"[download] 会话有效: {check_data.get('data', {})}")

    # 2. 拉取任务信息
    print(f"[download] 获取任务信息 (id={task_id})...")
    info = session.get(f"{API_BASE}/pushservice/task/info", params={"id": task_id}, timeout=30)
    info.raise_for_status()
    info_data = info.json()
    if info_data.get("code") != 0:
        print(f"[download] 获取任务信息失败: {info_data}", file=sys.stderr)
        sys.exit(3)
    print(f"[download] 任务: {info_data.get('data', {}).get('name', '(unknown)')}")

    # 3. 下载文件
    print(f"[download] 下载 Excel 到 {output}...")
    resp = session.get(
        f"{API_BASE}/pushservice/task/downloadFile",
        params={"id": task_id},
        timeout=120,
        allow_redirects=True,
    )
    resp.raise_for_status()

    # 校验是不是 Excel
    content_type = resp.headers.get("Content-Type", "")
    if "json" in content_type.lower():
        print(f"[download] 响应非文件: {resp.text[:500]}", file=sys.stderr)
        sys.exit(4)

    with open(output, "wb") as f:
        f.write(resp.content)

    size = len(resp.content)
    print(f"[download] 下载完成: {output} ({size:,} bytes)")


if __name__ == "__main__":
    main()
