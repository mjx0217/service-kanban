#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
班级明细服务沟通看板数据处理脚本

读取 Excel 底表 → 按沈阳基地聚合 → 生成 kanban_data.json

Usage:
    python preprocess.py <input.xlsx> [<output.json>]

依赖: pip install pandas openpyxl
"""
import sys
import json
import re
from pathlib import Path
from datetime import datetime
import pandas as pd


# Excel 列名(中文)
COL_BASE = "基地"
COL_SUBJECT = "学科"
COL_GROUP = "教学组"
COL_CT_ID = "班主任id"
COL_CT_NAME = "班主任名"
COL_PKG = "课包类型"
COL_INCLASS = "在班人次"
COL_BIND = "绑定人次"
COL_JOINGROUP = "进群人次"
COL_COVER = "覆盖人次_近7日"
COL_INTERACT = "交互人次_近7日"
COL_ACTIVE = "活跃人次_近7日"
COL_CONFIRM = "预出勤确认人次"
COL_ATTEND = "预出勤人次"

METRIC_COLS = [
    COL_INCLASS, COL_BIND, COL_JOINGROUP, COL_COVER,
    COL_INTERACT, COL_ACTIVE, COL_CONFIRM, COL_ATTEND,
]

TARGET_BASE = "沈阳"  # 固定处理沈阳基地


def pct(num, denom):
    """百分比:保留 1 位小数"""
    if not denom or denom <= 0:
        return 0
    return round(num / denom * 100, 1)


def parse_updated_from_filename(path):
    """从文件名提取更新时间,例如 班级明细_服务沟通数据-2026-09-03 08_30_00.xlsx"""
    name = Path(path).stem
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})\s+(\d{2})_(\d{2})_(\d{2})", name)
    if not m:
        return datetime.now().strftime("%Y-%m-%d %H:%M")
    y, mo, d, h, mi, s = m.groups()
    return f"{y}-{mo}-{d} {h}:{mi}"


def add_rates(row):
    """给一行数据加上比率字段"""
    ic = row["inClassCount"] or 1
    row["bindRate"] = pct(row["bindCount"], ic)
    row["joinRate"] = pct(row["joinGroupCount"], ic)
    row["confirmRate"] = pct(row["confirmCount"], ic)
    row["attendRate"] = pct(row["attendCount"], ic)
    row["coverRate"] = pct(row["coverCount"], ic)
    row["interactRate"] = pct(row["interactCount"], ic)
    row["activeRate"] = pct(row["activeCount"], ic)
    row["unbindCount"] = ic - row["bindCount"]
    return row


def agg_simple(rows, group_col):
    """按 group_col 字段分组聚合,返回带比率的对象列表"""
    out = {}
    metric_to_field = {
        COL_INCLASS: "inClassCount", COL_BIND: "bindCount",
        COL_JOINGROUP: "joinGroupCount", COL_COVER: "coverCount",
        COL_INTERACT: "interactCount", COL_ACTIVE: "activeCount",
        COL_CONFIRM: "confirmCount", COL_ATTEND: "attendCount",
    }
    for r in rows:
        key = r[group_col]
        if key not in out:
            obj = {"base": TARGET_BASE, group_col: key}
            for k in metric_to_field.values():
                obj[k] = 0
            out[key] = obj
        for src, dst in metric_to_field.items():
            out[key][dst] += r[src] or 0
    return [add_rates(v) for v in out.values()]


def main():
    if len(sys.argv) < 2:
        print("Usage: python preprocess.py <input.xlsx> [<output.json>]", file=sys.stderr)
        sys.exit(1)
    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "kanban_data.json"

    print(f"[preprocess] Reading: {input_path}")
    df = pd.read_excel(input_path)

    # 类型转换 + 沈阳过滤
    df = df[df[COL_BASE] == TARGET_BASE].copy()
    for c in METRIC_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    print(f"[preprocess] {TARGET_BASE} rows: {len(df)}")

    rows = df.to_dict("records")

    # META
    meta = {
        "updated": parse_updated_from_filename(input_path),
        "totalInClass": int(df[COL_INCLASS].sum()),
        "totalBind": int(df[COL_BIND].sum()),
        "totalJoinGroup": int(df[COL_JOINGROUP].sum()),
        "totalCover": int(df[COL_COVER].sum()),
        "totalInteract": int(df[COL_INTERACT].sum()),
        "totalActive": int(df[COL_ACTIVE].sum()),
    }
    meta["bindRate"] = pct(meta["totalBind"], meta["totalInClass"])
    meta["joinRate"] = pct(meta["totalJoinGroup"], meta["totalInClass"])
    meta["coverRate"] = pct(meta["totalCover"], meta["totalInClass"])
    meta["interactRate"] = pct(meta["totalInteract"], meta["totalInClass"])
    meta["activeRate"] = pct(meta["totalActive"], meta["totalInClass"])

    # SUBJECTS
    subjects = agg_simple(rows, COL_SUBJECT)

    # GROUPS
    groups = agg_simple(rows, COL_GROUP)

    # CLASS_TEACHERS
    ct_map = {}
    ct_pkgs = {}
    for r in rows:
        cid = str(r[COL_CT_ID])
        if cid not in ct_map:
            ct_map[cid] = {
                "ctid": cid,
                "name": r[COL_CT_NAME],
                "group": r[COL_GROUP],
                "base": TARGET_BASE,
                "inClassCount": 0, "bindCount": 0, "joinGroupCount": 0, "coverCount": 0,
                "interactCount": 0, "activeCount": 0, "confirmCount": 0, "attendCount": 0,
                "subjects": set(), "pkgs": set(),
            }
            ct_pkgs[cid] = set()
        o = ct_map[cid]
        o["inClassCount"] += r[COL_INCLASS]
        o["bindCount"] += r[COL_BIND]
        o["joinGroupCount"] += r[COL_JOINGROUP]
        o["coverCount"] += r[COL_COVER]
        o["interactCount"] += r[COL_INTERACT]
        o["activeCount"] += r[COL_ACTIVE]
        o["confirmCount"] += r[COL_CONFIRM]
        o["attendCount"] += r[COL_ATTEND]
        o["subjects"].add(r[COL_SUBJECT])
        o["pkgs"].add(r[COL_PKG])
        ct_pkgs[cid].add(r[COL_PKG])

    class_teachers = []
    for o in ct_map.values():
        o = add_rates(o)
        o["subjects"] = sorted(o["subjects"])
        o["pkgs"] = sorted(o["pkgs"])
        o["replyRate"] = 0
        o["avgRespTime"] = 0
        class_teachers.append(o)

    # CT_PKG_LIST
    cp_map = {}
    for r in rows:
        cid = str(r[COL_CT_ID])
        pkg = r[COL_PKG]
        key = (cid, pkg)
        if key not in cp_map:
            cp_map[key] = {
                "ctid": cid,
                "name": r[COL_CT_NAME],
                "group": r[COL_GROUP],
                "base": TARGET_BASE,
                "pkg": pkg,
                "inClassCount": 0, "bindCount": 0, "joinGroupCount": 0, "coverCount": 0,
                "interactCount": 0, "activeCount": 0, "confirmCount": 0, "attendCount": 0,
            }
        o = cp_map[key]
        o["inClassCount"] += r[COL_INCLASS]
        o["bindCount"] += r[COL_BIND]
        o["joinGroupCount"] += r[COL_JOINGROUP]
        o["coverCount"] += r[COL_COVER]
        o["interactCount"] += r[COL_INTERACT]
        o["activeCount"] += r[COL_ACTIVE]
        o["confirmCount"] += r[COL_CONFIRM]
        o["attendCount"] += r[COL_ATTEND]

    ct_pkg_list = []
    for (cid, pkg), o in cp_map.items():
        o = add_rates(o)
        o["subjects"] = sorted(ct_map[cid]["subjects"])
        o["pkgs"] = sorted(ct_pkgs[cid])
        o["replyRate"] = 0
        o["avgRespTime"] = 0
        ct_pkg_list.append(o)

    meta["ctCount"] = len(class_teachers)
    meta["groupCount"] = len(groups)
    meta["subjectCount"] = len(subjects)
    meta["teacherCount"] = len(class_teachers)

    out = {
        "meta": meta,
        "subjects": subjects,
        "groups": groups,
        "classTeachers": class_teachers,
        "ctPkgList": ct_pkg_list,
    }

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print(f"[preprocess] Wrote: {output_path}")
    print(f"[preprocess] meta: {json.dumps(meta, ensure_ascii=False)}")
    print(f"[preprocess] subjects={len(subjects)}, groups={len(groups)}, "
          f"classTeachers={len(class_teachers)}, ctPkgList={len(ct_pkg_list)}")


if __name__ == "__main__":
    main()
