#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
企业微信客户群统计报表生成脚本

功能：
1) 获取所有客户群与成员
2) 计算每日新增/退群，以及当月流失率（以月初基线为参考）
3) 可选：接入会话存档统计每日发言人数以计算互动率（发言人数/群总人数）
4) 使用 pandas 生成 Excel 报表，保存到本地
5) 依赖：requests, pandas, openpyxl

运行：
    python wechat_group_stats.py

配置：
    在下方填写 CORP_ID、APP_SECRET、AGENT_ID

说明：
- 群总人数与新增/退群统计以“外部联系人”（客户）为准，不包含企业内部成员。
- 会话存档对客户群的获取与解密较为复杂，此处提供可扩展的占位实现。如需启用，请在
  get_daily_group_speaking_user_count 中接入企业会话存档 SDK，并返回发言用户数。
"""

import os
import json
import time
import math
import glob
import shutil
import logging
import datetime as dt
from typing import Dict, List, Optional, Tuple, Set

import requests
import pandas as pd

# ======================== 配置区（请填写） ========================
CORP_ID: str = "YOUR_CORP_ID"
APP_SECRET: str = "YOUR_APP_SECRET"  # 对应有权限调用外部联系人/客户群接口的应用Secret
AGENT_ID: str = "YOUR_AGENT_ID"      # 这里预留但本脚本未直接使用
# ==============================================================

# 常量与目录
WORKSPACE_ROOT = "/workspace"
DATA_DIR = os.path.join(WORKSPACE_ROOT, "data")
REPORT_DIR = os.path.join(WORKSPACE_ROOT, "reports")
ACCESS_TOKEN_PATH = os.path.join(DATA_DIR, "access_token.json")
STATE_PATH = os.path.join(DATA_DIR, "state.json")

WECHAT_API_BASE = "https://qyapi.weixin.qq.com/cgi-bin"

# 日志
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)


# ======================== 工具函数 ========================
def ensure_dirs() -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(REPORT_DIR, exist_ok=True)


def today_str(tz: Optional[dt.tzinfo] = None) -> str:
    if tz is None:
        return dt.date.today().isoformat()
    return dt.datetime.now(tz=tz).date().isoformat()


def month_str(date_obj: Optional[dt.date] = None) -> str:
    d = date_obj or dt.date.today()
    return f"{d.year:04d}-{d.month:02d}"


def load_json(path: str, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def save_json(path: str, data) -> None:
    tmp = f"{path}.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    shutil.move(tmp, path)


# ======================== 企业微信 API ========================
def get_access_token(corp_id: str, corp_secret: str) -> str:
    ensure_dirs()
    now_ts = int(time.time())

    cached = load_json(ACCESS_TOKEN_PATH, default={})
    token = None
    if cached:
        token = cached.get("access_token")
        expires_at = cached.get("expires_at", 0)
        if token and now_ts < int(expires_at) - 60:  # 提前60秒刷新
            return token

    url = f"{WECHAT_API_BASE}/gettoken"
    params = {"corpid": corp_id, "corpsecret": corp_secret}
    resp = requests.get(url, params=params, timeout=10)
    resp.raise_for_status()
    data = resp.json()
    if data.get("errcode") != 0:
        raise RuntimeError(f"gettoken failed: {data}")

    token = data["access_token"]
    expires_in = int(data.get("expires_in", 7200))
    save_json(ACCESS_TOKEN_PATH, {
        "access_token": token,
        "expires_at": now_ts + expires_in,
        "fetched_at": now_ts,
    })
    return token


def wecom_post_json(url_path: str, token: str, payload: dict) -> dict:
    url = f"{WECHAT_API_BASE}/{url_path}"
    params = {"access_token": token}
    resp = requests.post(url, params=params, json=payload, timeout=20)
    resp.raise_for_status()
    data = resp.json()
    if data.get("errcode") != 0:
        raise RuntimeError(f"POST {url_path} failed: {data}")
    return data


def fetch_all_group_chat_ids(token: str) -> List[str]:
    """获取所有客户群 chat_id 列表（status_filter=0 全部）"""
    chat_ids: List[str] = []
    cursor: Optional[str] = None
    while True:
        payload = {
            "status_filter": 0,
            "limit": 1000,
        }
        if cursor:
            payload["cursor"] = cursor
        data = wecom_post_json("externalcontact/groupchat/list", token, payload)
        lst = data.get("group_chat_list", [])
        for item in lst:
            chat_id = item.get("chat_id")
            if chat_id:
                chat_ids.append(chat_id)
        cursor = data.get("next_cursor")
        if not cursor:
            break
        time.sleep(0.2)
    return chat_ids


def fetch_group_chat_detail(token: str, chat_id: str) -> dict:
    data = wecom_post_json("externalcontact/groupchat/get", token, {"chat_id": chat_id})
    group_chat = data.get("group_chat", {})
    return group_chat


# ======================== 数据抽取与快照 ========================
def normalize_external_members(member_list: List[dict]) -> Tuple[Set[str], Set[str]]:
    """
    返回 (external_ids, internal_user_ids)
    - external 使用 external_userid 作为标识
    - internal 使用 userid 作为标识
    """
    external_ids: Set[str] = set()
    internal_ids: Set[str] = set()
    for m in member_list or []:
        m_type = m.get("type")  # 1: 企业成员; 2: 外部联系人
        if m_type == 2:
            ext_id = m.get("external_userid")
            if ext_id:
                external_ids.add(ext_id)
        elif m_type == 1:
            uid = m.get("userid")
            if uid:
                internal_ids.add(uid)
    return external_ids, internal_ids


def build_today_snapshot(token: str, snapshot_date: Optional[str] = None) -> dict:
    """构建当日快照：每个群的外部联系人成员集合与群名"""
    snapshot_date = snapshot_date or today_str()
    chat_ids = fetch_all_group_chat_ids(token)
    logging.info(f"获取到客户群数量: {len(chat_ids)}")

    groups: Dict[str, dict] = {}
    for idx, chat_id in enumerate(chat_ids, start=1):
        try:
            detail = fetch_group_chat_detail(token, chat_id)
        except Exception as e:
            logging.warning(f"获取群详情失败 chat_id={chat_id}: {e}")
            time.sleep(0.5)
            continue

        name = detail.get("name") or detail.get("chat_id") or chat_id
        member_list = detail.get("member_list", [])
        external_ids, internal_ids = normalize_external_members(member_list)

        groups[chat_id] = {
            "name": name,
            "external_member_ids": sorted(list(external_ids)),
            "internal_member_ids": sorted(list(internal_ids)),
            "total_external_members": len(external_ids),
            "total_internal_members": len(internal_ids),
        }
        if idx % 50 == 0:
            logging.info(f"已处理 {idx}/{len(chat_ids)} 个群...")
        time.sleep(0.05)

    snapshot = {
        "date": snapshot_date,
        "groups": groups,
    }
    # 保存当日快照
    ensure_dirs()
    out_path = os.path.join(DATA_DIR, f"group_members_{snapshot_date}.json")
    save_json(out_path, snapshot)
    logging.info(f"已保存当日快照: {out_path}")
    return snapshot


def find_latest_snapshot(before_date: str) -> Optional[str]:
    """查找 before_date 之前最近的一次快照文件路径"""
    pattern = os.path.join(DATA_DIR, "group_members_*.json")
    files = sorted(glob.glob(pattern))
    if not files:
        return None
    try:
        target = dt.date.fromisoformat(before_date)
    except Exception:
        return None

    latest_path = None
    latest_date = None
    for p in files:
        base = os.path.basename(p)
        part = base.replace("group_members_", "").replace(".json", "")
        try:
            d = dt.date.fromisoformat(part)
        except Exception:
            continue
        if d < target and (latest_date is None or d > latest_date):
            latest_date = d
            latest_path = p
    return latest_path


def load_snapshot(path: str) -> dict:
    return load_json(path, default={})


# ======================== 月基线状态 ========================
class MonthBaseline:
    def __init__(self, state_path: str = STATE_PATH):
        self.state_path = state_path
        self.state = load_json(self.state_path, default={})

    def get_month_key(self) -> str:
        return month_str()

    def get_baseline(self) -> Dict[str, List[str]]:
        month_key = self.get_month_key()
        return self.state.get("month_baseline", {}).get(month_key, {})

    def set_baseline(self, groups_external: Dict[str, List[str]]) -> None:
        month_key = self.get_month_key()
        if "month_baseline" not in self.state:
            self.state["month_baseline"] = {}
        self.state["month_baseline"][month_key] = groups_external
        save_json(self.state_path, self.state)


# ======================== 会话存档（可选实现占位） ========================
def get_daily_group_speaking_user_count(chat_id: str, date_str: str) -> Optional[int]:
    """
    可选：接入企业微信“会话内容存档”统计当天发言用户数。
    由于客户群会话存档接入涉及密钥获取与消息解密，实际落地需使用官方/第三方SDK，
    并根据消息中的 roomid 与 chat_id 做关联后统计去重的发言用户数。

    如需启用，请在此函数内实现并返回 int；若无法获取，返回 None。
    """
    # 占位：默认不启用，返回 None
    return None


# ======================== 指标计算 ========================
def compute_daily_delta(
    today_snapshot: dict,
    prev_snapshot: Optional[dict]
) -> Dict[str, Dict[str, int]]:
    """计算每个群的新增/退群（外部联系人维度）"""
    result: Dict[str, Dict[str, int]] = {}
    today_groups: Dict[str, dict] = today_snapshot.get("groups", {})
    prev_groups: Dict[str, dict] = prev_snapshot.get("groups", {}) if prev_snapshot else {}

    for chat_id, info in today_groups.items():
        today_ext = set(info.get("external_member_ids", []))
        prev_ext = set(prev_groups.get(chat_id, {}).get("external_member_ids", []))
        added = len(today_ext - prev_ext)
        left = len(prev_ext - today_ext)
        result[chat_id] = {
            "added": added,
            "left": left,
        }
    return result


def compute_monthly_churn_rate(
    baseline: MonthBaseline,
    today_snapshot: dict
) -> Dict[str, Optional[float]]:
    """
    月流失率 = (基线外部联系人 - 今日外部联系人)/基线外部联系人。
    若无基线或分母为0，返回 None。
    """
    baseline_map = baseline.get_baseline()  # {chat_id: [ext_ids]}

    # 若本月基线不存在，则以今日为基线
    if not baseline_map:
        groups_external = {
            chat_id: info.get("external_member_ids", [])
            for chat_id, info in today_snapshot.get("groups", {}).items()
        }
        baseline.set_baseline(groups_external)
        baseline_map = groups_external

    rates: Dict[str, Optional[float]] = {}
    for chat_id, info in today_snapshot.get("groups", {}).items():
        today_ext = set(info.get("external_member_ids", []))
        base_ext = set(baseline_map.get(chat_id, []))
        base_n = len(base_ext)
        if base_n == 0:
            rates[chat_id] = None
        else:
            lost = len(base_ext - today_ext)
            rates[chat_id] = round(lost / base_n, 4)
    return rates


# ======================== 报表生成 ========================
def build_report_dataframe(
    today_snapshot: dict,
    daily_delta: Dict[str, Dict[str, int]],
    monthly_churn_rate: Dict[str, Optional[float]],
    date_str: str
) -> pd.DataFrame:
    rows: List[dict] = []
    for chat_id, info in today_snapshot.get("groups", {}).items():
        name = info.get("name", chat_id)
        total_external = int(info.get("total_external_members", 0))

        # 会话存档：发言人数（可选）
        speaking_users = get_daily_group_speaking_user_count(chat_id, date_str)
        interaction_rate = None
        if speaking_users is not None and total_external > 0:
            interaction_rate = round(speaking_users / total_external, 4)

        rows.append({
            "群名": name,
            "群ID": chat_id,
            "日期": date_str,
            "新增人数": int(daily_delta.get(chat_id, {}).get("added", 0)),
            "退群人数": int(daily_delta.get(chat_id, {}).get("left", 0)),
            "总人数": total_external,
            "互动率": interaction_rate,
            "月流失率": monthly_churn_rate.get(chat_id),
        })

    df = pd.DataFrame(rows, columns=[
        "群名", "群ID", "日期", "新增人数", "退群人数", "总人数", "互动率", "月流失率"
    ])
    return df


def save_report_excel(df: pd.DataFrame, date_str: str) -> str:
    ensure_dirs()
    filename = f"wechat_group_stats_{date_str}.xlsx"
    out_path = os.path.join(REPORT_DIR, filename)
    # 使用 openpyxl 引擎写入 xlsx
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="日报")
    logging.info(f"报表已生成: {out_path}")
    return out_path


# ======================== 主流程 ========================
def main():
    if (not CORP_ID or CORP_ID.startswith("YOUR_")) or (not APP_SECRET or APP_SECRET.startswith("YOUR_")) or (not AGENT_ID or AGENT_ID.startswith("YOUR_")):
        raise SystemExit("请先在脚本顶部填写真实的 CORP_ID、APP_SECRET、AGENT_ID 再运行。")

    ensure_dirs()

    date_str = today_str()
    logging.info(f"开始生成客户群日报: {date_str}")

    token = get_access_token(CORP_ID, APP_SECRET)

    # 1) 构建当日快照
    today_snapshot = build_today_snapshot(token, snapshot_date=date_str)

    # 2) 读取最近一次历史快照，计算新增/退群
    latest_path = find_latest_snapshot(before_date=date_str)
    prev_snapshot = load_snapshot(latest_path) if latest_path else None
    daily_delta = compute_daily_delta(today_snapshot, prev_snapshot)

    # 3) 计算月流失率（基于月初/本月首次运行的基线）
    baseline = MonthBaseline()
    monthly_churn = compute_monthly_churn_rate(baseline, today_snapshot)

    # 4) 组装报表并导出 Excel
    df = build_report_dataframe(today_snapshot, daily_delta, monthly_churn, date_str)
    save_report_excel(df, date_str)

    logging.info("全部完成。")


if __name__ == "__main__":
    main()