#!/usr/bin/env python3
"""
リベ大デンタルクリニック武蔵小杉院 ダッシュボード更新スクリプト
毎朝7時にGitHub Actionsから自動実行される

データ取得方針:
  - 診療実績 → Googleスプレッドシート（月次完全データ）
  - 経営分析 → Chatworkログ（Claude APIで分析）
"""

import os
import io
import csv
import json
import re
import requests
from datetime import datetime, timezone, timedelta
import anthropic

# ============================================================
# 設定
# ============================================================
CHATWORK_TOKEN = os.environ['CHATWORK_API_TOKEN']
CLAUDE_API_KEY = os.environ['CLAUDE_API_KEY']

JST = timezone(timedelta(hours=9))

# スプレッドシートID（診療実績）
SHEET_ID = '1c3yPZdER5i4e0syuGkyFabF3cuG7VF_kVZBQ_rNGVL4'

# 監視ルームID
ROOMS = {
    'daily': '410972239',   # 日報・業務連絡チャット
    'dr':    '422210775',   # Dr_すり合わせチャット
    'jimu':  '421087514',   # 医療事務チャット
}

ROOM_NAMES = {
    '410972239': '日報・業務連絡チャット',
    '422210775': 'Dr_すり合わせチャット',
    '421087514': '医療事務 大塚さんチャット',
}

CW_BASE    = 'https://api.chatwork.com/v2'
CW_HEADERS = {'X-ChatWorkToken': CHATWORK_TOKEN}
DATA_FILE  = 'data.json'


# ============================================================
# ① Googleスプレッドシートから診療実績を取得
# ============================================================
def fetch_spreadsheet_reports():
    """スプレッドシートのCSVを取得してパース"""
    url = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv'
    try:
        resp = requests.get(url, allow_redirects=True, timeout=30)
        if resp.status_code != 200:
            print(f'[WARN] Spreadsheet HTTP {resp.status_code}')
            return []
        resp.encoding = 'utf-8'
        reader = csv.reader(io.StringIO(resp.text))
        rows = list(reader)
    except Exception as e:
        print(f'[ERROR] fetch_spreadsheet_reports: {e}')
        return []

    now = datetime.now(JST)
    year = now.year

    reports = []
    for row in rows:
        if not row or not row[0]:
            continue
        date_str = row[0].strip()

        # MM/DD 形式のみ処理
        if not re.match(r'^\d{2}/\d{2}$', date_str):
            continue

        def col(i):
            if i >= len(row):
                return 0
            v = row[i].strip().replace(',', '')
            try:
                return int(float(v)) if v else 0
            except ValueError:
                return 0

        patients = col(2)
        total    = col(12)
        if patients == 0 and total == 0:
            continue  # 休診日スキップ

        # タイムスタンプ（ソート・重複排除用）
        try:
            m, d = date_str.split('/')
            dt = datetime(year, int(m), int(d), 12, 0, 0, tzinfo=JST)
            ts = int(dt.timestamp())
        except Exception:
            ts = 0

        reports.append({
            'date':      date_str,
            'timestamp': ts,
            'insurance': {'count': col(4),  'amount': col(5)},
            'jihi':      {'count': col(6),  'amount': col(7)},
            'kyosei':    {'count': col(8),  'amount': col(9)},
            'hanpan':    {'count': col(10), 'amount': col(11)},
            'total':     total,
            'jissitsu':  patients,
            'jihiRate':  col(13),
        })

    # 日付順ソート
    reports.sort(key=lambda x: x['timestamp'])
    print(f'  スプレッドシート: {len(reports)}件取得')
    return reports


# ============================================================
# ② Chatwork API からメッセージ取得（分析用）
# ============================================================
def get_messages(room_id):
    try:
        resp = requests.get(
            f'{CW_BASE}/rooms/{room_id}/messages',
            headers=CW_HEADERS,
            params={'force': 1},
            timeout=30
        )
        if resp.status_code == 200:
            return resp.json()
        print(f'[WARN] Chatwork {resp.status_code} / room={room_id}')
        return []
    except Exception as e:
        print(f'[ERROR] get_messages({room_id}): {e}')
        return []


# ============================================================
# ③ Claude API で経営分析
# ============================================================
def format_context(all_messages, daily_reports):
    lines = []
    for room_id, msgs in all_messages.items():
        name = ROOM_NAMES.get(room_id, room_id)
        lines.append(f'\n=== {name} ===')
        recent = msgs[-50:] if len(msgs) > 50 else msgs
        for msg in recent:
            dt = datetime.fromtimestamp(msg.get('send_time', 0), tz=JST)
            body = msg.get('body', '').strip()
            if body:
                lines.append(f'[{dt.strftime("%m/%d %H:%M")}] {body}')

    if daily_reports:
        lines.append('\n=== 今月の診療実績 ===')
        for r in daily_reports:
            lines.append(
                f"{r['date']}: 合計¥{r['total']:,} / 実質{r['jissitsu']}人 "
                f"(保険¥{r['insurance']['amount']:,} / 自費¥{r['jihi']['amount']:,} "
                f"/ 矯正¥{r['kyosei']['amount']:,})"
            )
    return '\n'.join(lines)

def analyze_with_claude(all_messages, daily_reports):
    client = anthropic.Anthropic(api_key=CLAUDE_API_KEY)
    context = format_context(all_messages, daily_reports)
    today_str = datetime.now(JST).strftime('%Y年%m月%d日')

    prompt = f"""あなたは歯科クリニック専門の経営コンサルタントです。
以下は「リベ大デンタルクリニック武蔵小杉院」の{today_str}時点のChatworkログと診療実績です。

{context}

---
院長が毎朝確認する経営レポートとして、以下のJSON形式で返してください。
日本語で、具体的・実践的に記述してください。数値・人名・出来事を積極的に使ってください。

{{
  "summary": "経営状況の総括（200字程度）",
  "goodPoints": ["良い点1（具体的に）", "良い点2", "良い点3"],
  "improvements": ["改善点1（なぜ問題かも含めて）", "改善点2", "改善点3"],
  "actionPlans": ["アクションプラン1（誰が・何を・いつまでに）", "アクションプラン2", "アクションプラン3"],
  "risks": ["リスク1（放置した場合の影響も）", "リスク2"]
}}

JSONのみを返してください。"""

    try:
        message = client.messages.create(
            model='claude-sonnet-4-6',
            max_tokens=2000,
            messages=[{'role': 'user', 'content': prompt}]
        )
        text = message.content[0].text.strip()
        m = re.search(r'\{.*\}', text, re.DOTALL)
        if m:
            return json.loads(m.group())
    except Exception as e:
        print(f'[ERROR] Claude API: {e}')

    return {
        'summary': '分析データを取得できませんでした。次回の実行をお待ちください。',
        'goodPoints': [], 'improvements': [], 'actionPlans': [], 'risks': []
    }


# ============================================================
# メイン
# ============================================================
def main():
    now = datetime.now(JST)
    print(f'=== 更新開始: {now.strftime("%Y-%m-%d %H:%M JST")} ===')

    # 1. スプレッドシートから診療実績取得
    print('スプレッドシートから診療実績を取得中...')
    reports = fetch_spreadsheet_reports()

    # 2. Chatworkからメッセージ取得（分析用）
    print('Chatworkメッセージ取得中...')
    all_messages = {}
    for name, room_id in ROOMS.items():
        msgs = get_messages(room_id)
        all_messages[room_id] = msgs
        print(f'  {ROOM_NAMES[room_id]}: {len(msgs)}件')

    # 3. 月次集計
    monthly_total     = sum(r['total']                for r in reports)
    monthly_patients  = sum(r['jissitsu']             for r in reports)
    monthly_jihi      = sum(r['jihi']['amount']       for r in reports)
    monthly_insurance = sum(r['insurance']['amount']  for r in reports)
    monthly_kyosei    = sum(r['kyosei']['amount']      for r in reports)
    monthly_hanpan    = sum(r['hanpan']['amount']      for r in reports)
    work_days         = len(reports)
    jihi_rate         = round(monthly_jihi / monthly_total * 100) if monthly_total > 0 else 0
    avg_patients      = round(monthly_patients / work_days) if work_days > 0 else 0

    # 4. Claude分析
    print('Claude APIで経営分析中...')
    analysis = analyze_with_claude(all_messages, reports)

    # 5. data.json 書き出し
    data = {
        'updatedAt':       now.isoformat(),
        'updatedAtLabel':  now.strftime('%Y年%m月%d日 %H:%M'),
        'targetMonth':     now.strftime('%Y年%m月'),
        'dailyReports':    reports,
        'monthly': {
            'total':       monthly_total,
            'patients':    monthly_patients,
            'insurance':   monthly_insurance,
            'jihi':        monthly_jihi,
            'kyosei':      monthly_kyosei,
            'hanpan':      monthly_hanpan,
            'jihiRate':    jihi_rate,
            'workDays':    work_days,
            'avgPatients': avg_patients,
        },
        'latest':   reports[-1] if reports else None,
        'analysis': analysis,
    }

    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f'=== 完了 ===')
    print(f'  診療日数: {work_days}日')
    print(f'  月次売上: ¥{monthly_total:,}')
    print(f'  平均患者数: {avg_patients}人/日')
    print(f'  自費率: {jihi_rate}%')

if __name__ == '__main__':
    main()
