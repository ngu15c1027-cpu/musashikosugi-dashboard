#!/usr/bin/env python3
"""
リベ大デンタルクリニック武蔵小杉院 ダッシュボード更新スクリプト
毎朝7時にGitHub Actionsから自動実行される
"""

import os
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

# 監視ルームID（URLの #!rid の後の番号）
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
# Chatwork API
# ============================================================
def get_messages(room_id):
    """指定ルームの最新メッセージ（最大100件）を取得"""
    try:
        resp = requests.get(
            f'{CW_BASE}/rooms/{room_id}/messages',
            headers=CW_HEADERS,
            params={'force': 1},
            timeout=30
        )
        if resp.status_code == 200:
            return resp.json()
        print(f'[WARNING] Chatwork API {resp.status_code} / room={room_id}')
        return []
    except Exception as e:
        print(f'[ERROR] get_messages({room_id}): {e}')
        return []


# ============================================================
# 日報パース
# 対象フォーマット:
#   本日の診療
#   保険　25人　¥264,100
#   自費　1人　¥80,000
#   矯正　4人　¥0
#   物販　0人　¥0
#   ----------------------
#   計　¥344,100
#   実質人数：30人
#   今月の自費率：24%
# ============================================================
REPORT_RE = re.compile(
    r'本日の診療'
    r'.*?保険[\s\u3000]+(\d+)人[\s\u3000]+[¥￥]([\d,]+)'
    r'.*?自費[\s\u3000]+(\d+)人[\s\u3000]+[¥￥]([\d,]+)'
    r'.*?矯正[\s\u3000]+(\d+)人[\s\u3000]+[¥￥]([\d,]+)'
    r'.*?物販[\s\u3000]+(\d+)人[\s\u3000]+[¥￥]([\d,]+)'
    r'.*?計[\s\u3000]+[¥￥]([\d,]+)'
    r'.*?実質人数[：:]\s*(\d+)人'
    r'.*?今月の自費率[：:]\s*(\d+)%',
    re.DOTALL
)

def to_int(s):
    return int(s.replace(',', ''))

def parse_daily_reports(messages):
    """日報フォーマットのメッセージをパースしてリストで返す"""
    reports = []
    for msg in messages:
        body = msg.get('body', '')
        if '本日の診療' not in body:
            continue
        m = REPORT_RE.search(body)
        if not m:
            print(f'[WARN] 日報フォーマット不一致: {body[:80]}')
            continue
        send_time = msg.get('send_time', 0)
        dt = datetime.fromtimestamp(send_time, tz=JST)
        reports.append({
            'date':      dt.strftime('%m/%d'),
            'timestamp': send_time,
            'insurance': {'count': int(m.group(1)),  'amount': to_int(m.group(2))},
            'jihi':      {'count': int(m.group(3)),  'amount': to_int(m.group(4))},
            'kyosei':    {'count': int(m.group(5)),  'amount': to_int(m.group(6))},
            'hanpan':    {'count': int(m.group(7)),  'amount': to_int(m.group(8))},
            'total':     to_int(m.group(9)),
            'jissitsu':  int(m.group(10)),
            'jihiRate':  int(m.group(11)),
        })
    return reports


# ============================================================
# 既存データの読み込み・マージ
# ============================================================
def load_existing():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {'dailyReports': []}

def merge_reports(existing, new_reports):
    """日付をキーにマージ（新データ優先）。今月分のみ保持。"""
    merged = {r['date']: r for r in existing}
    for r in new_reports:
        merged[r['date']] = r
    now = datetime.now(JST)
    month_prefix = now.strftime('%m/')
    sorted_list = sorted(merged.values(), key=lambda x: x.get('timestamp', 0))
    return [r for r in sorted_list if r['date'].startswith(month_prefix)]


# ============================================================
# Claude API で経営分析
# ============================================================
def format_messages_for_claude(all_messages, daily_reports):
    """Claude に渡すコンテキストを組み立てる"""
    lines = []

    # 各ルームの最新50件
    for room_id, msgs in all_messages.items():
        name = ROOM_NAMES.get(room_id, room_id)
        lines.append(f'\n=== {name} ===')
        recent = msgs[-50:] if len(msgs) > 50 else msgs
        for msg in recent:
            dt = datetime.fromtimestamp(msg.get('send_time', 0), tz=JST)
            body = msg.get('body', '').strip()
            if body:
                lines.append(f'[{dt.strftime("%m/%d %H:%M")}] {body}')

    # 今月の診療実績サマリー
    if daily_reports:
        lines.append('\n=== 今月の診療実績 ===')
        for r in daily_reports:
            lines.append(
                f"{r['date']}: 合計¥{r['total']:,} / 実質{r['jissitsu']}人 "
                f"(保険¥{r['insurance']['amount']:,} / 自費¥{r['jihi']['amount']:,} "
                f"/ 矯正¥{r['kyosei']['amount']:,} / 物販¥{r['hanpan']['amount']:,})"
            )

    return '\n'.join(lines)

def analyze_with_claude(all_messages, daily_reports):
    """Claudeでチャットログと実績を分析し、JSONで返す"""
    client = anthropic.Anthropic(api_key=CLAUDE_API_KEY)
    context = format_messages_for_claude(all_messages, daily_reports)
    today_str = datetime.now(JST).strftime('%Y年%m月%d日')

    prompt = f"""あなたは歯科クリニック専門の経営コンサルタントです。
以下は「リベ大デンタルクリニック武蔵小杉院」の{today_str}時点のChatworkログと診療実績です。

{context}

---
上記の情報をもとに、院長が毎朝確認する経営レポートとして、以下のJSON形式で分析結果を返してください。
日本語で、具体的・実践的に記述してください。曖昧な表現は避け、数値や固有名詞を使って書いてください。

{{
  "summary": "経営状況の総括（200字程度）。数値を交えて現状を端的に説明する。",
  "goodPoints": [
    "具体的な良い点（数値・人名・出来事を含めて）",
    "良い点2",
    "良い点3"
  ],
  "improvements": [
    "具体的な改善が必要な点（なぜ問題かも含めて）",
    "改善点2",
    "改善点3"
  ],
  "actionPlans": [
    "アクションプラン（誰が・何を・いつまでに、の形式で）",
    "アクションプラン2",
    "アクションプラン3"
  ],
  "risks": [
    "注意が必要なリスク（放置した場合の影響も含めて）",
    "リスク2"
  ]
}}

JSONのみを返してください。説明文は不要です。"""

    try:
        message = client.messages.create(
            model='claude-sonnet-4-6',
            max_tokens=2000,
            messages=[{'role': 'user', 'content': prompt}]
        )
        text = message.content[0].text.strip()
        # JSON部分を抽出
        json_match = re.search(r'\{.*\}', text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group())
    except Exception as e:
        print(f'[ERROR] Claude API: {e}')

    return {
        'summary': '分析データを取得できませんでした。次回の実行をお待ちください。',
        'goodPoints': [],
        'improvements': [],
        'actionPlans': [],
        'risks': []
    }


# ============================================================
# メイン
# ============================================================
def main():
    now = datetime.now(JST)
    print(f'=== 更新開始: {now.strftime("%Y-%m-%d %H:%M JST")} ===')

    # 1. 既存データ読み込み
    existing = load_existing()

    # 2. Chatwork メッセージ取得
    print('Chatwork メッセージ取得中...')
    all_messages = {}
    for name, room_id in ROOMS.items():
        msgs = get_messages(room_id)
        all_messages[room_id] = msgs
        print(f'  {ROOM_NAMES[room_id]}: {len(msgs)}件')

    # 3. 日報パース → マージ
    new_reports = parse_daily_reports(all_messages[ROOMS['daily']])
    print(f'日報パース: {len(new_reports)}件')
    merged = merge_reports(existing.get('dailyReports', []), new_reports)
    print(f'今月累計: {len(merged)}日分')

    # 4. 月次集計
    monthly_total    = sum(r['total']                   for r in merged)
    monthly_patients = sum(r['jissitsu']                for r in merged)
    monthly_jihi     = sum(r['jihi']['amount']          for r in merged)
    monthly_insurance= sum(r['insurance']['amount']     for r in merged)
    monthly_kyosei   = sum(r['kyosei']['amount']        for r in merged)
    monthly_hanpan   = sum(r['hanpan']['amount']        for r in merged)
    work_days        = len(merged)
    jihi_rate        = round(monthly_jihi / monthly_total * 100) if monthly_total > 0 else 0
    avg_patients     = round(monthly_patients / work_days) if work_days > 0 else 0

    # 5. Claude 分析
    print('Claude APIで経営分析中...')
    analysis = analyze_with_claude(all_messages, merged)
    print(f'分析完了: 良い点{len(analysis.get("goodPoints",[]))}件 / リスク{len(analysis.get("risks",[]))}件')

    # 6. data.json 書き出し
    data = {
        'updatedAt':      now.isoformat(),
        'updatedAtLabel': now.strftime('%Y年%m月%d日 %H:%M'),
        'targetMonth':    now.strftime('%Y年%m月'),
        'dailyReports':   merged,
        'monthly': {
            'total':      monthly_total,
            'patients':   monthly_patients,
            'insurance':  monthly_insurance,
            'jihi':       monthly_jihi,
            'kyosei':     monthly_kyosei,
            'hanpan':     monthly_hanpan,
            'jihiRate':   jihi_rate,
            'workDays':   work_days,
            'avgPatients':avg_patients,
        },
        'latest': merged[-1] if merged else None,
        'analysis': analysis,
    }

    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f'=== data.json 更新完了 ===')
    print(f'  月次売上: ¥{monthly_total:,}')
    print(f'  平均患者数: {avg_patients}人/日')
    print(f'  自費率: {jihi_rate}%')


if __name__ == '__main__':
    main()
