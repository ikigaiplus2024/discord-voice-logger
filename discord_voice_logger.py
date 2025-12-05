import discord
from discord.ext import commands
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timezone, timedelta
import os
import json
import asyncio
from concurrent.futures import ThreadPoolExecutor

# 環境変数から設定を取得
SPREADSHEET_ID = os.environ.get('SPREADSHEET_ID')
DISCORD_TOKEN = os.environ.get('DISCORD_TOKEN')
CREDENTIALS_JSON = os.environ.get('CREDENTIALS_JSON')

# Google Sheets の設定
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SHEET_NAME = 'ボイスログ'

# 日本時間（JST）のタイムゾーン設定
JST = timezone(timedelta(hours=9))

# スレッドプール
executor = ThreadPoolExecutor(max_workers=3)

# Google Sheets 認証（環境変数から）
def get_google_sheets_client():
    if CREDENTIALS_JSON:
        creds_dict = json.loads(CREDENTIALS_JSON)
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
    return gspread.authorize(creds)

# スプレッドシートの初期化
def initialize_sheet():
    try:
        client = get_google_sheets_client()
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        
        try:
            sheet = spreadsheet.worksheet(SHEET_NAME)
        except gspread.WorksheetNotFound:
            sheet = spreadsheet.add_worksheet(title=SHEET_NAME, rows=1000, cols=6)
        
        headers = ['日付', '名前', 'ID', '部屋の名前', '入室時間', '退出時間']
        if sheet.row_values(1) != headers:
            sheet.update([headers], 'A1:F1')
        
        print("✅ スプレッドシート初期化完了")
        return sheet
    except Exception as e:
        print(f"❌ スプレッドシート初期化エラー: {e}")
        return None

# ログをスプレッドシートに追加
def log_to_sheet(date, name, user_id, channel_name, join_time, leave_time=""):
    try:
        client = get_google_sheets_client()
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        sheet = spreadsheet.worksheet(SHEET_NAME)
        
        row = [date, name, str(user_id), channel_name, join_time, leave_time]
        sheet.append_row(row, value_input_option='USER_ENTERED')
        print(f"📝 入室記録: {name} - {channel_name} ({join_time})")
        
    except Exception as e:
        print(f"❌ スプレッドシート書き込みエラー: {e}")

# 退出時間を既存の行に更新
def update_leave_time(user_id, channel_name, leave_time):
    try:
        client = get_google_sheets_client()
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        sheet = spreadsheet.worksheet(SHEET_NAME)
        
        all_values = sheet.get_all_values()
        
        for i in range(len(all_values) - 1, 0, -1):
            row = all_values[i]
            if len(row) >= 6:
                if row[2] == str(user_id) and row[3] == channel_name and (len(row) < 6 or row[5] == ""):
                    sheet.update_cell(i + 1, 6, leave_time)
                    print(f"📝 退出記録: {row[1]} - {channel_name} ({leave_time})")
                    return True
        
        print(f"⚠️ 入室記録が見つかりません: UserID={user_id}, Channel={channel_name}")
        return False
        
    except Exception as e:
        print(f"❌ 退出時間更新エラー: {e}")
        return False

user_join_times = {}

intents = discord.Intents.default()
intents.voice_states = True
intents.guilds = True
intents.members = True

bot = commands.Bot(command_prefix='!', intents=intents)

@bot.event
async def on_ready():
    print(f'✅ {bot.user} としてログインしました')
    print('👀 ボイスチャンネルの監視を開始します...')
    print(f'🕐 現在の日本時間: {datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")}')
    
    # スプレッドシートの初期化を別スレッドで実行
    loop = asyncio.get_event_loop()
    await loop.run_in_executor(executor, initialize_sheet)

@bot.event
async def on_voice_state_update(member, before, after):
    now = datetime.now(JST)
    date = now.strftime('%Y-%m-%d')
    time_str = now.strftime('%H:%M:%S')
    
    # 非同期で実行するためのラッパー
    loop = asyncio.get_event_loop()
    
    if before.channel is None and after.channel is not None:
        key = f"{member.id}_{after.channel.id}"
        user_join_times[key] = time_str
        print(f"🟢 入室: {member.name} → {after.channel.name} ({time_str})")
        
        # 別スレッドで実行
        loop.run_in_executor(executor, log_to_sheet, date, member.name, member.id, after.channel.name, time_str, "")
    
    elif before.channel is not None and after.channel is None:
        key = f"{member.id}_{before.channel.id}"
        print(f"🔴 退出: {member.name} ← {before.channel.name} ({time_str})")
        
        # 別スレッドで実行
        loop.run_in_executor(executor, update_leave_time, member.id, before.channel.name, time_str)
        
        if key in user_join_times:
            del user_join_times[key]
    
    elif before.channel is not None and after.channel is not None and before.channel != after.channel:
        key_before = f"{member.id}_{before.channel.id}"
        print(f"🔄 移動: {member.name} {before.channel.name} → {after.channel.name}")
        
        # 別スレッドで実行
        loop.run_in_executor(executor, update_leave_time, member.id, before.channel.name, time_str)
        
        if key_before in user_join_times:
            del user_join_times[key_before]
        
        key_after = f"{member.id}_{after.channel.id}"
        user_join_times[key_after] = time_str
        
        # 別スレッドで実行
        loop.run_in_executor(executor, log_to_sheet, date, member.name, member.id, after.channel.name, time_str, "")

if __name__ == "__main__":
    if not DISCORD_TOKEN:
        print("❌ DISCORD_TOKEN が設定されていません")
    elif not SPREADSHEET_ID:
        print("❌ SPREADSHEET_ID が設定されていません")
    else:
        try:
            print("🚀 Bot を起動しています...")
            bot.run(DISCORD_TOKEN)
        except Exception as e:
            print(f"❌ Bot起動エラー: {e}")
