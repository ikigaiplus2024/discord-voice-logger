import discord
from discord.ext import commands, tasks
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timezone, timedelta
import os
import json
import asyncio

# 環境変数から設定を取得
SPREADSHEET_ID = os.environ.get('SPREADSHEET_ID')
DISCORD_TOKEN = os.environ.get('DISCORD_TOKEN')
CREDENTIALS_JSON = os.environ.get('CREDENTIALS_JSON')

# Google Sheets の設定
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SHEET_NAME = 'ボイスログ'

# 日本時間(JST)のタイムゾーン設定
JST = timezone(timedelta(hours=9))

# Google Sheets クライアント(グローバルで保持)
sheets_client = None
spreadsheet = None
sheet = None

# Google Sheets 認証
def get_google_sheets_client():
    global sheets_client
    if sheets_client is None:
        if CREDENTIALS_JSON:
            creds_dict = json.loads(CREDENTIALS_JSON)
            creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        else:
            creds = Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
        sheets_client = gspread.authorize(creds)
    return sheets_client

# スプレッドシートの初期化(同期関数)
def initialize_sheet_sync():
    global spreadsheet, sheet
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
        return True
    except Exception as e:
        print(f"❌ スプレッドシート初期化エラー: {e}")
        return False

# ログをスプレッドシートに追加(入室時)
def log_to_sheet(date, name, user_id, channel_name, join_time, leave_time=""):
    try:
        global sheet
        if sheet is None:
            print("⚠️ スプレッドシート未初期化")
            return
        
        row = [date, name, str(user_id), channel_name, join_time, leave_time]
        sheet.append_row(row, value_input_option='USER_ENTERED')
        print(f"📝 入室記録: {name} - {channel_name} ({join_time})")
        
    except Exception as e:
        print(f"❌ スプレッドシート書き込みエラー: {e}")

# 退出時間を既存の行に更新
def update_leave_time(user_id, channel_name, leave_time):
    try:
        global sheet
        if sheet is None:
            print("⚠️ スプレッドシート未初期化")
            return False
        
        all_values = sheet.get_all_values()
        
        for i in range(len(all_values) - 1, 0, -1):
            row = all_values[i]
            if len(row) >= 3:
                if row[2] == str(user_id) and row[3] == channel_name:
                    if len(row) < 6 or row[5] == "":
                        sheet.update_cell(i + 1, 6, leave_time)
                        print(f"📝 退出記録: {row[1]} - {channel_name} ({leave_time})")
                        return True
        
        print(f"⚠️ 入室記録が見つかりません: UserID={user_id}, Channel={channel_name}")
        return False
        
    except Exception as e:
        print(f"❌ 退出時間更新エラー: {e}")
        return False

# Discord Bot の設定
intents = discord.Intents.default()
intents.voice_states = True
intents.guilds = True
intents.members = True

bot = commands.Bot(command_prefix='!', intents=intents)

@bot.event
async def on_ready():
    """Bot起動時の処理"""
    print(f'✅ {bot.user} としてログインしました')
    print(f'🕐 現在の日本時間: {datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")}')
    print('👀 ボイスチャンネルの監視を開始します...')
    
    # スプレッドシート初期化を別スレッドで実行
    loop = asyncio.get_event_loop()
    success = await loop.run_in_executor(None, initialize_sheet_sync)
    if success:
        print("📊 スプレッドシートの記録準備完了")
    
    # 稼働確認ログを開始
    if not keep_alive.is_running():
        keep_alive.start()

@bot.event
async def on_voice_state_update(member, before, after):
    """ボイスチャンネルの入退室を検知"""
    
    # スプレッドシート未初期化の場合はスキップ
    if sheet is None:
        print("⏳ スプレッドシート初期化待機中...")
        return
    
    now = datetime.now(JST)
    date = now.strftime('%Y-%m-%d')
    time_str = now.strftime('%H:%M:%S')
    
    loop = asyncio.get_event_loop()
    
    # 入室検知
    if before.channel is None and after.channel is not None:
        print(f"🟢 入室: {member.display_name} → {after.channel.name} ({time_str})")
        
        await loop.run_in_executor(
            None,
            log_to_sheet,
            date,
            member.display_name,
            member.id,
            after.channel.name,
            time_str,
            ""
        )
    
    # 退出検知
    elif before.channel is not None and after.channel is None:
        print(f"🔴 退出: {member.display_name} ← {before.channel.name} ({time_str})")
        
        await loop.run_in_executor(
            None,
            update_leave_time,
            member.id,
            before.channel.name,
            time_str
        )
    
    # チャンネル移動検知
    elif before.channel is not None and after.channel is not None and before.channel != after.channel:
        print(f"🔄 移動: {member.display_name} {before.channel.name} → {after.channel.name}")
        
        # 前のチャンネルの退出時間を更新
        await loop.run_in_executor(
            None,
            update_leave_time,
            member.id,
            before.channel.name,
            time_str
        )
        
        # 新しいチャンネルへの入室を記録
        await loop.run_in_executor(
            None,
            log_to_sheet,
            date,
            member.display_name,
            member.id,
            after.channel.name,
            time_str,
            ""
        )

# 定期的な稼働確認ログ(Render.comのスリープ防止)
@tasks.loop(minutes=5)
async def keep_alive():
    print(f"💓 稼働中... {datetime.now(JST).strftime('%H:%M:%S')}")

# Bot を起動
if __name__ == "__main__":
    if not DISCORD_TOKEN:
        print("❌ DISCORD_TOKEN が設定されていません")
    elif not SPREADSHEET_ID:
        print("❌ SPREADSHEET_ID が設定されていません")
    else:
        print("🚀 Bot を起動しています...")
        bot.run(DISCORD_TOKEN)
