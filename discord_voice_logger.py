import discord
from discord.ext import commands
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timezone, timedelta
import os
import json
import time
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

# Google Sheets APIアクセス制御用のロック
sheets_lock = asyncio.Lock()

# Google Sheets 認証(環境変数から)
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

# Google Sheets APIへのリトライ付きアクセス(改良版)
def retry_sheets_operation(operation, max_retries=5, initial_delay=1):
    """Google Sheets操作を失敗時にリトライする(指数バックオフ)"""
    for attempt in range(max_retries):
        try:
            result = operation()
            # 成功後、短時間待機(APIレート制限対策)
            time.sleep(0.5)
            return result
        except Exception as e:
            if attempt < max_retries - 1:
                delay = initial_delay * (2 ** attempt)  # 指数バックオフ
                print(f"⚠️ Google Sheets API エラー(リトライ {attempt + 1}/{max_retries}): {e}")
                print(f"   {delay}秒後に再試行します...")
                time.sleep(delay)
            else:
                print(f"❌ Google Sheets API エラー(最終試行失敗): {e}")
                raise

# ログをスプレッドシートに追加(非同期対応版)
async def log_to_sheet_async(date, name, user_id, channel_name, join_time, leave_time=""):
    try:
        print(f"📊 [DEBUG] log_to_sheet開始: {name} (ID:{user_id}) - {channel_name}")
        
        # ロックを取得して排他制御
        async with sheets_lock:
            def add_row():
                client = get_google_sheets_client()
                spreadsheet = client.open_by_key(SPREADSHEET_ID)
                sheet = spreadsheet.worksheet(SHEET_NAME)
                row = [date, name, str(user_id), channel_name, join_time, leave_time]
                sheet.append_row(row, value_input_option='USER_ENTERED')
                return True
            
            # 非同期処理で実行
            loop = asyncio.get_event_loop()
            await loop.run_in_executor(None, lambda: retry_sheets_operation(add_row))
            
        print(f"✅ 入室記録成功: {name} - {channel_name} ({join_time})")
        return True
        
    except Exception as e:
        print(f"❌ スプレッドシート書き込みエラー: {e}")
        print(f"   データ: date={date}, name={name}, user_id={user_id}, channel={channel_name}")
        return False

# 退出時間を既存の行に更新(非同期対応版)
async def update_leave_time_async(user_id, channel_name, leave_time, display_name="不明"):
    try:
        print(f"📊 [DEBUG] update_leave_time開始: {display_name} (ID:{user_id}) - {channel_name}")
        
        # ロックを取得して排他制御
        async with sheets_lock:
            def update_operation():
                client = get_google_sheets_client()
                spreadsheet = client.open_by_key(SPREADSHEET_ID)
                sheet = spreadsheet.worksheet(SHEET_NAME)
                
                # 全データを取得
                all_values = sheet.get_all_values()
                print(f"📊 [DEBUG] 取得した行数: {len(all_values)}")
                
                # 最後の行から遡って、該当ユーザーの入室記録を探す
                for i in range(len(all_values) - 1, 0, -1):
                    row = all_values[i]
                    
                    # 安全に列データを取得
                    row_id = row[2] if len(row) > 2 else ""
                    row_channel = row[3] if len(row) > 3 else ""
                    row_leave_time = row[5] if len(row) > 5 else ""
                    
                    # IDと部屋名が一致し、退出時間が空欄の行を探す
                    if row_id == str(user_id) and row_channel == channel_name and row_leave_time == "":
                        print(f"📊 [DEBUG] 一致する行を発見: 行{i+1}")
                        sheet.update_cell(i + 1, 6, leave_time)
                        row_name = row[1] if len(row) > 1 else "不明"
                        print(f"✅ 退出記録成功: {row_name} - {channel_name} ({leave_time})")
                        return True
                
                print(f"⚠️ 入室記録が見つかりません: UserID={user_id}, Channel={channel_name}")
                return False
            
            # 非同期処理で実行
            loop = asyncio.get_event_loop()
            result = await loop.run_in_executor(None, lambda: retry_sheets_operation(update_operation))
        
        # フォールバック: 入室記録が見つからない場合、新しい行を作成
        if not result:
            print(f"⚠️ フォールバック処理: 新しい行として記録します")
            now = datetime.now(JST)
            date = now.strftime('%Y-%m-%d')
            await log_to_sheet_async(date, display_name, user_id, channel_name, "不明", leave_time)
            return True
        
        return result
        
    except Exception as e:
        print(f"❌ 退出時間更新エラー: {e}")
        print(f"   データ: user_id={user_id}, channel={channel_name}, leave_time={leave_time}")
        
        # エラー時もフォールバック処理を試みる
        try:
            print(f"⚠️ エラー後のフォールバック処理を実行")
            now = datetime.now(JST)
            date = now.strftime('%Y-%m-%d')
            await log_to_sheet_async(date, display_name, user_id, channel_name, "不明", leave_time)
        except Exception as fallback_error:
            print(f"❌ フォールバック処理も失敗: {fallback_error}")
        
        return False

# 入室時間を記録する辞書
user_join_times = {}

# Discord Bot の設定
intents = discord.Intents.default()
intents.voice_states = True
intents.guilds = True
intents.members = True
intents.message_content = True  # WARNINGを消すために追加

bot = commands.Bot(command_prefix='!', intents=intents)

@bot.event
async def on_ready():
    print(f'✅ {bot.user} としてログインしました')
    print('👀 ボイスチャンネルの監視を開始します...')
    print(f'🕐 現在の日本時間: {datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")}')
    
    # スプレッドシートの初期化
    loop = asyncio.get_event_loop()
    await loop.run_in_executor(None, initialize_sheet)

@bot.event
async def on_voice_state_update(member, before, after):
    """ボイスチャンネルの入退室を検知"""
    
    try:
        # 日本時間を取得
        now = datetime.now(JST)
        date = now.strftime('%Y-%m-%d')
        time_str = now.strftime('%H:%M:%S')
        
        print(f"\n{'='*60}")
        print(f"🎯 [イベント検知] {member.display_name} (ID:{member.id})")
        print(f"   Before: {before.channel.name if before.channel else 'None'}")
        print(f"   After: {after.channel.name if after.channel else 'None'}")
        print(f"   時刻: {time_str}")
        print(f"{'='*60}\n")
        
        # 入室検知
        if before.channel is None and after.channel is not None:
            key = f"{member.id}_{after.channel.id}"
            user_join_times[key] = time_str
            
            print(f"🟢 [入室] {member.display_name} → {after.channel.name} ({time_str})")
            
            await log_to_sheet_async(
                date=date,
                name=member.display_name,
                user_id=member.id,
                channel_name=after.channel.name,
                join_time=time_str,
                leave_time=""
            )
        
        # 退出検知
        elif before.channel is not None and after.channel is None:
            key = f"{member.id}_{before.channel.id}"
            
            print(f"🔴 [退出] {member.display_name} ← {before.channel.name} ({time_str})")
            
            await update_leave_time_async(
                user_id=member.id,
                channel_name=before.channel.name,
                leave_time=time_str,
                display_name=member.display_name
            )
            
            if key in user_join_times:
                del user_join_times[key]
        
        # チャンネル移動検知
        elif before.channel is not None and after.channel is not None and before.channel != after.channel:
            key_before = f"{member.id}_{before.channel.id}"
            
            print(f"🔄 [移動] {member.display_name}: {before.channel.name} → {after.channel.name} ({time_str})")
            
            # ステップ1: 前のチャンネルの退出時間を更新
            print(f"   ステップ1: {before.channel.name} の退出時刻を記録")
            await update_leave_time_async(
                user_id=member.id,
                channel_name=before.channel.name,
                leave_time=time_str,
                display_name=member.display_name
            )
            
            if key_before in user_join_times:
                del user_join_times[key_before]
            
            # ステップ2: 新しいチャンネルへの入室を記録
            print(f"   ステップ2: {after.channel.name} への入室を記録")
            key_after = f"{member.id}_{after.channel.id}"
            user_join_times[key_after] = time_str
            
            await log_to_sheet_async(
                date=date,
                name=member.display_name,
                user_id=member.id,
                channel_name=after.channel.name,
                join_time=time_str,
                leave_time=""
            )
            
            print(f"✅ [移動完了] 両方の記録が完了しました")
        
    except Exception as e:
        print(f"❌ on_voice_state_update内でエラー発生: {e}")
        import traceback
        traceback.print_exc()

# Bot を起動
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
