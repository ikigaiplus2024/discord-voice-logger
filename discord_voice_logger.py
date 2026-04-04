import discord
from discord.ext import commands
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timezone, timedelta
import os
import json
import asyncio

SPREADSHEET_ID = os.environ.get('SPREADSHEET_ID')
DISCORD_TOKEN = os.environ.get('DISCORD_TOKEN')
CREDENTIALS_JSON = os.environ.get('CREDENTIALS_JSON')

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SHEET_NAME = 'VCログ'
JST = timezone(timedelta(hours=9))

write_queue = asyncio.Queue()

def append_to_sheet(row):
    creds_dict = json.loads(CREDENTIALS_JSON)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    gc = gspread.authorize(creds)
    sheet = gc.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
    sheet.append_row(row, value_input_option='USER_ENTERED')

async def sheet_writer():
    """キューから1件ずつ確実に書き込む"""
    while True:
        row = await write_queue.get()
        try:
            await asyncio.to_thread(append_to_sheet, row)
            print(f'書き込み成功: {row}')
        except Exception as e:
            print(f'書き込みエラー: {e} | データ: {row}')
        finally:
            write_queue.task_done()
intents = discord.Intents.default()
intents.voice_states = True
intents.members = True
bot = commands.Bot(command_prefix='!', intents=intents)
@bot.event
async def on_ready():
    print(f'{bot.user} としてログイン完了')
    print(f'{datetime.now(JST).strftime("%Y/%m/%d %H:%M:%S")}')
    asyncio.create_task(sheet_writer())
@bot.event
async def on_voice_state_update(member, before, after):
    if before.channel == after.channel:
        return
     now = datetime.now(JST)
    date_str = now.strftime('%Y/%m/%d')
    time_str = now.strftime('%H:%M')
    username = member.name
     if before.channel is None:
        event = 'Member joined voice channel'
        room1 = after.channel.name
        room2 = ''
    elif after.channel is None:
        event = 'Member left voice channel'
        room1 = before.channel.name
        room2 = ''
    else:
        event = 'Member moved channel'
        room1 = before.channel.name
        room2 = after.channel.name
     row = [date_str, time_str, username, event, room1, room2]
    await write_queue.put(row)
    print(f'{event}: {username} ({room1}{"->"+room2 if room2 else ""})')
 bot.run(DISCORD_TOKEN)
