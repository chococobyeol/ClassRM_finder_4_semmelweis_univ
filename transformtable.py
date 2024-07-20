# transformtable.py
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment
import discord
from discord.ext import commands
import os

def transform_timetable(file_path, output_file_path):
    # Load the Excel file
    df = pd.read_excel(file_path)

    # Convert the date strings to datetime
    df['Exam start'] = pd.to_datetime(df['Exam start'], format='%Y. %m. %d. %H:%M:%S')
    df['Finish'] = pd.to_datetime(df['Finish'], format='%Y. %m. %d. %H:%M:%S')

    # Extract the day of the week and time from the 'Exam start' column
    df['Day'] = df['Exam start'].dt.day_name()
    df['Start Time'] = df['Exam start'].dt.strftime('%H:%M')
    df['End Time'] = df['Finish'].dt.strftime('%H:%M')

    # Combine 'Start Time' and 'End Time' to a single time slot column
    df['Time Slot'] = df['Start Time'] + ' - ' + df['End Time']

    # Create a combined summary with location
    df['Summary with Location'] = df.apply(lambda x: f"{x['Summary']}\n{x['검색결과']}", axis=1)

    # Remove duplicate locations within the same time slot and day
    df['Summary with Location'] = df.groupby(['Time Slot', 'Day'])['Summary with Location'].transform(lambda x: '\n'.join(pd.Series(x).unique()))

    # Define the order of days
    day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']

    # Create a pivot table with 'Day' as columns and 'Time Slot' as index, ordering the days
    pivot_table = df.pivot_table(index='Time Slot', columns='Day', values='Summary with Location', aggfunc=lambda x: ' '.join(set(str(v) for v in x)))
    pivot_table = pivot_table[day_order]

    # Create a new workbook and add the pivot table to it
    wb = Workbook()
    ws = wb.active

    # Add the headers
    ws.append(['Time Slot'] + day_order)

    # Add the data rows
    for time_slot, row in pivot_table.iterrows():
        ws.append([time_slot] + list(row))

    # Adjust the alignment to center
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')

    # Save the workbook
    wb.save(output_file_path)

    print(f"Timetable saved to {output_file_path}")

# 디스코드 명령어 핸들러
class TimetableConverter(commands.Cog):
    def __init__(self, bot):
        self.bot = bot

    @commands.command(name="시간표변환")
    async def convert_timetable(self, ctx):
        if not ctx.message.attachments:
            await ctx.send("엑셀 파일을 첨부해주세요.")
            return

        attachment = ctx.message.attachments[0]
        if not attachment.filename.endswith('.xlsx'):
            await ctx.send("지원되는 파일 형식은 .xlsx 입니다.")
            return

        file_path = f"/tmp/{attachment.filename}"
        output_file_path = f"/tmp/변환된_{attachment.filename}"

        await attachment.save(file_path)
        await ctx.send(f"엑셀 파일 '{attachment.filename}'을 변환 중입니다...")

        try:
            transform_timetable(file_path, output_file_path)
            await ctx.send(file=discord.File(output_file_path))
        except Exception as e:
            await ctx.send(f"시간표 변환 중 오류가 발생했습니다: {e}")
        finally:
            os.remove(file_path)
            os.remove(output_file_path)

# setup 함수
async def setup(bot):
    await bot.add_cog(TimetableConverter(bot))