import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment
import discord
from discord.ext import commands
import os

def transform_timetable(file_path, output_file_path):
    # 엑셀 파일을 로드합니다.
    df = pd.read_excel(file_path)

    # 날짜 문자열을 datetime으로 변환합니다.
    df['Exam start'] = pd.to_datetime(df['Exam start'], format='%Y. %m. %d. %H:%M:%S')
    df['Finish'] = pd.to_datetime(df['Finish'], format='%Y. %m. %d. %H:%M:%S')

    # 'Exam start' 열에서 요일과 시간을 추출합니다.
    df['Day'] = df['Exam start'].dt.day_name()
    df['Start Time'] = df['Exam start'].dt.strftime('%H:%M')
    df['End Time'] = df['Finish'].dt.strftime('%H:%M')

    # 'Start Time'과 'End Time'을 하나의 시간 슬롯 열로 결합합니다.
    df['Time Slot'] = df['Start Time'] + ' - ' + df['End Time']

    # 위치가 포함된 요약을 생성합니다.
    df['Summary with Location'] = df.apply(lambda x: f"{x['Summary']}\n{x['검색결과']}", axis=1)

    # 동일한 시간 슬롯 및 요일 내에서 중복 위치를 제거합니다.
    df['Summary with Location'] = df.groupby(['Time Slot', 'Day'])['Summary with Location'].transform(lambda x: '\n'.join(pd.Series(x).unique()))

    # 요일 순서를 정의합니다.
    day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']

    # 피벗 테이블을 생성하여 'Day'를 열로, 'Time Slot'을 인덱스로 사용하고, 요일을 정렬합니다.
    pivot_table = df.pivot_table(index='Time Slot', columns='Day', values='Summary with Location', aggfunc=lambda x: ' '.join(set(str(v) for v in x)))
    pivot_table = pivot_table[day_order]

    # 새 워크북을 만들고 피벗 테이블을 추가합니다.
    wb = Workbook()
    ws = wb.active

    # 헤더를 추가합니다.
    ws.append(['Time Slot'] + day_order)

    # 데이터 행을 추가합니다.
    for time_slot, row in pivot_table.iterrows():
        ws.append([time_slot] + list(row))

    # 정렬을 가운데로 조정합니다.
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')

    # 워크북을 저장합니다.
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