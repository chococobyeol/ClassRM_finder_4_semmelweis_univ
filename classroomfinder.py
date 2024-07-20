import discord
from discord.ext import commands
import aiohttp
from bs4 import BeautifulSoup
import logging
import pandas as pd
import os
import re

# 로깅 설정
logging.basicConfig(level=logging.INFO)

# 강의실 정보를 가져오는 함수
async def fetch_classroom_info(classroom_name):
    try:
        search_url = f"https://semmelweis.hu/registrar/information/classroom-finder/?search={classroom_name}"
        
        async with aiohttp.ClientSession() as session:
            async with session.get(search_url) as res:
                if res.status != 200:
                    logging.error(f"HTTP 요청 오류: 상태 코드 {res.status}")
                    return None
                html = await res.text()
        
        soup = BeautifulSoup(html, 'html.parser')
        
        # 모든 검색 결과 행을 찾습니다.
        results = soup.select("#tablepress-16 > tbody > tr")
        
        classroom_info_list = []
        for result in results:
            cells = result.select("td")
            if len(cells) >= 2:
                classroom_name_cell = cells[0].get_text(strip=True)
                classroom_location_cell = cells[1].get_text(strip=True)
                if classroom_name.lower().replace(' ', '') in classroom_name_cell.lower().replace(' ', ''):
                    classroom_info_list.append(f"{classroom_name_cell}: {classroom_location_cell}")
        
        if not classroom_info_list:
            return f"'{classroom_name}'에 대한 검색 결과를 찾을 수 없습니다."
        
        classroom_info = "; ".join(classroom_info_list)
        return classroom_info
    
    except Exception as e:
        logging.error(f"강의실 정보를 가져오는 중 오류 발생: {e}")
        return None

# 결과를 2000자 이하의 메시지로 분할하여 전송
async def send_long_message(ctx, message):
    for i in range(0, len(message), 2000):
        await ctx.send(message[i:i+2000])

# 엑셀 파일을 처리하는 함수
async def process_excel(file_path):
    df = pd.read_excel(file_path)
    return df

# 검색 결과를 엑셀 파일로 저장하는 함수
async def save_results_to_excel(df, output_path):
    df.to_excel(output_path, index=False)

# 엑셀 파일 검색 명령어
@commands.command(name="엑셀검색")
async def search_classroom_excel(ctx):
    if not ctx.message.attachments:
        await ctx.send("엑셀 파일을 첨부해주세요.")
        return
    
    attachment = ctx.message.attachments[0]
    if not attachment.filename.endswith('.xlsx'):
        await ctx.send("지원되는 파일 형식은 .xlsx 입니다.")
        return

    file_path = f"/tmp/{attachment.filename}"
    await attachment.save(file_path)
    await ctx.send(f"엑셀 파일 '{attachment.filename}'을 처리 중입니다...")

    df = await process_excel(file_path)
    results = []
    for location in df['Location']:
        if pd.isna(location):
            results.append("정보를 찾을 수 없습니다.")
            continue
        
        location = str(location).strip()  # 각 개별 검색어도 정리
        search_result = await fetch_classroom_info(location)
        results.append(search_result if search_result else "정보를 찾을 수 없습니다.")

    df['검색결과'] = results

    output_path = f"/tmp/검색결과_{attachment.filename}"
    await save_results_to_excel(df, output_path)

    await ctx.send(file=discord.File(output_path))

    # 임시 파일 삭제
    os.remove(file_path)
    os.remove(output_path)

# 강의실 검색 명령어
@commands.command(name="검색")
async def search_classroom(ctx, *, classroom_names: str):
    await ctx.send(f"강의실 '{classroom_names}' 정보를 검색 중입니다...")

    # 검색어에서 불필요한 공백과 서식을 제거
    cleaned_classroom_names = classroom_names.replace('\n', '').replace('\r', '').replace('\t', '').strip()
    
    results = []
    for classroom_name in cleaned_classroom_names.split('/'):
        classroom_name = classroom_name.strip()  # 각 개별 검색어도 정리
        location = await fetch_classroom_info(classroom_name)
        if location:
            results.append(f"강의실 '{classroom_name}'의 위치:\n{location}")
        else:
            results.append(f"강의실 '{classroom_name}' 정보를 찾을 수 없습니다.")
    
    final_result = "\n\n".join(results)
    await send_long_message(ctx, final_result)

# setup 함수: 봇의 설정을 초기화합니다.
def setup(bot):
    bot.add_command(search_classroom)
    bot.add_command(search_classroom_excel)