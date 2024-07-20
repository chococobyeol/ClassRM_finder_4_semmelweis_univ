import discord
from discord.ext import commands
import aiohttp
from bs4 import BeautifulSoup
import logging
from icalendar import Calendar, Event
from datetime import datetime
import os

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
            department = result.select_one("td.column-2 > div > div.class-department").get_text(strip=True)
            address = result.select_one("td.column-2 > div > div.class-address").get_text(strip=True)
            classroom_info_list.append(f"{department}: {address}")
        
        if not classroom_info_list:
            return None, None
        
        # 첫 번째 결과만 사용합니다.
        return classroom_info_list[0].split(": ")
    
    except Exception as e:
        logging.error(f"강의실 정보를 가져오는 중 오류 발생: {e}")
        return None, None

# 디스코드 명령어 핸들러
class CalendarSearcher(commands.Cog):
    def __init__(self, bot):
        self.bot = bot

    @commands.command(name="캘린더검색")
    async def search_calendar(self, ctx):
        if not ctx.message.attachments:
            await ctx.send("ics 파일을 첨부해주세요.")
            return

        attachment = ctx.message.attachments[0]
        if not attachment.filename.endswith('.ics'):
            await ctx.send("지원되는 파일 형식은 .ics 입니다.")
            return

        file_path = f"/tmp/{attachment.filename}"
        output_file_path = f"/tmp/변환된_{attachment.filename}"

        await attachment.save(file_path)
        await ctx.send(f"ics 파일 '{attachment.filename}'을 처리 중입니다...")

        try:
            with open(file_path, 'rb') as f:
                cal = Calendar.from_ical(f.read())

            new_cal = Calendar()

            for component in cal.walk():
                if component.name == "VEVENT":
                    location = component.get('LOCATION')
                    if location:
                        department, address = await fetch_classroom_info(location)
                        if address:
                            new_description = f"{location}\nDepartment: {department}\nAddress: {address}"
                            component['LOCATION'] = address
                            component['DESCRIPTION'] = new_description
                    new_cal.add_component(component)

            with open(output_file_path, 'wb') as f:
                f.write(new_cal.to_ical())

            await ctx.send(file=discord.File(output_file_path))
        except Exception as e:
            await ctx.send(f"캘린더 검색 및 변환 중 오류가 발생했습니다: {e}")
        finally:
            os.remove(file_path)
            os.remove(output_file_path)

# setup 함수
async def setup(bot):
    await bot.add_cog(CalendarSearcher(bot))