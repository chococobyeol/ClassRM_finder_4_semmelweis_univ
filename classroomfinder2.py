import discord
from discord.ext import commands
import aiohttp
from bs4 import BeautifulSoup
import logging
from icalendar import Calendar
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
                    return None, None, None, None
                html = await res.text()
        
        soup = BeautifulSoup(html, 'html.parser')
        
        # 모든 검색 결과 행을 찾습니다.
        results = soup.select("#tablepress-16 > tbody > tr")
        
        for result in results:
            department = result.select_one("td.column-1").get_text(strip=True)
            address = result.select_one("td.column-2").get_text(strip=True)
            if classroom_name.lower().replace(' ', '') in department.lower().replace(' ', ''):
                # 주소에서 숫자 앞부분을 잘라냅니다.
                address_match = re.search(r'\d.*', address)
                address_cleaned = address_match.group() if address_match else address
                
                # department에서 강의실 코드와 세부 정보를 분리합니다.
                department_parts = department.split(' - ', 1)
                classroom_code = department_parts[0]
                classroom_details = department_parts[1] if len(department_parts) > 1 else ""
                
                # 주소 부분을 제거하고 순수한 부서 이름만 추출합니다.
                pure_department = address.split(',')[0].strip()
                
                return classroom_code, classroom_details, pure_department, address_cleaned
        
        return None, None, None, None
    
    except Exception as e:
        logging.error(f"강의실 정보를 가져오는 중 오류 발생: {e}")
        return None, None, None, None

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
                        classroom_code, classroom_details, pure_department, address_cleaned = await fetch_classroom_info(location)
                        if address_cleaned:
                            new_description = f"{classroom_code} - {classroom_details}\nDepartment: {pure_department}"
                            component['LOCATION'] = address_cleaned
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