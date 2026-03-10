"""
Core
Made with ❤️by Z🐻
"""
import io
import os
import logging
import time
from typing import Optional, Dict, Any
import requests
from bs4 import BeautifulSoup
from PIL import Image
from pytesseract import image_to_string
import re


class ResidencePointsCrawler:
    def __init__(self, session_id: Optional[str] = None, max_retries: int = 20):
        self.session = requests.Session()
        self.max_retries = max_retries
        self.logger = logging.getLogger(__name__)

        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8'
        }

        if session_id:
            self.headers['Cookie'] = f'JSESSIONID={session_id}'

        self.POST_URL = 'https://jzzjf.rsj.sh.gov.cn/jzzjf/pingfen/query_jzz_score.jsp'
        self.CAPCHA_URL = 'https://jzzjf.rsj.sh.gov.cn/jzzjf/inc/code.jsp?Math.random()'

    def _recognize_captcha(self, image_bytes: bytes) -> Optional[str]:
        try:
            captcha_img = Image.open(io.BytesIO(image_bytes))

            captcha_img = captcha_img.resize((300, 100))
            captcha_img = captcha_img.convert("L")

            threshold = 127
            captcha_img = captcha_img.point(lambda p: p > threshold and 255)

            captcha = image_to_string(captcha_img, config='--psm 8 --oem 3')
            captcha = captcha.strip().replace(" ", "")

            if len(captcha) == 4:
                return captcha
            else:
                self.logger.debug(f"No matched captcha: {captcha}")
                return None

        except Exception as e:
            self.logger.error(f"Failed recognize captcha: {str(e)}")
            return None


    def _get_captcha(self) -> Optional[str]:
        try:
            response = self.session.get(self.CAPCHA_URL, headers=self.headers, timeout=10)
            response.raise_for_status()
            return self._recognize_captcha(response.content)
        except requests.RequestException as e:
            self.logger.error(f"Get captcha failed: {str(e)}")
            return None


    def query_points(self, name: str, pid: str, progress_callback=None) -> Dict[str, Any]:
        result = {
            'status': 'failed',
            'data': None,
            'error': None,
            'attempts': 0
        }

        for attempt in range(1, self.max_retries + 1):
            result['attempts'] = attempt

            try:
                if progress_callback:
                    progress_callback(attempt, self.max_retries, f"Try the {attempt}/{self.max_retries} time...")

                captcha = self._get_captcha()
                if not captcha:
                    self.logger.debug(f"The {attempt} try: captcha failed.")
                    time.sleep(0.5)
                    continue

                self.logger.debug(f"The {attempt} try: captcha result is {captcha}")

                form_data = {
                    'PersonName': name,
                    'Pid': pid,
                    'YanZheng': captcha,
                    'button1': '提交查询'
                }

                response = self.session.post(
                    self.POST_URL,
                    headers=self.headers,
                    data=form_data,
                    timeout=10
                )
                response.raise_for_status()

                if '验证码输入错误' in response.text:
                    self.logger.debug(f"The {attempt} try: captcha failed.")
                    time.sleep(0.5)
                    continue

                if '未查询到相关记录' in response.text:
                    result['status'] = 'not_found'
                    result['error'] = '未查询到相关记录，请检查姓名和身份证号是否正确'
                    self.logger.warning(f"Not found record for {name}")
                    return result

                try:
                    bs = BeautifulSoup(response.content, 'html.parser')
                    table = bs.find('table').find('table')

                    if table:
                        pretty_text = re.sub(r'\s+', ' ', table.text).strip()

                        parsed_data = self._parse_result_table(table)

                        result['status'] = 'success'
                        result['data'] = {
                            'raw_text': pretty_text,
                            'parsed_data': parsed_data
                        }

                        self.logger.info(f"Successfully query for{name}. (The {attempt} try)")
                        return result
                    else:
                        result['error'] = '页面结构错误，无法解析'
                        self.logger.error("No found result table")

                except Exception as e:
                    result['error'] = f'解析失败 {str(e)}'
                    self.logger.error("Failed to parse result table")

            except requests.Timeout:
                self.logger.warning(f"The {attempt} try failed: Time out")
                result['error'] = f'网络请求失败 {str(e)}'
                time.sleep(1)

            except requests.RequestException:
                self.logger.warning(f"The {attempt} try failed: Request failed")
                result['error'] = f'网络请求失败 {str(e)}'
                time.sleep(1)

        if result['status'] == 'failed' and not result['error']:
            result['error'] = f'查询失败，已经尝试 {self.max_retries} 次'

        self.logger.error(f"Failed to query for {name}, reach maximum tries {self.max_retries}")
        return result

    def _parse_result_table(self, table) -> Dict[str, str]:
        parsed_data = {}

        try:
            rows = table.find_all('tr')
            for row in rows:
                cells = row.find_all(['td', 'th'])
                if len(cells) >= 2:
                    for i in range(0, len(cells) - 1, 2):
                        key = cells[i].get_text(strip=True)
                        value = cells[i + 1].get_text(strip=True)
                        if key and value:
                            parsed_data[key] = value
        except Exception as e:
            self.logger.error(f"Failed parsed the data in table: {str(e)}")

        return parsed_data

    def close(self):
        self.session.close()
