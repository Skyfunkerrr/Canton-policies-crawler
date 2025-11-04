# 文件名: Step1_crawl.py
# 依赖: pip install requests pandas openpyxl
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import pandas as pd
import os
import sys


def get_resource_path(relative_path):
    """获取资源文件的正确路径"""
    if getattr(sys, 'frozen', False):
        # 被编译成 EXE
        base_path = sys._MEIPASS
    else:
        # 开发环境
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)


# 配置
TOTAL_IDS = 1000000
MAX_WORKERS = 50
TIMEOUT = 5  # 单次请求超时（秒）
OUTPUT_FILE = get_resource_path("数据/Step1_初筛网站_requests.xlsx")
# 每次批量提交的任务数（防止一次性提交过多 future 导致内存占用过大）
BATCH_SIZE = 20000  # 可根据内存和并发调整

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Connection": "keep-alive",
    "Referer": "https://search.gd.gov.cn/"
}


# 创建带重试的 Session
def make_session():
    s = requests.Session()
    retries = Retry(total=2, backoff_factor=0.3,
                    status_forcelist=(500, 502, 503, 504),
                    allowed_methods=frozenset(["GET"]))
    s.mount("https://", HTTPAdapter(max_retries=retries))
    s.mount("http://", HTTPAdapter(max_retries=retries))
    s.headers.update(HEADERS)
    return s


def check_page(i, session=None):
    """
    检查单个 id 页面，返回 [id, title, url] 或 None
    session: requests.Session 实例（若无则内部创建短期 session）
    """
    url = f"https://search.gd.gov.cn/search/file/{i}"
    close_session = False
    if session is None:
        session = make_session()
        close_session = True
    try:
        resp = session.get(url, timeout=TIMEOUT)
        if resp.status_code == 200:
            m = re.search(r"<title>(.*?)</title>", resp.text, flags=re.IGNORECASE | re.S)
            if m:
                title = m.group(1).strip()
                if "测试" not in title:
                    print(f"找到有效页面: {i} - {title}")
                    return [i, title, url]


    except Exception:
        # 可以在这里打印异常以便调试
        # import traceback; traceback.print_exc()
        pass
    finally:
        if close_session:
            session.close()
    return None


def main():
    data = []

    start_id = 1
    end_id = TOTAL_IDS

    # 分批提交，避免一次性生成过多 future
    current = start_id
    checked_count = 0
    found_count = 0

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        while current <= end_id:
            batch_end = min(current + BATCH_SIZE - 1, end_id)
            futures = {executor.submit(check_page, i): i for i in range(current, batch_end + 1)}
            for idx, future in enumerate(as_completed(futures), 1):
                checked_count += 1
                result = future.result()
                if result:
                    data.append(result)
                    found_count += 1

                if checked_count % 1000 == 0:
                    print(f"已检查 {checked_count} 个页面，找到 {found_count} 个有效页面")

            # 下一批
            current = batch_end + 1

            # time.sleep(0.1)

    print("Step1 Finished")
    return data


if __name__ == "__main__":
    data = main()
    df = pd.DataFrame(data, columns=['id', 'title', 'url'])
    df.to_excel(OUTPUT_FILE, index=False)
    print("完成，结果保存到", OUTPUT_FILE)
