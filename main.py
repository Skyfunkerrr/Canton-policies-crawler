import pandas as pd
import re
import time
import random
import os
import sys
import tkinter as tk
from tkinter import messagebox
from collections import deque
from pathlib import Path
from DrissionPage import ChromiumPage, ChromiumOptions


def get_resource_path(relative_path):
    """获取资源文件的正确路径"""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def get_output_path(filename):
    """获取输出文件保存路径 - 保存到桌面"""
    output_dir = Path.home() / "Desktop" / "爬虫结果"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir / filename


def step2_filter(level="区级以上", gov="仅政府"):
    """Step2: 过滤网站"""
    print(f"\n[Step2] 开始过滤网站... 层级={level}, 政府={gov}")

    df = pd.read_excel(get_resource_path('数据/Step1_初筛网站_requests.xlsx'))
    print(f"[Step2] 原始数据: {len(df)} 条")

    if gov == "仅政府":
        df = df[df["title"].str.contains("政府", na=False)]
        df = df[df["title"].str.contains("广东省人民政府门户网站|市|区|街道", na=False,regex=True)]
        print(f"[Step2] 过滤后: {len(df)} 条")

    if level == "镇/街道":
        df = df[df["title"].str.contains("街道|镇", na=False, regex=True)]

    elif level == "区级以上":
        df = df[~df["title"].str.contains("街道|镇", na=False, regex=True)]

    print(f"[Step2] 最终: {len(df)} 条记录")
    return df


def step3_crawl(df_websites, keyword="城乡统筹", level="区级以上", gov="仅政府"):
    """Step3: 爬虫爬取文件 - 从DataFrame读取，返回结果DataFrame"""
    print(f"\n[Step3] 开始爬虫... 关键词={keyword}")
    print(f"[Step3] 读取 {len(df_websites)} 个网站，开启 Edge 浏览器...")

    try:
        option = ChromiumOptions()
        print(f"[Step3] ✓ 创建 ChromiumOptions")

        option.binary_location = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
        print(f"[Step3] ✓ 设置 Edge 路径")

        page = ChromiumPage(option)
        print(f"[Step3] ✓ 打开浏览器成功")

        result = []

        q = deque()
        for _, row in df_websites.iterrows():
            q.append({"url": row["url"], "title": row["title"]})

        print(f"[Step3] ✓ 队列初始化完成，共 {len(q)} 个网站")

        site_count = 0
        while q:
            task = q.popleft()
            url = task.get("url")
            title = task.get("title")

            if not url:
                print(f"[Step3] ⚠️ 跳过：URL为空")
                continue

            site_count += 1
            print(f"\n[爬虫] 处理网站 {site_count}/{len(df_websites)}: {title}")

            try:
                print(f"  [1] 正在访问网址: {url}")
                page.get(url)
                print(f"  [1] ✓ 网页加载成功")

                print(f"  [2] 正在查找搜索框...")
                search_box = page.ele("@id=input-keywords", timeout=5)
                search_button = page.ele("@class=list-search-button", timeout=5)

                if search_box and search_button:
                    print(f"  [2] ✓ 找到搜索框和按钮")
                    search_box.input(keyword)
                    print(f"  [2] ✓ 输入关键词: {keyword}")

                    search_button.click()
                    print(f"  [2] ✓ 点击搜索按钮")
                    time.sleep(2)
                else:
                    print(f"  [2] ⚠️ 未找到搜索框或按钮，跳过此网站")
                    continue

                print(f"  [3] 正在滚动到底部...")
                page.scroll.to_bottom()
                print(f"  [3] ✓ 滚动完成")

                print(f"  [4] 正在查找分页元素...")
                page_list = page.ele("@id=page-list", timeout=5)

                if not page_list:
                    print(f"  [4] ⚠️ 未找到分页元素")
                    continue

                totalpages = page_list.eles("tag:a", timeout=3)
                print(f"  [4] ✓ 找到 {len(totalpages)} 页分页链接")

                if len(totalpages) == 0:
                    print(f"  [4] ⚠️ 没有分页结果，跳过")
                    continue

                print(f"  [5] 正在查找列表项...")
                list_body = page.ele("@class=list-body", timeout=5)

                if not list_body:
                    print(f"  [5] ⚠️ 未找到列表容器")
                    continue

                list_items = list_body.eles("tag:div@class=list-item  file", timeout=5)
                print(f"  [5] ✓ 找到 {len(list_items)} 个列表项")

                for idx, list_item in enumerate(list_items):
                    try:
                        data_url = list_item.attr("data-url")
                        data_title_elem = list_item.ele("tag:a@class=title", timeout=2)

                        if data_title_elem:
                            data_title = data_title_elem.text
                            data_title = data_title.replace("<em>", "").replace("</em>", "").strip()
                            print(f"    ✓ {data_title}")
                            result.append([title, data_title, data_url])
                    except Exception as e:
                        print(f"    ⚠️ 解析列表项 {idx} 失败: {str(e)}")
                        continue

                if len(totalpages) > 1:
                    print(f"  [6] 开始翻页处理（共 {len(totalpages)} 页）...")
                    num = 1
                    while num < len(totalpages):
                        try:
                            print(f"    [6.{num}] 翻页中...")
                            page.scroll.to_bottom()

                            page_button_current = page.ele("@id=page-list", timeout=5).eles("tag:a@class=item cur",
                                                                                            timeout=3)
                            if not page_button_current:
                                print(f"    [6.{num}] ⚠️ 未找到当前页按钮")
                                break

                            page_button = page_button_current[0].nexts("tag:a@class=item", timeout=3)
                            if not page_button:
                                print(f"    [6.{num}] ⚠️ 未找到下一页按钮")
                                break

                            page_button[0].click()
                            print(f"    [6.{num}] ✓ 点击下一页")
                            time.sleep(2)

                            next_list_body = page.ele("@class=list-body", timeout=5)
                            next_list_items = next_list_body.eles("tag:div@class=list-item  file", timeout=5)
                            print(f"    [6.{num}] ✓ 找到 {len(next_list_items)} 个列表项")

                            for next_list_item in next_list_items:
                                try:
                                    data_url = next_list_item.attr("data-url")
                                    data_title_elem = next_list_item.ele("tag:a@class=title", timeout=2)

                                    if data_title_elem:
                                        data_title = data_title_elem.text
                                        data_title = data_title.replace("<em>", "").replace("</em>", "").strip()
                                        print(f"      ✓ {data_title}")
                                        result.append([title, data_title, data_url])
                                except Exception as e:
                                    print(f"      ⚠️ 解析失败: {str(e)}")
                                    continue
                            num += 1
                        except Exception as e:
                            print(f"    [6.{num}] ⚠️ 翻页处理失败: {str(e)}")
                            break

                print(f"  ✓ 本网站处理完成")
                time.sleep(random.randint(1, 2))

            except Exception as e:
                import traceback
                error_detail = traceback.format_exc()
                print(f"  ❌ 处理失败: {str(e)}")
                print(f"  错误详情:\n{error_detail}")
                continue

        print(f"\n[Step3] 正在关闭浏览器...")
        page.quit()
        print(f"[Step3] ✓ 浏览器已关闭")

        df_result = pd.DataFrame(result, columns=["数据源", "title", "url"])
        print(f"\n[Step3] 完成！爬取 {len(df_result)} 条文件")
        return df_result

    except Exception as e:
        import traceback
        error_detail = traceback.format_exc()
        print(f"\n[Step3] ❌ 致命错误: {str(e)}")
        print(f"[Step3] 错误详情:\n{error_detail}")
        raise


def step4_filter_title(df_crawled, keyword, level="区级以上", gov="仅政府"):
    """Step4: 过滤标题 - 保存到桌面"""
    print(f"\n[Step4] 开始过滤标题... 关键词={keyword}")

    result = []
    for i in range(len(df_crawled)):
        result.append([df_crawled.loc[i, "数据源"], df_crawled.loc[i, "title"], df_crawled.loc[i, "url"]])

    result_df = pd.DataFrame(result, columns=["数据源", "title", "url"])
    output_path = get_output_path(f"Step4_{keyword}_{level}_{gov}文件网站.xlsx")
    result_df.to_excel(output_path, index=False)
    print(f"[Step4] 完成！过滤后 {len(result)} 条记录")
    print(f"[Step4] 📁 保存到: {output_path}")
    return result_df


def get_config():
    """显示配置界面"""
    root = tk.Tk()
    root.title("广东省政策文件爬虫")
    root.geometry("300x350")

    tk.Label(root, text="请选择文件效力层级", font=("Arial", 12, "bold")).pack(pady=10)
    level_var = tk.StringVar(value="区级以上")
    tk.Radiobutton(root, text="区级以上", variable=level_var, value="区级以上").pack(anchor=tk.W, padx=30)
    tk.Radiobutton(root, text="镇/街道", variable=level_var, value="镇/街道").pack(anchor=tk.W, padx=30)
    tk.Radiobutton(root, text="所有层级", variable=level_var, value="所有层级").pack(anchor=tk.W, padx=30)

    tk.Label(root, text="是否仅政府", font=("Arial", 12, "bold")).pack(pady=10)
    gov_var = tk.StringVar(value="仅政府")
    tk.Radiobutton(root, text="仅政府", variable=gov_var, value="仅政府").pack(anchor=tk.W, padx=30)
    tk.Radiobutton(root, text="所有机关", variable=gov_var, value="所有机关").pack(anchor=tk.W, padx=30)

    tk.Label(root, text="输入搜索关键词", font=("Arial", 12, "bold")).pack(pady=10)
    keyword_entry = tk.Entry(root, width=30)
    keyword_entry.insert(0, "城乡统筹")
    keyword_entry.pack(pady=5)

    def on_ok():
        config = {
            "level": level_var.get(),
            "gov": gov_var.get(),
            "keyword": keyword_entry.get().strip() or "城乡统筹"
        }
        root.config_data = config
        root.destroy()

    tk.Button(root, text="确定", command=on_ok, width=15).pack(pady=15)
    root.mainloop()
    return getattr(root, "config_data", None)


if __name__ == "__main__":
    config = get_config()
    if config is None:
        exit()

    try:
        print("=" * 60)
        print("广东省政策文件爬虫 - 开始运行")
        print("=" * 60)

        df_step2 = step2_filter(config["level"], config["gov"])
        df_step3 = step3_crawl(df_step2, config["keyword"], config["level"], config["gov"])
        df_step4 = step4_filter_title(df_step3, config["keyword"], config["level"], config["gov"])

        print("\n" + "=" * 60)
        messagebox.showinfo("完成",
                            f"✓ 全部任务完成！\n\n层级: {config['level']}\n关键词: {config['keyword']}\n\n最终结果: {len(df_step4)} 条记录\n\n文件已保存到桌面的'爬虫结果'文件夹")
        print("=" * 60)
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        print(f"\n[错误] {error_msg}")
        messagebox.showerror("错误", f"执行出错:\n{str(e)}")
