import re
from openpyxl import Workbook
import os
import datetime

# 获取系统当前日期和时间
now = datetime.datetime.now()
date_str = now.strftime('%Y%m%d%H%M')

# 获取桌面路径
desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')

# 修改代码以接收拖放的文件名
if __name__ == "__main__":
    import sys
    # 检查是否有文件参数传入
    if len(sys.argv) > 1:
        # 获取拖放的文件路径
        txt_file_path = sys.argv[1]
        # 检查文件扩展名是否正确
        if txt_file_path.lower().endswith('.txt'):
            # 创建一个新的Excel工作簿
            wb = Workbook()
            ws = wb.active
            ws.title = "视频数据"

            # 定义表头
            headers = ["视频标题", "发布日期", "观看量", "人均观看时长", "点赞量", "收藏数", "评论数", "弹幕数", "分享数", "直接涨粉数"]
            ws.append(headers)

            # 正则表达式，用于匹配数据
            title_pattern = re.compile(r'<span data-v-70715371="" class="title">(.*?)</span>')
            publish_time_pattern = re.compile(r'<span data-v-70715371="" class="publish-time">.*?(\d{4}-\d{2}-\d{2})</span>')
            watch_count_pattern = re.compile(r'<label data-v-70715371="">观看量</label> <b data-v-70715371="" class="align-text">(.*?)</b>')
            avg_watch_time_pattern = re.compile(r'<label data-v-70715371="">人均观看时长</label> <b data-v-70715371="">(.*?)</b>')
            like_count_pattern = re.compile(r'<label data-v-70715371="">点赞量</label> <b data-v-70715371="" class="align-text">(.*?)</b>')
            favorite_count_pattern = re.compile(r'<label data-v-70715371="">收藏数</label> <b data-v-70715371="" class="align-text">(.*?)</b>')
            comment_count_pattern = re.compile(r'<label data-v-70715371="">评论数</label> <b data-v-70715371="" class="align-text-m">(.*?)</b>')
            danmu_count_pattern = re.compile(r'<label data-v-70715371="">弹幕数</label> <b data-v-70715371="" class="align-text-m">(.*?)</b>')
            share_count_pattern = re.compile(r'<label data-v-70715371="">分享数</label> <b data-v-70715371="" class="align-text-m">(.*?)</b>')
            direct_fans_count_pattern = re.compile(r'<label data-v-70715371="">直接涨粉数</label> <b data-v-70715371="">(.*?)</b>')

            # 读取txt文件
            with open(txt_file_path, 'r', encoding='utf-8') as file:
                content = file.read()

           # 查找所有匹配的数据
            titles = title_pattern.findall(content)
            publish_times = publish_time_pattern.findall(content)
            watch_counts = watch_count_pattern.findall(content)
            avg_watch_times = avg_watch_time_pattern.findall(content)
            like_counts = like_count_pattern.findall(content)
            favorite_counts = favorite_count_pattern.findall(content)
            comment_counts = comment_count_pattern.findall(content)
            danmu_counts = danmu_count_pattern.findall(content)
            share_counts = share_count_pattern.findall(content)
            direct_fans_counts = direct_fans_count_pattern.findall(content)

# 将找到的数据汇总到Excel中
for i in range(len(titles)):
    ws.append([
        titles[i], publish_times[i], watch_counts[i], avg_watch_times[i],
        like_counts[i], favorite_counts[i], comment_counts[i], danmu_counts[i],
        share_counts[i], direct_fans_counts[i]
    ])


            # 构建Excel文件名
excel_file_name = f"{date_str}_video_data.xlsx"
excel_file_path = os.path.join(desktop_path, excel_file_name)

            # 保存Excel文件到桌面
wb.save(excel_file_path)

            # 打印日志提醒用户成功导出文件
print(f"文件已成功导出至桌面: '{excel_file_path}'。")
