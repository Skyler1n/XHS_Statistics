import re
from openpyxl import Workbook
import os
import datetime

# 获取系统当前日期和时间
now = datetime.datetime.now()
date_str = now.strftime('%Y%m%d%H%M')

# 获取桌面路径
desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')

# 接收拖放的文件名
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
            headers = ["作品名称", "发布时间", "审核状态", "观看量", "人均观看时长", "点赞量", "收藏数", "评论数", "分享数", "直接涨粉数", "弹幕数"]
            ws.append(headers)

            # 正则表达式，用于匹配数据
            title_pattern = re.compile(r'<span data-v-70715371="" class="title">(.*?)</span>')
            permission_pattern = re.compile(r'class="play">(.*?)</div> <div data-v-70715371="" class="info-text"><span data-v-70715371="" class="title">')
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


            # 转换包含“万”的字符串为数字
            def convert_to_number(text):
                if '万' in text:
                    # 去掉'万'并转换为数字
                    return int(float(text.replace('万', '')) * 10000)
                else:
                    return int(text)

           # 查找所有匹配的数据
            permissions = permission_pattern.findall(content)
            titles = title_pattern.findall(content)
            publish_times = publish_time_pattern.findall(content)
            watch_counts = [convert_to_number(count) for count in watch_count_pattern.findall(content)]
            avg_watch_times = avg_watch_time_pattern.findall(content)
            like_counts = [convert_to_number(count) for count in like_count_pattern.findall(content)]
            favorite_counts = [convert_to_number(count) for count in favorite_count_pattern.findall(content)]
            comment_counts = [convert_to_number(count) for count in comment_count_pattern.findall(content)]
            danmu_counts = [convert_to_number(count) for count in danmu_count_pattern.findall(content)]
            share_counts = [convert_to_number(count) for count in share_count_pattern.findall(content)]
            direct_fans_counts = [convert_to_number(count) for count in direct_fans_count_pattern.findall(content)]

            # 处理审查状态
            for i in range(len(permissions)):
                if permissions[i] == " <!---->":
                    permissions[i] = "公开"
                else:
                    permissions[i] = "仅自己可见"

            # 将找到的数据汇总到Excel中
            for i in range(len(titles)):
                ws.append([
                    titles[i], publish_times[i], permissions[i], watch_counts[i], avg_watch_times[i],
                    like_counts[i], favorite_counts[i], comment_counts[i],share_counts[i], 
                    direct_fans_counts[i], danmu_counts[i]
                ])

            # 构建Excel文件名
            excel_file_name = f"{date_str}_video_data.xlsx"
            excel_file_path = os.path.join(desktop_path, excel_file_name)

            # 保存Excel文件到桌面
            wb.save(excel_file_path)

            # 打印日志提醒用户成功导出文件
            print(f"文件已成功导出至桌面: '{excel_file_path}'。")
