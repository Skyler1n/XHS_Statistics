import re
from openpyxl import Workbook
import os
import datetime
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import configparser

# 获取系统当前日期和时间
now = datetime.datetime.now()
date_str = now.strftime('%Y%m%d%H%M')

# 获取桌面路径
#desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')

# 临时路径
temp_dir = os.path.join(os.environ['TEMP'], 'xhsexp')

# 确保 xhsexp 目录存在
if not os.path.exists(temp_dir):
    os.makedirs(temp_dir)

# 配置文件名
config_file = os.path.join(temp_dir, 'config.ini')

# 读取配置文件
config = configparser.ConfigParser()
if os.path.exists(config_file):
    config.read(config_file)
    output_path = config.get('PATH', 'output_path', fallback=os.path.join(os.environ['USERPROFILE'], 'Desktop'))
    show_popup = config.getboolean('POPUP', 'show_popup', fallback=True)
else:
    output_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    show_popup = True
    config['PATH'] = {'output_path': output_path}
    config['POPUP'] = {'show_popup': show_popup}
    with open(config_file, 'w') as configfile:
        config.write(configfile)

# 有关GUI
if __name__ == "__main__":
    import sys

    def save_settings():
        config['PATH']['output_path'] = output_path_var.get()
        config['POPUP']['show_popup'] = str(show_popup_var.get())
        with open(config_file, 'w') as configfile:
            config.write(configfile)
        messagebox.showinfo("保存设置", "所有设置已保存。")

    def restore_defaults():
        global output_path  # Declare that we're modifying the global variable
        global show_popup
        output_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        show_popup = True
        output_path_var.set(output_path)
        show_popup_var.set(show_popup)
        if os.path.exists(config_file):
            os.remove(config_file)
        save_settings()
        messagebox.showinfo("恢复默认", "所有设置已恢复默认。")


    def browse_output_path():
        directory = filedialog.askdirectory()
        if directory:
            output_path_var.set(os.path.normpath(directory))

    # 创建GUI
    root = tk.Tk()
    root.title("小红书视频数据导出工具-控制面板")

    # 创建选项卡
    tab_control = tk.ttk.Notebook(root)
    tab_path = tk.ttk.Frame(tab_control)
    tab_popup = tk.ttk.Frame(tab_control)
    tab_about = tk.ttk.Frame(tab_control)
    tab_control.add(tab_path, text='路径设置')
    tab_control.add(tab_popup, text='弹窗设置')
    tab_control.add(tab_about, text='关于作者')
    tab_control.pack(expand=1, fill="both")

    # 路径选项卡
    output_path_var = tk.StringVar(value=output_path)
    tk.Label(tab_path, text="输出路径:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    output_path_entry = tk.Entry(tab_path, textvariable=output_path_var, width=50)
    output_path_entry.grid(row=0, column=1, padx=5, pady=5)
    browse_button = tk.Button(tab_path, text="选择路径", command=browse_output_path)
    browse_button.grid(row=0, column=2, padx=5, pady=5)
    #tk.Button(tab_path, text="恢复默认路径", command=restore_defaults).grid(row=1, column=1, columnspan=2, sticky="we", padx=5, pady=5)

    # 设置grid的列权重，使得输入框可以扩大
    tab_path.grid_columnconfigure(1, weight=1)

    # 弹窗选项卡
    show_popup_var = tk.BooleanVar(value=show_popup)
    tk.Checkbutton(tab_popup, text="导出成功后弹窗提示", variable=show_popup_var).pack(fill="x", padx=5, pady=5)

    # 作者信息选项卡
    tk.Label(tab_about, text="软件版本：V4.0").pack(padx=5, pady=5)
    tk.Label(tab_about, text="软件编写：Skyler1n").pack(padx=5, pady=5)

    # 更新日志按钮
    def show_changelog():
        changelog = """ V1.0：实现基本功能。\n V2.0：实现txt拖拽到程序直接生成xls。\n V3.0：增加审核状态。\n V3.1：将XX换成纯数字以方便统计。\n V3.2：增加xls导出成功提示。\n V4.0：全新GUI界面，支持设置导出路径和是否进行弹窗。\n"""
        messagebox.showinfo("更新日志", changelog)

    tk.Button(tab_about, text="更新日志", command=show_changelog).pack(padx=5, pady=5)

    # 保存和恢复默认按钮
    tk.Button(root, text="保存所有设置", command=save_settings).pack(side=tk.LEFT, padx=5, pady=5)
    tk.Button(root, text="全部恢复默认", command=restore_defaults).pack(side=tk.RIGHT, padx=5, pady=5)

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
            excel_file_path = os.path.join(output_path_var.get(), excel_file_name)
            #excel_file_path = os.path.join(desktop_path, excel_file_name)


            # 保存Excel文件到桌面
            wb.save(excel_file_path)
            
            # 显示导出成功弹窗
            if show_popup_var.get():
                messagebox.showinfo("导出完成", f"已导出文件到{excel_file_path}")
        else:
            messagebox.showerror("错误", f"文件格式不对，请检查！")
    else:
        root.mainloop()
