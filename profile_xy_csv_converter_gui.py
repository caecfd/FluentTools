import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import scrolledtext
import re
import os
import csv


# 文件格式转换类定义
class FileConverter:
    """
    文件格式转换类，支持CSV、PROF和XY文件之间的相互转换
    """
    
    def convert_csv_to_prof(self, csv_path):
        """
        将CSV文件转换为PROF文件
        
        Args:
            csv_path: CSV文件路径
            
        Returns:
            str: 生成的PROF文件路径
        """
        with open(csv_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            headers = next(reader)
            data = list(reader)
        rows = len(data)
        # 提取不带路径的文件名（不含扩展名）
        name = os.path.splitext(os.path.basename(csv_path))[0]
        
        prof_content = f'(({name} point {rows} 1)\n'
        for col_idx, col_name in enumerate(headers):
            col_data = [row[col_idx] for row in data]
            prof_content += f'({col_name} '
            for i, val in enumerate(col_data):
                prof_content += f'{val} '
                if (i + 1) % 6 == 0 or i == len(col_data) - 1:
                    prof_content += '\n'
            prof_content += ')\n'
        prof_content += ')'
        
        # 使用完整路径生成PROF文件
        prof_path = os.path.splitext(csv_path)[0] + '.prof'
        with open(prof_path, 'w', encoding='utf-8') as f:
            f.write(prof_content)
        print(f'生成 {prof_path}，{rows} 行，{len(headers)} 列')
        return prof_path
    
    def convert_csv_to_xy(self, csv_path):
        """
        将CSV文件转换为XY文件
        
        Args:
            csv_path: CSV文件路径
            
        Returns:
            str: 生成的XY文件路径
        """
        with open(csv_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            headers = next(reader)
            data = list(reader)

        if len(headers) != 2 or len(data[0]) != 2:
            raise ValueError('CSV必须有精确2列')

        xy_path = csv_path.rsplit('.', 1)[0] + '.xy'
        with open(xy_path, 'w', encoding='utf-8') as f:
            f.write(f'(title "{headers[1]}")\n')
            f.write(f'(labels "{headers[0]}" "{headers[1]}")\n')
            f.write('\n')
            f.write('((xy/key/label "location")\n')
            for row in data:
                f.write(f'{row[0]}\t{row[1]}\n')
            f.write(')\n')

        rows = len(data)
        print(f'生成 {xy_path}，{rows} 行，2 列')
        return xy_path
    
    def convert_prof_to_csv(self, file_path):
        """
        将PROF文件转换为CSV文件
        
        Args:
            file_path: PROF文件路径
            
        Returns:
            str: 生成的CSV文件路径
        """
        with open(file_path, 'r') as f:
            lines = [line.strip() for line in f.readlines() if line.strip()]
        if not lines:
            raise ValueError('文件为空')
        # 首行配置 ((name kw rows ignore)
        first = lines[0]
        if not first.startswith('((') or not first.endswith(')'):
            raise ValueError('配置行格式无效')
        config_str = first[2:-1].strip()  # 去内外层括号
        config = re.split(r'\s+', config_str)
        name = config[0]
        rows = int(config[2])
        columns = []
        i = 1
        while i < len(lines):
            line = lines[i]
            if line.startswith('(') and not line.endswith(')'):
                parts = line[1:].split()
                if parts:
                    col_name = parts[0]
                else:
                    col_name = 'unnamed'
                col_data = []
                # 提取列名行剩余数据
                nums_header = re.findall(r'[-+]??\d*\.?\d+(?:[eE][-+]?\d+)?', ' '.join(parts[1:]))
                col_data.extend(nums_header)
                i += 1
                while i < len(lines) and not lines[i].endswith(')'):
                    nums = re.findall(r'[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?', lines[i])
                    col_data.extend(nums)
                    i += 1
                # 处理结束行
                if i < len(lines) and lines[i].endswith(')'):
                    nums = re.findall(r'[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?', lines[i])
                    col_data.extend(nums)
                    columns.append((col_name, col_data))
                i += 1
            else:
                i += 1
        headers = [c[0] for c in columns]
        col_data_list = [c[1] for c in columns]
        for j, cd in enumerate(col_data_list):
            if len(cd) != rows:
                print(f'警告：列 {headers[j]} 有 {len(cd)} 行，预期 {rows} 行')
        # 使用输入文件的路径生成同名但扩展名为csv的输出文件路径
        csv_path = os.path.splitext(file_path)[0] + '.csv'
        with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(headers)
            for r in range(rows):
                row = [col_data_list[k][r] if r < len(col_data_list[k]) else '' for k in range(len(columns))]
                writer.writerow(row)
        print(f'生成 {csv_path}，{rows} 行，{len(headers)} 列')
        return csv_path
    
    def convert_xy_to_csv(self, xy_path):
        """
        将XY文件转换为CSV文件
        
        Args:
            xy_path: XY文件路径
            
        Returns:
            str: 生成的CSV文件路径
        """
        with open(xy_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        headers = None
        data = []
        found_labels = False

        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if line.startswith('(labels'):
                match = re.match(r'\(labels\s+"([^"]+)"\s+"([^"]+)"\)', line)
                if match:
                    headers = [match.group(1), match.group(2)]
                    found_labels = True
                i += 1
                continue
            if found_labels and line == ')':
                break
            if found_labels and line and not line.startswith('(') and not line.startswith(')'):
                parts = re.findall(r'[-+]?[0-9]*\.?[0-9]+(?:[eE][-+]?[0-9]+)?', line)
                if len(parts) >= 2:
                    data.append([parts[0], parts[1]])
            i += 1

        if not headers:
            raise ValueError('未找到 labels 行')

        csv_path = xy_path.rsplit('.', 1)[0] + '.csv'
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(data)

        rows = len(data)
        print(f'生成 {csv_path}，{rows} 行，2 列')
        return csv_path

    def convert(self, input_file, conversion_type):
        """
        根据指定的转换类型执行相应的转换操作
        
        Args:
            input_file: 输入文件路径
            conversion_type: 转换类型（CSV2PROF、PROF2CSV、CSV2XY、XY2CSV）
            
        Returns:
            str: 生成的输出文件路径
            
        Raises:
            ValueError: 当转换类型无效时
        """
        if conversion_type == "CSV2PROF":
            return self.convert_csv_to_prof(input_file)
        elif conversion_type == "PROF2CSV":
            return self.convert_prof_to_csv(input_file)
        elif conversion_type == "CSV2XY":
            return self.convert_csv_to_xy(input_file)
        elif conversion_type == "XY2CSV":
            return self.convert_xy_to_csv(input_file)
        else:
            raise ValueError(f'无效的转换类型: {conversion_type}')

# 初始化FileConverter类实例
file_converter = FileConverter()

def browse_file():
    # 根据转换类型设置对应的文件过滤器
    conv_type = selected_conversion.get()
    file_filters = [("所有文件", "*.*")]
    
    if conv_type == "CSV2PROF":
        file_filters = [("CSV文件", "*.csv"), ("所有文件", "*.*")]
    elif conv_type == "PROF2CSV":
        file_filters = [("PROF文件", "*.prof"), ("所有文件", "*.*")]
    elif conv_type == "CSV2XY":
        file_filters = [("CSV文件", "*.csv"), ("所有文件", "*.*")]
    elif conv_type == "XY2CSV":
        file_filters = [("XY文件", "*.xy"), ("所有文件", "*.*")]
    
    filename = filedialog.askopenfilename(
        title="选择输入文件",
        filetypes=file_filters
    )
    if filename:
        entry.delete(0, tk.END)
        entry.insert(0, filename)

def perform_conversion():
    input_file = entry.get().strip()
    conv_type = selected_conversion.get()
    if not input_file or not conv_type:
        error_msg = "请选择转换方式并选择输入文件！"
        messagebox.showerror("错误", error_msg)
        update_output_log(error_msg)
        return
    
    # 根据转换类型确定目标扩展名
    extension_map = {
        "CSV2PROF": ".prof",
        "PROF2CSV": ".csv",
        "CSV2XY": ".xy",
        "XY2CSV": ".csv"
    }
    
    # 生成输出文件路径（保持相同目录和文件名，只修改扩展名）
    base_name = os.path.splitext(input_file)[0]
    output_file = base_name + extension_map[conv_type]
    
    # workspace = r"d:\STAR CCM+\VM14\test"  # 不再需要
    
    # 更新日志显示开始转换
    update_output_log(f"开始转换：{conv_type}\n输入文件：{input_file}\n输出文件：{output_file}\n")
    
    try:
        # 使用FileConverter类进行转换，不再调用子进程
        output_file = file_converter.convert(input_file, conv_type)
        
        # 更新输出文件路径显示
        output_entry.config(state=tk.NORMAL)
        output_entry.delete(0, tk.END)
        output_entry.insert(0, output_file)
        # 保持NORMAL状态，不再设置为DISABLED
        # output_entry.config(state=tk.DISABLED)
        
        # 更新日志
        log_text = f"转换完成！\n输入文件：{input_file}\n输出文件：{output_file}\n"
        update_output_log(log_text)
        
        messagebox.showinfo("成功", f"转换完成！\n生成文件：{output_file}")
        
    except Exception as e:
        error_msg = f"转换出错！\n错误：{str(e)}"
        messagebox.showerror("失败", error_msg)
        update_output_log(error_msg)

def open_output_file():
    """打开输出文件"""
    output_path = output_entry.get().strip()
    if output_path and os.path.exists(output_path):
        try:
            os.startfile(output_path)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件：{str(e)}")
    else:
        messagebox.showinfo("提示", "输出文件不存在")

def update_output_log(text):
    """更新输出日志文本框"""
    output_log.config(state=tk.NORMAL)
    output_log.delete(1.0, tk.END)
    output_log.insert(tk.END, text)
    output_log.config(state=tk.DISABLED)

root = tk.Tk()
root.title("文件转换器 (CSV ↔ PROF/XY)")
root.geometry("720x500")  # 增大窗口尺寸以适应新控件

# 设置ttk主题
style = ttk.Style(root)
# 尝试使用更现代的主题，如果可用
available_themes = style.theme_names()
if 'clam' in available_themes:
    style.theme_use('clam')
elif 'vista' in available_themes:
    style.theme_use('vista')
elif 'alt' in available_themes:
    style.theme_use('alt')

# 自定义控件样式
# 按钮样式
style.configure('TButton',
    padding=6,
    relief="flat",
    font=('Microsoft YaHei UI', 9))

# 转换按钮特殊样式
style.configure('Convert.TButton',
    foreground='#ffffff',
    background='#4a7abc',
    padding=8,
    font=('Microsoft YaHei UI', 9, 'bold'))

# 浏览和打开按钮样式（稍淡的蓝色）
style.configure('Browse.TButton',
    foreground='#ffffff',
    background='#7daae8',
    padding=6,
    font=('Microsoft YaHei UI', 9))

# 输入框样式
style.configure('TEntry',
    padding=5,
    font=('Microsoft YaHei UI', 9))

# 下拉框样式
style.configure('TCombobox',
    padding=5,
    font=('Microsoft YaHei UI', 9))

# 标签样式
style.configure('TLabel',
    font=('Microsoft YaHei UI', 9))

# ScrolledText样式
style.configure('TScrolledText',
    font=('Microsoft YaHei UI', 9))

# 创建主框架
main_frame = ttk.Frame(root, padding=10)
main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10)

# 添加标题区域
title_frame = ttk.Frame(main_frame)
title_frame.grid(row=0, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0,15))

# 程序标题
title_label = ttk.Label(title_frame, text="文件转换器", font=('Microsoft YaHei UI', 12, 'bold'))
title_label.pack(side=tk.LEFT, padx=10)

# 标题下的分隔线
title_separator = ttk.Separator(main_frame, orient='horizontal')
title_separator.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0,15))

# 配置权重，确保界面能够响应式调整
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

# 优化列配置，设置更加合理的宽度和权重
main_frame.columnconfigure(0, weight=0, minsize=90)  # 标签列稍宽一些
main_frame.columnconfigure(1, weight=1)  # 内容列自适应宽度
main_frame.columnconfigure(2, weight=0)  # 按钮列不扩展
main_frame.columnconfigure(3, weight=0)

# 转换方式 - 使用单选按钮组
# 创建一个变量来存储选中的单选按钮
selected_conversion = tk.StringVar(value="CSV2PROF")
ttk.Label(main_frame, text="转换方式：").grid(row=2, column=0, sticky=tk.NW, pady=(0,10))

# 创建单选按钮组框架
radio_frame = ttk.Frame(main_frame)
radio_frame.grid(row=2, column=1, sticky=(tk.W), padx=(10, 5), pady=(0,10))

# 为radio_frame配置列权重，确保四个单选按钮能很好地排列
for i in range(4):
    radio_frame.columnconfigure(i, weight=0, minsize=100)

# 创建单选按钮
conversion_types = ["CSV2PROF", "PROF2CSV", "CSV2XY", "XY2CSV"]
for i, conv_type in enumerate(conversion_types):
    # 将单选按钮按一行四列横向排列
    row = 0
    col = i
    ttk.Radiobutton(radio_frame, text=conv_type, variable=selected_conversion, value=conv_type).grid(
        row=row, column=col, sticky=tk.W, padx=10, pady=5)

# 输入文件
ttk.Label(main_frame, text="输入文件：").grid(row=3, column=0, sticky=tk.W, pady=(10,10))
entry = ttk.Entry(main_frame, width=50)
entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=(10,10), padx=(10,5))
entry.bind('<Button-1>', lambda e: browse_file())  # 绑定点击事件

# 移除原来的浏览按钮
# browse_btn = ttk.Button(main_frame, text="浏览", command=browse_file, style='Browse.TButton')
# browse_btn.grid(row=3, column=2, sticky=tk.W, padx=(5,10), pady=(10,10))

# 转换按钮（应用自定义样式）
convert_btn = ttk.Button(main_frame, text="转换", command=perform_conversion, style='Convert.TButton')
convert_btn.grid(row=4, column=1, sticky=tk.W, padx=(10, 0), pady=(0,15))

# 输出文件
ttk.Label(main_frame, text="输出文件：").grid(row=5, column=0, sticky=tk.W, pady=(10,10))
output_entry = ttk.Entry(main_frame, width=50)
output_entry.grid(row=5, column=1, sticky=(tk.W, tk.E), pady=(10,10), padx=(10,5))
output_entry.config(state=tk.NORMAL)  # 允许点击交互
output_entry.bind('<Button-1>', lambda e: open_output_file())  # 绑定点击事件

# 移除原来的打开按钮
# open_btn = ttk.Button(main_frame, text="打开", command=open_output_file, style='Browse.TButton')
# open_btn.grid(row=5, column=2, sticky=tk.W, padx=(5,10), pady=(10,10))

# 输出日志文本框
ttk.Label(main_frame, text="执行信息：").grid(row=6, column=0, sticky=tk.NW, pady=(10,5))
output_log = scrolledtext.ScrolledText(main_frame, height=8, wrap=tk.WORD, state=tk.DISABLED, font=('Microsoft YaHei UI', 9))
output_log.grid(row=6, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10,10), pady=(10,5))

# 确保日志文本框可以垂直扩展
main_frame.rowconfigure(6, weight=1)

root.mainloop()