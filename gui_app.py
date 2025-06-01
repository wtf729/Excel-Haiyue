from tkinterdnd2 import DND_FILES, TkinterDnD
import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
from data_sort import data_sort_func
from config.constants import COLUMN_TYPE_MAPPING, COLUMN_TYPES_INPUT_SELECTIONS, COLUMN_TYPES_OUTPUT_SELECTIONS, OUTPUT_PRESETS, PRESET_CONFIG

def excel_col_to_index(col_str):
    col_str = col_str.upper()
    index = 0
    for char in col_str:
        if not char.isalpha():
            raise ValueError(f"列格式不正确：{col_str}")
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1

def apply_preset(preset_name):
    if preset_name not in PRESET_CONFIG:
        return

    configs = PRESET_CONFIG[preset_name]
    for i in range(3):
        if i < len(configs):
            config = configs[i]
            output_enabled_vars[i].set(config.get("enabled", True))
            output_name_style_vars[i].set(config.get("style", "中文"))
            output_sheet_name_vars[i].set(config.get("sheet_name", f"输出工作表{i+1}"))

            column_names = config.get("columns", [])
            for j in range(10):
                col_type = column_names[j] if j < len(column_names) else "无"
                cn_label = next((k for k, v in COLUMN_TYPE_MAPPING.items() if v == col_type), "无")
                output_column_type_vars_list[i][j].set(cn_label)
        else:
            output_enabled_vars[i].set(False)
            output_sheet_name_vars[i].set(f"输出工作表{i+1}")
            for j in range(10):
                output_column_type_vars_list[i][j].set("无")

# --- GUI 初始化 ---
root = TkinterDnD.Tk()
root.title("海悦动画 - 产量表助手")
root.geometry("1000x800")
root.iconbitmap("assets/app_icon.ico")

file_path_var = tk.StringVar()
file_path_var.set("请将 Excel 文件拖入此处")

tk.Label(root, text="拖入 Excel 文件：").pack(pady=(5, 0))

file_label = tk.Label(
    root,
    textvariable=file_path_var,
    bg="lightgrey",
    relief="groove",
    width=80,
    height=2,
    anchor="center",
    justify="center",
    font=("Microsoft YaHei", 10)
)
file_label.pack(pady=(0, 5), padx=10)

def on_file_drop(event):
    path = event.data.strip('{}')
    if path.lower().endswith(('.xls', '.xlsx')):
        file_path_var.set(path)
        status_text.insert(tk.END, f"已选择文件：{path}\n")
    else:
        status_text.insert(tk.END, "请拖入 Excel 文件（.xls 或 .xlsx）\n")

file_label.drop_target_register(DND_FILES)
file_label.dnd_bind("<<Drop>>", on_file_drop)

frame = tk.Frame(root)
frame.pack(pady=10)

tk.Label(frame, text="起始行:").grid(row=0, column=0)
start_row_entry = tk.Entry(frame, width=5)
start_row_entry.insert(0, "")
start_row_entry.grid(row=0, column=1)

tk.Label(frame, text="起始列:").grid(row=0, column=2)
start_col_entry = tk.Entry(frame, width=5)
start_col_entry.insert(0, "")
start_col_entry.grid(row=0, column=3)

tk.Label(frame, text="结束行:").grid(row=0, column=4)
end_row_entry = tk.Entry(frame, width=5)
end_row_entry.insert(0, "")
end_row_entry.grid(row=0, column=5)

tk.Label(frame, text="结束列:").grid(row=0, column=6)
end_col_entry = tk.Entry(frame, width=5)
end_col_entry.insert(0, "")
end_col_entry.grid(row=0, column=7)

price_frame = tk.Frame(root)
price_frame.pack(pady=(10, 0))

tk.Label(price_frame, text="动画单价:").grid(row=0, column=0)
price_ani_input = tk.Entry(price_frame, width=8)
price_ani_input.grid(row=0, column=1)

tk.Label(price_frame, text="上色单价:").grid(row=0, column=2)
price_coloring_input = tk.Entry(price_frame, width=8)
price_coloring_input.grid(row=0, column=3)

tk.Label(price_frame, text="一原单价:").grid(row=0, column=4)
price_1_yuan_input = tk.Entry(price_frame, width=8)
price_1_yuan_input.grid(row=0, column=5)

tk.Label(price_frame, text="二原单价:").grid(row=0, column=6)
price_2_yuan_input = tk.Entry(price_frame, width=8)
price_2_yuan_input.grid(row=0, column=7)


# 输入列类型设置
default_column_types = [
    "传票号", "无", "无", "无",
    "无", "无", "无", "无"
]

column_type_vars = []

tk.Label(root, text="输入列类型设置：").pack(pady=(10, 0))
type_frame = tk.Frame(root)
type_frame.pack()

for i in range(8):
    tk.Label(type_frame, text=f"第{i+1}列").grid(row=0, column=i*2, padx=3, pady=2, sticky="e")
    default_value = default_column_types[i] if i < len(default_column_types) else "无"
    var = tk.StringVar(value=default_value)
    combo = ttk.Combobox(type_frame, values=COLUMN_TYPES_INPUT_SELECTIONS, textvariable=var, width=8, state="readonly")
    combo.grid(row=1, column=i*2, padx=3, pady=2)
    column_type_vars.append(var)

# 输出设置区域
tk.Label(root, text="输出参数预设：").pack(pady=(10, 0))
preset_var = tk.StringVar(value="无")
preset_dropdown = ttk.Combobox(root, values=OUTPUT_PRESETS, textvariable=preset_var, width=10, state="readonly")
preset_dropdown.pack()
preset_dropdown.bind("<<ComboboxSelected>>", lambda e: apply_preset(preset_var.get()))

def create_output_sheet_section(sheet_num):
    section_frame = tk.LabelFrame(root, text=f"输出工作表{sheet_num}", padx=5, pady=5)
    section_frame.pack(padx=10, pady=5)

    # 让 section_frame 的列支持居中对齐
    section_frame.grid_columnconfigure(0, weight=1)

    # 包裹启用/下拉/命名输入的小框架（整行水平居中）
    top_row_frame = tk.Frame(section_frame)
    top_row_frame.grid(row=0, column=0, pady=3)

    # 启用勾选框
    enabled_var = tk.BooleanVar(value=False)
    check = tk.Checkbutton(top_row_frame, text=f"启用输出工作表{sheet_num}", variable=enabled_var)
    check.pack(side="left", padx=10)

    # 输出列名样式组合
    name_style_frame = tk.Frame(top_row_frame)
    name_style_frame.pack(side="left", padx=10)

    tk.Label(name_style_frame, text="输出列名样式:").pack(side="left")
    name_style_var = tk.StringVar(value="中文")
    name_style_dropdown = ttk.Combobox(name_style_frame, values=["中文", "日文"], textvariable=name_style_var, width=6,
                                       state="readonly")
    name_style_dropdown.pack(side="left")

    # 工作表名称组合
    sheet_name_frame = tk.Frame(top_row_frame)
    sheet_name_frame.pack(side="left", padx=10)

    tk.Label(sheet_name_frame, text="工作表名称:").pack(side="left")
    sheet_name_var = tk.StringVar(value=f"输出工作表{sheet_num}")
    sheet_name_entry = tk.Entry(sheet_name_frame, textvariable=sheet_name_var, width=14)
    sheet_name_entry.pack(side="left")

    # 输出列设置区域
    row_frame = tk.Frame(section_frame)
    row_frame.grid(row=1, column=0, pady=5)



    vars = []
    combos = []

    for i in range(10):
        item_frame = tk.Frame(row_frame)
        item_frame.pack(side="left", padx=5)

        tk.Label(item_frame, text=f"第{i + 1}列").pack()
        var = tk.StringVar(value="无")
        combo = ttk.Combobox(item_frame, values=COLUMN_TYPES_OUTPUT_SELECTIONS, textvariable=var, width=7, state="readonly")
        combo.pack()

        vars.append(var)
        combos.append(combo)

    def toggle_state(*_):
        state = "readonly" if enabled_var.get() else "disabled"
        for combo in combos:
            combo.config(state=state)
        name_style_dropdown.config(state=state)
        sheet_name_entry.config(state="normal" if enabled_var.get() else "disabled")

    enabled_var.trace_add("write", toggle_state)
    toggle_state()

    return enabled_var, vars, name_style_var, sheet_name_var



output_enabled_vars = []
output_column_type_vars_list = []
output_name_style_vars = []
output_sheet_name_vars = []

for i in range(1, 4):
    enabled_var, col_vars, name_style_var, sheet_name_var = create_output_sheet_section(i)
    output_enabled_vars.append(enabled_var)
    output_column_type_vars_list.append(col_vars)
    output_name_style_vars.append(name_style_var)
    output_sheet_name_vars.append(sheet_name_var)

# 执行处理函数
def process_excel():
    file_path = file_path_var.get()
    if not file_path:
        status_text.insert(tk.END, "请先拖入一个 Excel 文件。\n")
        return

    try:
        prices = {}
        for label, var in [
            ("动画单价", price_ani_input),
            ("上色单价", price_coloring_input),
            ("一原单价", price_1_yuan_input),
            ("二原单价", price_2_yuan_input),
        ]:
            value = var.get().strip()
            if value:
                try:
                    num = float(value)
                    if num <= 0:
                        raise ValueError
                    prices[label] = num
                except ValueError:
                    status_text.insert(tk.END, f"{label} 请输入正数！\n")
                    return

        df = pd.read_excel(file_path)
        start_row = int(start_row_entry.get()) - 2
        end_row = int(end_row_entry.get()) - 1
        start_col = excel_col_to_index(start_col_entry.get())
        end_col = excel_col_to_index(end_col_entry.get()) + 1

        selected_df = df.iloc[start_row:end_row, start_col:end_col]

        selected_column_types = [var.get() for var in column_type_vars]
        selected_non_none = [x for x in selected_column_types if x != "无"]
        if len(selected_non_none) != len(set(selected_non_none)):
            status_text.insert(tk.END, "输入列分类存在重复，请检查。\n")
            return

        # 输出列分类检查（仅检查已启用的工作表）
        for i in range(3):
            if output_enabled_vars[i].get():
                output_labels = [var.get() for var in output_column_type_vars_list[i]]
                output_non_none = [x for x in output_labels if x != "无"]
                if len(output_non_none) != len(set(output_non_none)):
                    status_text.insert(tk.END, f"输出工作表{i + 1}列分类存在重复，请检查。\n")
                    return

        # 过滤掉列类型为“无”的列
        valid_columns = [
            (idx, COLUMN_TYPE_MAPPING.get(label))
            for idx, label in enumerate(selected_column_types)
            if label != "无"
        ]
        # 从 selected_df 中选择有效列
        selected_df = selected_df.iloc[:, [idx for idx, _ in valid_columns]]
        # 提取内部字段名
        internal_column_names = [name for _, name in valid_columns]

        output_sheets_config = []
        for i in range(3):
            if output_enabled_vars[i].get():
                labels = output_column_type_vars_list[i]
                internal_names = [COLUMN_TYPE_MAPPING.get(var.get(), None) for var in labels]
                output_sheets_config.append({
                    "enabled": True,
                    "columns": internal_names,
                    "style": output_name_style_vars[i].get(),
                    "sheet_name": output_sheet_name_vars[i].get()
                })

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not save_path:
            status_text.insert(tk.END, "用户取消了保存操作。\n")
            return

        try:
            # 调用 data_sort_func 并传入所有配置
            data_sort_func(
                selected_df,
                internal_column_names,
                prices,
                output_sheets_config,
                save_path
            )
            status_text.insert(tk.END, f"处理完成，保存至：{save_path}\n")

        except Exception as e:
            status_text.insert(tk.END, f"发生错误：{e}\n")

    except Exception as e:
        status_text.insert(tk.END, f"发生错误：{e}\n")

# 执行按钮
button = tk.Button(root, text="执行", command=process_excel)
button.pack(pady=10)

status_text = tk.Text(root, height=5)
status_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

root.mainloop()
