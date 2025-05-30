from tkinterdnd2 import DND_FILES, TkinterDnD
import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
from data_sort import data_sort_func

COLUMN_TYPES = [
    "传票号", "株式会社", "片名", "话数",
    "动画数量", "上色数量", "一原数量", "二原数量",
    "无"
]

COLUMN_TYPE_MAPPING = {
    "传票号": "order_number",
    "株式会社": "company_name",
    "片名": "animation_name",
    "话数": "animation_episode",
    "动画数量": "count_ani",
    "上色数量": "count_coloring",
    "一原数量": "count_1_yuan",
    "二原数量": "count_2_yuan",
    "无": None
}

def excel_col_to_index(col_str):
    col_str = col_str.upper()
    index = 0
    for char in col_str:
        if not char.isalpha():
            raise ValueError(f"列格式不正确：{col_str}")
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1

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
        start_row = int(start_row_entry.get()) - 1
        end_row = int(end_row_entry.get())
        start_col = excel_col_to_index(start_col_entry.get())
        end_col = excel_col_to_index(end_col_entry.get()) + 1

        selected_df = df.iloc[start_row:end_row, start_col:end_col]

        selected_column_types = [var.get() for var in column_type_vars]

        # 检查重复
        selected_non_none = [x for x in selected_column_types if x != "无"]
        if len(selected_non_none) != len(set(selected_non_none)):
            status_text.insert(tk.END, "列分类存在重复，请检查。\n")
            return

        internal_column_names = [COLUMN_TYPE_MAPPING.get(label, None) for label in selected_column_types]

        result_df = data_sort_func(selected_df, internal_column_names)

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            result_df.to_excel(save_path, index=False)
            status_text.insert(tk.END, f"处理完成，保存至：{save_path}\n")
        else:
            status_text.insert(tk.END, "用户取消了保存操作。\n")

    except Exception as e:
        status_text.insert(tk.END, f"发生错误：{e}\n")

# --- GUI 初始化 ---
root = TkinterDnD.Tk()
root.title("海悦动画 - 产量表助手")
root.geometry("1000x450")



file_path_var = tk.StringVar()
file_path_var.set("请将 Excel 文件拖入此处")  # 初始提示文字

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
file_label.pack(pady=(0, 5), padx=10, anchor="center")  # 不再 fill="x"

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

# 输入范围设置
tk.Label(frame, text="起始行:").grid(row=0, column=0)
start_row_entry = tk.Entry(frame, width=5)
start_row_entry.insert(0, "3")
start_row_entry.grid(row=0, column=1)

tk.Label(frame, text="结束行:").grid(row=0, column=2)
end_row_entry = tk.Entry(frame, width=5)
end_row_entry.insert(0, "100")
end_row_entry.grid(row=0, column=3)

tk.Label(frame, text="起始列:").grid(row=0, column=4)
start_col_entry = tk.Entry(frame, width=5)
start_col_entry.insert(0, "A")
start_col_entry.grid(row=0, column=5)

tk.Label(frame, text="结束列:").grid(row=0, column=6)
end_col_entry = tk.Entry(frame, width=5)
end_col_entry.insert(0, "G")
end_col_entry.grid(row=0, column=7)

# 单价输入
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

# 下拉框设置
default_column_types = [
    "传票号", "株式会社", "片名", "话数",
    "动画数量", "上色数量", "二原数量", "无",
    "无", "无"
]

column_type_vars = []

tk.Label(root, text="列类型设置：").pack(pady=(10, 0))
type_frame = tk.Frame(root)
type_frame.pack()

for i in range(10):
    tk.Label(type_frame, text=f"第{i+1}列").grid(row=0, column=i*2, padx=3, pady=2, sticky="e")
    default_value = default_column_types[i] if i < len(default_column_types) else "无"
    var = tk.StringVar(value=default_value)
    combo = ttk.Combobox(type_frame, values=COLUMN_TYPES, textvariable=var, width=8, state="readonly")
    combo.grid(row=1, column=i*2, padx=3, pady=2, sticky="w")
    column_type_vars.append(var)

# 执行按钮
button = tk.Button(root, text="执行", command=process_excel)
button.pack(pady=10)

# 状态输出框
status_text = tk.Text(root, height=5)
status_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

root.mainloop()
