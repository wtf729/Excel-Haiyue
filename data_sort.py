import pandas as pd
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 映射表
COLUMN_TYPE_MAPPING = {
    "传票号": "order_number",
    "株式会社": "company_name",
    "片名": "animation_name",
    "话数": "animation_episode",
    "动画数量": "count_ani",
    "上色数量": "count_coloring",
    "一原数量": "count_1_yuan",
    "二原数量": "count_2_yuan"
}

def data_sort_func(selected_df, internal_column_names, prices, output_sheets_config, save_path):
    # 1. 列重命名
    df = selected_df.copy()
    df.columns = internal_column_names

    # 2. 删除全为空的行
    df.dropna(how='all', inplace=True)

    # 3. 验证 count_* 列
    count_cols = [col for col in ["count_ani", "count_coloring", "count_1_yuan", "count_2_yuan"] if col in df.columns]
    error_rows = []

    for col in count_cols:
        temp_col = pd.to_numeric(df[col], errors='coerce')
        invalid_mask = ~((temp_col.notna()) & (temp_col % 1 == 0) & (temp_col >= 0))
        if invalid_mask.any():
            for idx in df[invalid_mask].index:
                error_rows.append(f"第 {idx + 2} 行（列名：{col}）的值不为非负整数：{df.at[idx, col]}")

    if error_rows:
        raise ValueError("检测到以下单元格不是合法的非负整数：\n" + "\n".join(error_rows))

    for col in count_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).infer_objects(copy=False).astype(int)

    # 4. 排序
    sort_keys = [k for k in ["animation_name", "animation_episode", "order_number"] if k in df.columns]
    df.sort_values(by=sort_keys, ascending=True, inplace=True)

    # 5. 写入 Excel（检查点工作表）
    reverse_column_map = {v: k for k, v in COLUMN_TYPE_MAPPING.items()}
    df_for_excel = df.rename(columns=reverse_column_map)

    with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
        df_for_excel.to_excel(writer, sheet_name='data_sorted', index=False)
        worksheet = writer.sheets['data_sorted']

        # 设置格式
        align_center = Alignment(horizontal='center', vertical='center')
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )

        for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row,
                                       min_col=1, max_col=worksheet.max_column):
            for cell in row:
                col_letter = get_column_letter(cell.column)
                col_name = worksheet[f"{col_letter}1"].value

                # 设置列宽
                if worksheet[f"{col_letter}1"].row == 1:
                    if col_name in ["株式会社", "片名"]:
                        worksheet.column_dimensions[col_letter].width = 40

                # 居中对齐特定列
                if col_name in ["话数", "动画数量", "上色数量", "一原数量", "二原数量"]:
                    cell.alignment = align_center

                # 所有格子添加边框
                cell.border = thin_border

    # 注意：df 仍保持英文列名，用于后续输出其他工作表
    return df
