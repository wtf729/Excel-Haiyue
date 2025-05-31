import pandas as pd
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter
from config.constants import COLUMN_TYPE_MAPPING, COLUMN_TYPE_OUTPUT_CN, COLUMN_TYPE_OUTPUT_JP

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
        temp_col = pd.to_numeric(df[col], errors='coerce')  # 转换为数值（NaN 代表非法值）
        invalid_mask = ~((temp_col.notna()) & (temp_col >= 0))  # 非负浮点数的条件

        if invalid_mask.any():
            for idx in df[invalid_mask].index:
                error_rows.append(f"第 {idx + 2} 行（列名：{col}）的值不是非负浮点数：{df.at[idx, col]}")

    if error_rows:
        raise ValueError("检测到以下单元格不是合法的非负浮点数：\n" + "\n".join(error_rows))

    # 校验通过后，统一转为 float 类型，非法值用 0 填充
    for col in count_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0).astype(float)

    # 4. 排序
    sort_keys = [k for k in ["animation_name", "animation_episode", "order_number"] if k in df.columns]
    df.sort_values(by=sort_keys, ascending=True, inplace=True)

    # 5. 写入 Excel（检查点工作表）
    df_sorted_for_excel = df.rename(columns=COLUMN_TYPE_OUTPUT_CN)

    # 1. 删除 order_number 列
    df_for_calc = df.drop(columns=["order_number"]) if "order_number" in df.columns else df.copy()

    # 2. 动态选取存在的 count_* 列
    count_cols = [col for col in ["count_ani", "count_coloring", "count_1_yuan", "count_2_yuan"] if
                  col in df_for_calc.columns]

    # 3. 分组汇总
    # 预设想要用来分组的列
    desired_group_keys = ["company_name", "animation_name", "animation_episode"]
    # 实际存在于 DataFrame 中的分组列
    actual_group_keys = [col for col in desired_group_keys if col in df_for_calc.columns]
    # 如果没有分组列，直接复制原始数据；否则进行分组汇总
    if actual_group_keys:
        grouped_df = df_for_calc.groupby(actual_group_keys, as_index=False)[count_cols].sum()
    else:
        grouped_df = df_for_calc.copy()

    # 4. 转换回中文列名
    df_calculated_for_excel = grouped_df.rename(columns=COLUMN_TYPE_OUTPUT_CN)


    with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
        align_center = Alignment(horizontal='center', vertical='center')
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        # 写入 calculated 表
        df_calculated_for_excel.to_excel(writer, sheet_name='data_calculated', index=False)
        worksheet_calculated = writer.sheets['data_calculated']
        # 设置格式
        for row in worksheet_calculated.iter_rows(min_row=1, max_row=worksheet_calculated.max_row,
                                       min_col=1, max_col=worksheet_calculated.max_column):
            for cell in row:
                col_letter = get_column_letter(cell.column)
                col_name = worksheet_calculated[f"{col_letter}1"].value
                # 设置列宽
                if worksheet_calculated[f"{col_letter}1"].row == 1:
                    if col_name in ["株式会社", "片名"]:
                        worksheet_calculated.column_dimensions[col_letter].width = 50
                # 居中对齐特定列
                if col_name in ["话数", "动画数量", "上色数量", "一原数量", "二原数量"]:
                    cell.alignment = align_center
                # 所有格子添加边框
                cell.border = thin_border
        # 写入 sorted 表
        df_sorted_for_excel.to_excel(writer, sheet_name='data_sorted', index=False)
        worksheet_sorted = writer.sheets['data_sorted']
        # 设置格式
        for row in worksheet_sorted.iter_rows(min_row=1, max_row=worksheet_sorted.max_row,
                                       min_col=1, max_col=worksheet_sorted.max_column):
            for cell in row:
                col_letter = get_column_letter(cell.column)
                col_name = worksheet_sorted[f"{col_letter}1"].value
                # 设置列宽
                if worksheet_sorted[f"{col_letter}1"].row == 1:
                    if col_name in ["株式会社", "片名"]:
                        worksheet_sorted.column_dimensions[col_letter].width = 50
                # 居中对齐特定列
                if col_name in ["话数", "动画数量", "上色数量", "一原数量", "二原数量"]:
                    cell.alignment = align_center
                # 所有格子添加边框
                cell.border = thin_border

    return


# TODO
# count_cols校验改为浮点数，输出时将小数点后没有数字的数显示为整数形式
# 调整使用COLUMN_TYPE_OUTPUT_CN进行输出
# 排除公司列不存在的情况