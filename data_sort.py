import pandas as pd

def data_sort_func(df, column_type_list):
    # 获取列对应的标记类型
    new_columns = []
    for idx, col_type in enumerate(column_type_list):
        if idx < df.shape[1]:
            new_columns.append(col_type if col_type else f"unknown_{idx}")

    df = df.iloc[:, :len(new_columns)]
    df.columns = new_columns

    if "animation_name" in df.columns:
        df = df.sort_values(by="animation_name", ascending=True)


    # 你可以在此添加更多的清洗规则，例如合并列、生成新列、统计总量等

    return df