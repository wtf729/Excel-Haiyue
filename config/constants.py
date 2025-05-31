COLUMN_TYPE_MAPPING = {
    "传票号": "order_number",
    "株式会社": "company_name",
    "片名": "animation_name",
    "话数": "animation_episode",
    "动画数量": "count_ani",
    "上色数量": "count_coloring",
    "一原数量": "count_1_yuan",
    "二原数量": "count_2_yuan",
    "动画单价": "price_ani",
    "上色单价": "price_coloring",
    "一原单价": "price_1_yuan",
    "二原单价": "price_2_yuan",
    "动画总价": "total_ani",
    "上色总价": "total_coloring",
    "一原总价": "total_1_yuan",
    "二原总价": "total_2_yuan",
    "无": None,
    "---------": None
}

COLUMN_TYPE_OUTPUT_CN = {
    "order_number": "传票号",
    "company_name": "株式会社",
    "animation_name": "片名",
    "animation_episode": "话数",
    "count_ani": "动画",
    "count_coloring": "上色",
    "count_1_yuan": "一原",
    "count_2_yuan": "二原",
    "price_ani": "动画单价",
    "price_coloring": "上色单价",
    "price_1_yuan": "一原单价",
    "price_2_yuan": "二原单价",
    "total_ani": "动画总价",
    "total_coloring": "上色总价",
    "total_1_yuan": "一原总价",
    "total_2_yuan": "二原总价"
}

COLUMN_TYPE_OUTPUT_JP = {
    "order_number": "NO",
    "company_name": "株式会社",
    "animation_name": "タ　イ　ト　ル(デジタル)",
    "animation_episode": "话数",
    "count_ani": "動画",
    "count_coloring": "仕上げ",
    "count_1_yuan": "L/Oー作监",
    "count_2_yuan": "2原",
    "price_ani": "単価",
    "price_coloring": "単価",
    "price_1_yuan": "単価",
    "price_2_yuan": "単価",
    "total_ani": "合　計",
    "total_coloring": "合　計",
    "total_1_yuan": "合　計",
    "total_2_yuan": "合　計"
}

COLUMN_TYPES_INPUT_SELECTIONS = [
    "传票号", "株式会社", "片名", "话数",
    "动画数量", "上色数量", "一原数量", "二原数量",
    "无"
]

COLUMN_TYPES_OUTPUT_SELECTIONS = [
    "传票号", "株式会社", "片名", "话数",
    "动画数量", "上色数量", "一原数量", "二原数量", "---------",
    "动画单价", "上色单价", "一原单价", "二原单价", "---------",
    "动画总价", "上色总价", "一原总价", "二原总价",
    "无"
]