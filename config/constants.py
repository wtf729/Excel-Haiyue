# constants.py

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
    "无": None
}

COLUMN_TYPES = [
    "传票号", "株式会社", "片名", "话数",
    "动画数量", "上色数量", "一原数量", "二原数量",
    "动画单价", "上色单价", "一原单价", "二原单价",
    "无"
]