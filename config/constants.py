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
    "order_number": "NO.",
    "company_name": "株式会社",
    "animation_name": "タィトル",
    "animation_episode": "话数",
    "count_ani": "動画",
    "count_coloring": "仕上げ",
    "count_1_yuan": "L/Oー作监",
    "count_2_yuan": "2原",
    "price_ani": "単価",
    "price_coloring": "単価",
    "price_1_yuan": "単価",
    "price_2_yuan": "単価",
    "total_ani": "合計",
    "total_coloring": "合計",
    "total_1_yuan": "合計",
    "total_2_yuan": "合計"
}

COLUMN_TYPES_INPUT_SELECTIONS = [
    "无",
    "传票号", "株式会社", "片名", "话数",
    "动画数量", "上色数量", "一原数量", "二原数量"
]

COLUMN_TYPES_OUTPUT_SELECTIONS = [
    "无",
    "株式会社", "片名", "话数",
    "动画数量", "上色数量", "一原数量", "二原数量", "---------",
    "动画单价", "上色单价", "一原单价", "二原单价", "---------",
    "动画总价", "上色总价", "一原总价", "二原总价"
]


INPUT_PRESETS = ["无", "全部"]

INPUT_PRESET_CONFIG = {
    "无": {
        "column_types": ["传票号"]
    },
    "全部": {
        "column_types": ["传票号", "株式会社", "片名", "话数", "动画数量", "上色数量", "一原数量", "二原数量"]
    }
}



OUTPUT_PRESETS = [
    "无",
    "Asahi", "MSJ"
]

OUTPUT_PRESET_CONFIG = {
    "无": [
        {
            "enabled": False,
            "style": "中文",
            "columns": []
        }
    ],
    "Asahi": [
        {
            "enabled": True,
            "style": "日文",
            "sheet_name": "旭-动仕",
            "columns": [
                "animation_name", "animation_episode",
                "price_ani", "count_ani", "total_ani",
                "price_coloring", "count_coloring", "total_coloring"
            ]
        },
        {
            "enabled": True,
            "style": "日文",
            "sheet_name": "旭-LO,二原",
            "columns": [
                "animation_name", "animation_episode",
                "price_1_yuan", "count_1_yuan", "total_1_yuan",
                "price_2_yuan", "count_2_yuan", "total_2_yuan"
            ]
        }
    ],
    "MSJ": [
        {
            "enabled": True,
            "style": "中文",
            "sheet_name": "",
            "columns": [
                "company_name", "animation_name", "animation_episode",
                "count_ani", "count_coloring", "count_1_yuan", "count_2_yuan"
            ]
        }
    ]
}