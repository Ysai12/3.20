import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

# 固定随机种子，保证每次生成的数据完全一致
np.random.seed(42)
random.seed(42)

# 配置参数：数据量在150-200之间（这里取180条，可自行修改）
num_records = 180
start_date = datetime(2025, 1, 1)
end_date = datetime(2025, 12, 31)

# 枚举选项（贴合图书馆实际业务）
gender_list = ["男", "女"]
age_range_list = ["18岁及以下", "19-25岁", "26-35岁", "36-45岁", "46岁及以上"]
card_type_list = ["普通读者卡", "学生卡", "教职工卡", "VIP卡"]
item_list = ["图书借阅", "文献传递", "打印复印", "文创购买", "专题讲座"]
region_list = ["一楼大厅", "二楼社科区", "三楼自科区", "四楼电子阅览区", "五楼特藏区"]
payment_list = ["微信支付", "支付宝", "读者卡余额", "现金"]
status_list = ["已完成", "已取消"]
librarian_list = [f"L{str(i).zfill(3)}" for i in range(1, 11)]  # 馆员ID：L001-L010


# 生成单条读者消费/借阅记录
def generate_record(rid):
    # 基础信息
    gender = random.choice(gender_list)
    age_range = random.choice(age_range_list)
    card_type = random.choice(card_type_list)
    item = random.choice(item_list)
    times = random.randint(1, 10)  # 消费/借阅次数

    # 不同项目的单价（贴合实际：借阅免费，消费项目有对应价格）
    price_map = {
        "图书借阅": 0.0,  # 图书借阅免费
        "文献传递": 2.0,  # 文献传递2元/次
        "打印复印": 0.5,  # 打印复印0.5元/次
        "文创购买": random.uniform(10, 100),  # 文创10-100元/件
        "专题讲座": 30.0  # 讲座30元/次
    }
    price = price_map[item]
    total = round(times * price, 2)  # 计算消费总额（自动验证：次数×单价）

    # 时间信息
    random_days = random.randint(0, (end_date - start_date).days)
    trans_date = (start_date + timedelta(days=random_days)).strftime("%Y-%m-%d")
    hour = random.randint(8, 21)  # 图书馆营业时间8:00-22:00
    time_slot = f"{hour}:00-{hour + 1}:00"

    # 其他信息
    librarian = random.choice(librarian_list)
    region = random.choice(region_list)
    payment = random.choice(payment_list)
    status = random.choice(status_list)

    return {
        "读者ID": f"R{str(rid).zfill(4)}",
        "读者姓名": f"读者{str(rid).zfill(4)}",
        "性别": gender,
        "年龄区间": age_range,
        "读者卡类型": card_type,
        "借阅/消费项目": item,
        "借阅/消费次数": times,
        "单价(元)": round(price, 2),
        "消费总额(元)": total,
        "借阅/消费日期": trans_date,
        "借阅/消费时段": time_slot,
        "馆员ID": librarian,
        "借阅/消费区域": region,
        "支付方式": payment,
        "借阅/消费状态": status
    }


# 生成所有数据并导出为Excel
if __name__ == "__main__":
    # 生成180条数据（可修改num_records调整数量）
    data = [generate_record(i + 1) for i in range(num_records)]
    df = pd.DataFrame(data)

    # 导出Excel（index=False 不生成行号）
    excel_path = "图书馆读者借阅消费数据.xlsx"
    df.to_excel(excel_path, index=False, engine="openpyxl")

    print(f"✅ 数据生成完成！")
    print(f"📊 生成记录数：{num_records} 条")
    print(f"📁 文件保存路径：{excel_path}")
    # 预览前5条数据，方便核对
    print("\n📌 数据预览（前5条）：")
    print(df.head())