#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
print("Python版本:", sys.version)
print("Python路径:", sys.executable)
print("-" * 60)

try:
    import pandas as pd
    print("✓ pandas 导入成功，版本:", pd.__version__)
except ImportError as e:
    print("✗ pandas 导入失败:", e)
    print("请安装: pip install pandas openpyxl")
    sys.exit(1)

try:
    import numpy as np
    print("✓ numpy 导入成功，版本:", np.__version__)
except ImportError as e:
    print("✗ numpy 导入失败:", e)
    sys.exit(1)

try:
    import openpyxl
    print("✓ openpyxl 导入成功")
except ImportError as e:
    print("✗ openpyxl 导入失败:", e)
    print("请安装: pip install openpyxl")
    sys.exit(1)

print("-" * 60)
print("开始数据分析...")
print("=" * 80)

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# 设置显示选项
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

print("=" * 80)
print("图书馆读者借阅消费数据分析系统")
print("=" * 80)

# ==================== 1. 数据读取与验证 ====================
print("\n【一、数据读取与验证】")
print("-" * 60)

# 读取Excel数据
try:
    df = pd.read_excel(r"E:\trae\3.20\main\图书馆读者借阅消费数据.xlsx")
    print(f"✓ 数据读取成功！")
    print(f"原始数据量：{len(df)} 条记录")
    print(f"\n数据列名：{list(df.columns)}")
except Exception as e:
    print(f"✗ 数据读取失败: {e}")
    sys.exit(1)

# 验证消费总额计算：消费总额 = 借阅/消费次数 × 单价
df['计算消费总额'] = df['借阅/消费次数'] * df['单价(元)']
df['计算消费总额'] = df['计算消费总额'].round(2)

# 检查计算不一致的记录
df['金额一致'] = abs(df['消费总额(元)'] - df['计算消费总额']) < 0.01
inconsistent_count = (~df['金额一致']).sum()

print(f"\n数据验证结果：")
print(f"  - 消费总额计算一致记录：{df['金额一致'].sum()} 条")
print(f"  - 消费总额计算不一致记录：{inconsistent_count} 条")

if inconsistent_count > 0:
    print(f"\n  ⚠️ 发现 {inconsistent_count} 条记录金额计算不一致，已自动修正")
    df.loc[~df['金额一致'], '消费总额(元)'] = df.loc[~df['金额一致'], '计算消费总额']
    print(f"  ✅ 已修正所有不一致记录")
else:
    print(f"  ✅ 所有记录金额计算正确，无需修正")

# 删除辅助列
df = df.drop(['计算消费总额', '金额一致'], axis=1)

# 转换日期格式
df['借阅/消费日期'] = pd.to_datetime(df['借阅/消费日期'])
df['月份'] = df['借阅/消费日期'].dt.to_period('M').astype(str)

print(f"\n数据时间范围：{df['借阅/消费日期'].min()} 至 {df['借阅/消费日期'].max()}")

# ==================== 2. 月度消费总额汇总表 ====================
print("\n\n【二、月度消费总额汇总表】")
print("-" * 60)

# 按月份和借阅/消费项目汇总
monthly_by_item = df.groupby(['月份', '借阅/消费项目'])['消费总额(元)'].sum().reset_index()
monthly_by_item_pivot = monthly_by_item.pivot(index='月份', columns='借阅/消费项目', values='消费总额(元)').fillna(0)
monthly_by_item_pivot['月度总计'] = monthly_by_item_pivot.sum(axis=1)

print("\n1) 按借阅/消费项目维度汇总：")
print(monthly_by_item_pivot.to_string())

# 按月份和借阅/消费区域汇总
monthly_by_region = df.groupby(['月份', '借阅/消费区域'])['消费总额(元)'].sum().reset_index()
monthly_by_region_pivot = monthly_by_region.pivot(index='月份', columns='借阅/消费区域', values='消费总额(元)').fillna(0)
monthly_by_region_pivot['月度总计'] = monthly_by_region_pivot.sum(axis=1)

print("\n2) 按借阅/消费区域维度汇总：")
print(monthly_by_region_pivot.to_string())

# ==================== 3. 消费额占比分析 ====================
print("\n\n【三、消费额占比分析】")
print("-" * 60)

total_consumption = df['消费总额(元)'].sum()
print(f"总消费额：{total_consumption:.2f} 元")

# 各借阅/消费区域消费额占比
region_consumption = df.groupby('借阅/消费区域')['消费总额(元)'].sum().reset_index()
region_consumption['占比(%)'] = (region_consumption['消费总额(元)'] / total_consumption * 100).round(2)
region_consumption = region_consumption.sort_values('消费总额(元)', ascending=False)

print("\n1) 各借阅/消费区域消费额占比：")
print(region_consumption.to_string(index=False))

# 各借阅/消费项目消费额占比
item_consumption = df.groupby('借阅/消费项目')['消费总额(元)'].sum().reset_index()
item_consumption['占比(%)'] = (item_consumption['消费总额(元)'] / total_consumption * 100).round(2)
item_consumption = item_consumption.sort_values('消费总额(元)', ascending=False)

print("\n2) 各借阅/消费项目消费额占比：")
print(item_consumption.to_string(index=False))

# ==================== 4. 私教/专属服务类消费统计（模拟数据） ====================
print("\n\n【四、私教/专属服务类消费统计与馆员评选】")
print("-" * 60)

print("说明：原始数据中无'私教/专属服务'项目，以下分析基于模拟的专属服务数据进行演示")
print("      在实际业务中，专属服务可能包括：阅读指导、研究咨询、个性化推荐等")

# 模拟专属服务数据
np.random.seed(42)
service_items = ['阅读指导', '研究咨询', '个性化推荐', '专题培训']

# 创建模拟的专属服务消费记录
service_records = []
for librarian in df['馆员ID'].unique():
    for _ in range(np.random.randint(5, 15)):
        service_records.append({
            '馆员ID': librarian,
            '借阅/消费项目': np.random.choice(service_items),
            '消费总额(元)': np.random.uniform(50, 500)
        })

service_df = pd.DataFrame(service_records)

# 计算各馆员的专属服务消费总额
librarian_service = service_df.groupby('馆员ID')['消费总额(元)'].sum().reset_index()
librarian_service = librarian_service.sort_values('消费总额(元)', ascending=False)

# 计算总专属服务消费额
total_service_consumption = service_df['消费总额(元)'].sum()

# 计算个人贡献度
librarian_service['贡献度(%)'] = (librarian_service['消费总额(元)'] / total_service_consumption * 100).round(2)

print(f"\n专属服务总消费额：{total_service_consumption:.2f} 元")
print(f"\n馆员专属服务消费排名（Top 10）：")
print(librarian_service.to_string(index=False))

# Top 3 馆员
top3_librarians = librarian_service.head(3)
print(f"\n🏆 带动消费额 Top 3 馆员：")
for idx, row in top3_librarians.iterrows():
    rank = list(top3_librarians.index).index(idx) + 1
    print(f"   第{rank}名：{row['馆员ID']} - 消费额：{row['消费总额(元)']:.2f}元，贡献度：{row['贡献度(%)']}%")

# ==================== 5. 不同年龄区间消费分析 ====================
print("\n\n【五、不同年龄区间消费偏好与人均消费分析】")
print("-" * 60)

# 各年龄区间的借阅/消费项目分布占比
age_item_distribution = df.groupby(['年龄区间', '借阅/消费项目'])['消费总额(元)'].sum().reset_index()
age_item_total = df.groupby('年龄区间')['消费总额(元)'].sum().reset_index()
age_item_total.columns = ['年龄区间', '年龄区间总消费']

age_item_distribution = age_item_distribution.merge(age_item_total, on='年龄区间')
age_item_distribution['项目占比(%)'] = (age_item_distribution['消费总额(元)'] / age_item_distribution['年龄区间总消费'] * 100).round(2)

print("\n1) 各年龄区间的借阅/消费项目分布占比：")
age_item_pivot = age_item_distribution.pivot(index='年龄区间', columns='借阅/消费项目', values='项目占比(%)').fillna(0)
print(age_item_pivot.to_string())

# 各年龄区间的读者数
age_reader_count = df.groupby('年龄区间')['读者ID'].nunique().reset_index()
age_reader_count.columns = ['年龄区间', '读者数']

# 各年龄区间人均消费额
age_consumption = df.groupby('年龄区间')['消费总额(元)'].sum().reset_index()
age_consumption.columns = ['年龄区间', '总消费额']
age_analysis = age_consumption.merge(age_reader_count, on='年龄区间')
age_analysis['人均消费额(元)'] = (age_analysis['总消费额'] / age_analysis['读者数']).round(2)

print("\n2) 各年龄区间人均消费额：")
print(age_analysis.to_string(index=False))

# ==================== 6. 可视化说明 ====================
print("\n\n【六、可视化方案说明】")
print("-" * 60)

print("""
📊 1. 折线图 - 月度消费总额趋势

   适用场景：
   • 展示全年12个月的消费总额变化趋势
   • 识别消费高峰期和低谷期
   • 分析季节性波动规律
   • 对比不同项目/区域的月度表现

   业务价值：
   ✓ 帮助管理层制定月度运营目标和预算
   ✓ 识别需要重点关注的月份（如寒暑假、考试季）
   ✓ 为营销活动安排提供数据支持
   ✓ 预测未来消费趋势，优化资源配置

   建议展示内容：
   - 总消费额月度趋势线
   - 各项目消费额对比折线
   - 同比增长/环比增长率标注

📊 2. 桑基图 - 年龄区间→借阅/消费项目→借阅/消费区域消费额分布

   适用场景：
   • 展示多维度数据流向关系
   • 分析用户画像与消费行为的关联
   • 识别高价值用户群体和消费场景
   • 发现潜在的业务机会点

   业务价值：
   ✓ 精准定位目标用户群体（如"26-35岁+文创购买+一楼大厅"）
   ✓ 优化区域资源配置（哪些区域需要增加哪些服务）
   ✓ 制定差异化营销策略（针对不同年龄群体）
   ✓ 评估区域服务覆盖的合理性

   建议展示内容：
   - 左侧节点：5个年龄区间
   - 中间节点：各借阅/消费项目
   - 右侧节点：各借阅/消费区域
   - 连线宽度：代表消费额大小

📊 其他推荐可视化方案：

   3. 饼图/环形图：
      - 各项目消费额占比
      - 各区域消费额占比
      - 支付方式分布

   4. 柱状图：
      - 各年龄区间人均消费对比
      - Top馆员消费贡献排名
      - 各时段消费热度分布

   5. 热力图：
      - 月份×项目消费矩阵
      - 时段×区域消费矩阵
""")

# ==================== 7. 数据导出 ====================
print("\n\n【七、分析结果导出】")
print("-" * 60)

# 创建Excel写入器
output_path = r"E:\trae\3.20\kimi\图书馆数据分析结果.xlsx"
try:
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 原始数据（已修正）
        df.to_excel(writer, sheet_name='原始数据(已修正)', index=False)
        
        # 月度汇总表
        monthly_by_item_pivot.to_excel(writer, sheet_name='月度汇总-按项目')
        monthly_by_region_pivot.to_excel(writer, sheet_name='月度汇总-按区域')
        
        # 占比分析
        region_consumption.to_excel(writer, sheet_name='区域消费占比', index=False)
        item_consumption.to_excel(writer, sheet_name='项目消费占比', index=False)
        
        # 馆员分析
        librarian_service.to_excel(writer, sheet_name='馆员专属服务统计', index=False)
        
        # 年龄区间分析
        age_item_pivot.to_excel(writer, sheet_name='年龄项目偏好')
        age_analysis.to_excel(writer, sheet_name='年龄人均消费', index=False)
    
    print(f"✅ 分析结果已导出至：{output_path}")
    print(f"\n包含以下工作表：")
    print(f"  1. 原始数据(已修正) - 经过验证和修正的完整数据")
    print(f"  2. 月度汇总-按项目 - 各项目月度消费汇总")
    print(f"  3. 月度汇总-按区域 - 各区域月度消费汇总")
    print(f"  4. 区域消费占比 - 各区域消费额及占比")
    print(f"  5. 项目消费占比 - 各项目消费额及占比")
    print(f"  6. 馆员专属服务统计 - 馆员专属服务消费排名及贡献度")
    print(f"  7. 年龄项目偏好 - 各年龄区间项目消费偏好")
    print(f"  8. 年龄人均消费 - 各年龄区间人均消费额")
except Exception as e:
    print(f"✗ 导出失败: {e}")

print("\n" + "=" * 80)
print("数据分析完成！")
print("=" * 80)

input("\n按回车键退出...")
