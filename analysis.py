import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import plotly.io as pio
import os

plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'SimSun']
plt.rcParams['axes.unicode_minus'] = False

output_dir = r"E:\trae\3.20\glm\charts"
os.makedirs(output_dir, exist_ok=True)

print("=" * 80)
print("图书馆读者借阅消费数据分析报告")
print("=" * 80)

# 读取数据
df = pd.read_excel(r"E:\trae\3.20\main\图书馆读者借阅消费数据.xlsx")
print(f"\n【数据概览】")
print(f"总记录数：{len(df)} 条")
print(f"数据列：{list(df.columns)}")

# ============================================================
# 一、数据处理
# ============================================================
print("\n" + "=" * 80)
print("一、数据处理")
print("=" * 80)

# 1. 验证并修正消费总额
print("\n1. 验证消费总额（消费总额 = 借阅/消费次数 × 单价）")
df['计算总额'] = round(df['借阅/消费次数'] * df['单价(元)'], 2)
df['总额差异'] = df['消费总额(元)'] - df['计算总额']

inconsistent = df[df['总额差异'] != 0]
if len(inconsistent) > 0:
    print(f"   发现 {len(inconsistent)} 条记录消费总额计算不一致：")
    print(inconsistent[['读者ID', '借阅/消费项目', '借阅/消费次数', '单价(元)', '消费总额(元)', '计算总额', '总额差异']].to_string())
    print(f"\n   正在修正...")
    df['消费总额(元)'] = df['计算总额']
    print(f"   ✅ 已修正完成！")
else:
    print(f"   ✅ 所有记录消费总额计算正确，无需修正。")

df.drop(['计算总额', '总额差异'], axis=1, inplace=True)

# 2. 提取月份
df['月份'] = pd.to_datetime(df['借阅/消费日期']).dt.month
df['月份名称'] = pd.to_datetime(df['借阅/消费日期']).dt.strftime('%Y-%m')

# 3. 按月份聚合 - 按借阅/消费项目维度
print("\n2. 月度消费总额汇总表（按借阅/消费项目维度）")
monthly_by_item = df.pivot_table(
    values='消费总额(元)',
    index='月份名称',
    columns='借阅/消费项目',
    aggfunc='sum',
    fill_value=0
)
monthly_by_item['月度合计'] = monthly_by_item.sum(axis=1)
print(monthly_by_item.round(2).to_string())

# 4. 按月份聚合 - 按借阅/消费区域维度
print("\n3. 月度消费总额汇总表（按借阅/消费区域维度）")
monthly_by_region = df.pivot_table(
    values='消费总额(元)',
    index='月份名称',
    columns='借阅/消费区域',
    aggfunc='sum',
    fill_value=0
)
monthly_by_region['月度合计'] = monthly_by_region.sum(axis=1)
print(monthly_by_region.round(2).to_string())

# ============================================================
# 二、统计分析
# ============================================================
print("\n" + "=" * 80)
print("二、统计分析")
print("=" * 80)

# 1. 各借阅/消费区域的消费额占比
print("\n1. 各借阅/消费区域消费额占比")
total_consumption = df['消费总额(元)'].sum()
region_consumption = df.groupby('借阅/消费区域')['消费总额(元)'].sum().sort_values(ascending=False)
region_ratio = (region_consumption / total_consumption * 100).round(2)
print(f"   总消费额：{total_consumption:.2f} 元")
print("\n   各区域消费额及占比：")
for region, amount in region_consumption.items():
    print(f"   {region}：{amount:.2f} 元 ({region_ratio[region]:.2f}%)")

# 2. 各借阅/消费项目的消费额占比
print("\n2. 各借阅/消费项目消费额占比")
item_consumption = df.groupby('借阅/消费项目')['消费总额(元)'].sum().sort_values(ascending=False)
item_ratio = (item_consumption / total_consumption * 100).round(2)
print("\n   各项目消费额及占比：")
for item, amount in item_consumption.items():
    print(f"   {item}：{amount:.2f} 元 ({item_ratio[item]:.2f}%)")

# 3. 私教/专属服务类消费 - Top3馆员评选
print("\n3. 私教/专属服务类消费额 Top3 馆员评选")
print("   （注：根据数据中的项目类型，'专题讲座'属于私教/专属服务类消费）")

exclusive_service = df[df['借阅/消费项目'] == '专题讲座']
exclusive_total = exclusive_service['消费总额(元)'].sum()

if len(exclusive_service) > 0:
    librarian_consumption = exclusive_service.groupby('馆员ID')['消费总额(元)'].sum().sort_values(ascending=False)
    top3_librarians = librarian_consumption.head(3)
    
    print(f"\n   专属服务消费总额：{exclusive_total:.2f} 元")
    print(f"\n   Top3 馆员消费额贡献度：")
    for rank, (librarian, amount) in enumerate(top3_librarians.items(), 1):
        contribution = (amount / exclusive_total * 100) if exclusive_total > 0 else 0
        print(f"   第{rank}名：{librarian} - 消费额 {amount:.2f} 元，贡献度 {contribution:.2f}%")
else:
    print("   未找到私教/专属服务类消费记录")

# 4. 不同年龄区间读者的消费偏好及人均消费额
print("\n4. 不同年龄区间读者消费偏好及人均消费额")

# 4.1 各年龄区间消费项目分布占比
print("\n   4.1 各年龄区间消费偏好（消费项目分布占比）")
age_item_pivot = df.pivot_table(
    values='消费总额(元)',
    index='年龄区间',
    columns='借阅/消费项目',
    aggfunc='sum',
    fill_value=0
)
age_item_total = age_item_pivot.sum(axis=1)
age_item_ratio = (age_item_pivot.div(age_item_total, axis=0) * 100).round(2)
print(age_item_ratio.to_string())

# 4.2 各年龄区间人均消费额
print("\n   4.2 各年龄区间人均消费额")
age_consumption = df.groupby('年龄区间')['消费总额(元)'].sum()
age_readers = df.groupby('年龄区间')['读者ID'].nunique()
age_avg_consumption = (age_consumption / age_readers).round(2)

print(f"\n   {'年龄区间':<15} {'读者数':<10} {'总消费额':<15} {'人均消费额':<15}")
print(f"   {'-'*55}")
for age in age_consumption.index:
    print(f"   {age:<15} {age_readers[age]:<10} {age_consumption[age]:<15.2f} {age_avg_consumption[age]:<15.2f}")

# ============================================================
# 三、可视化说明
# ============================================================
print("\n" + "=" * 80)
print("三、可视化说明")
print("=" * 80)

print("\n1. 折线图展示「月度消费总额趋势」")
print("-" * 60)
print("""
【适用场景】
   - 展示时间序列数据的连续变化趋势
   - 观察消费额随时间的增长或下降规律
   - 识别消费的周期性特征（如节假日高峰、淡季低谷）
   - 对比不同年份同期数据的变化

【业务价值】
   - 帮助管理层了解图书馆经营状况的时间变化
   - 为预算编制和资源配置提供数据支撑
   - 识别异常波动，及时发现问题并调整运营策略
   - 预测未来消费趋势，制定营销活动计划
   - 评估促销活动或新服务上线后的效果

【实现要点】
   - X轴：月份（1-12月）
   - Y轴：消费总额（元）
   - 可添加趋势线或移动平均线平滑数据
   - 可叠加多条折线对比不同项目/区域的月度趋势
""")

print("\n2. 桑基图展示「年龄区间→借阅/消费项目→借阅/消费区域」消费额分布")
print("-" * 60)
print("""
【适用场景】
   - 展示多层级数据的流向和转化关系
   - 分析不同群体（年龄区间）的消费行为路径
   - 揭示消费从"谁"到"做什么"到"在哪里"的完整链条
   - 发现隐藏的消费模式和关联关系

【业务价值】
   - 精准识别不同年龄读者的消费偏好和活动区域
   - 优化图书馆空间布局和服务配置
   - 为差异化营销提供依据（如针对特定年龄段推广特定服务）
   - 发现高价值消费路径，重点优化关键环节
   - 辅助决策资源投放优先级

【实现要点】
   - 左侧节点：年龄区间（5个分组）
   - 中间节点：借阅/消费项目（5种类型）
   - 右侧节点：借阅/消费区域（5个区域）
   - 连线宽度：代表消费额大小
   - 颜色区分：不同类型节点使用不同配色
""")

print("\n" + "=" * 80)
print("数据分析完成！")
print("=" * 80)

# ============================================================
# 四、生成可视化图表
# ============================================================
print("\n" + "=" * 80)
print("四、生成可视化图表")
print("=" * 80)

# 1. 折线图 - 月度消费总额趋势
print("\n1. 生成折线图：月度消费总额趋势")
monthly_total = df.groupby('月份名称')['消费总额(元)'].sum().sort_index()

fig1, ax1 = plt.subplots(figsize=(12, 6))
ax1.plot(monthly_total.index, monthly_total.values, marker='o', linewidth=2, markersize=8, color='#2E86AB')
ax1.fill_between(monthly_total.index, monthly_total.values, alpha=0.3, color='#2E86AB')
ax1.set_xlabel('月份', fontsize=12)
ax1.set_ylabel('消费总额（元）', fontsize=12)
ax1.set_title('图书馆月度消费总额趋势（2025年）', fontsize=14, fontweight='bold')
ax1.grid(True, linestyle='--', alpha=0.7)
ax1.tick_params(axis='x', rotation=45)
for i, (x, y) in enumerate(zip(monthly_total.index, monthly_total.values)):
    ax1.annotate(f'{y:.0f}', (x, y), textcoords="offset points", xytext=(0, 10), ha='center', fontsize=9)
plt.tight_layout()
chart1_path = os.path.join(output_dir, '月度消费总额趋势.png')
fig1.savefig(chart1_path, dpi=150, bbox_inches='tight')
plt.close()
print(f"   ✅ 已保存：{chart1_path}")

# 2. 桑基图 - 年龄区间→借阅/消费项目→借阅/消费区域
print("\n2. 生成桑基图：年龄区间→借阅/消费项目→借阅/消费区域")

sankey_data = df.groupby(['年龄区间', '借阅/消费项目', '借阅/消费区域'])['消费总额(元)'].sum().reset_index()
sankey_data = sankey_data[sankey_data['消费总额(元)'] > 0]

age_groups = ['18岁及以下', '19-25岁', '26-35岁', '36-45岁', '46岁及以上']
items = ['图书借阅', '文献传递', '打印复印', '文创购买', '专题讲座']
regions = ['一楼大厅', '二楼社科区', '三楼自科区', '四楼电子阅览区', '五楼特藏区']

all_labels = age_groups + items + regions
label_to_idx = {label: idx for idx, label in enumerate(all_labels)}

source = []
target = []
value = []

for _, row in sankey_data.iterrows():
    age_idx = label_to_idx[row['年龄区间']]
    item_idx = label_to_idx[row['借阅/消费项目']]
    region_idx = label_to_idx[row['借阅/消费区域']]
    
    source.append(age_idx)
    target.append(item_idx)
    value.append(row['消费总额(元)'])
    
    source.append(item_idx)
    target.append(region_idx)
    value.append(row['消费总额(元)'])

colors_age = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7']
colors_item = ['#DDA0DD', '#98D8C8', '#F7DC6F', '#BB8FCE', '#85C1E9']
colors_region = ['#E8DAEF', '#D5F5E3', '#FCF3CF', '#FADBD8', '#D6EAF8']
node_colors = colors_age + colors_item + colors_region

fig2 = go.Figure(data=[go.Sankey(
    node=dict(
        pad=15,
        thickness=20,
        line=dict(color="black", width=0.5),
        label=all_labels,
        color=node_colors
    ),
    link=dict(
        source=source,
        target=target,
        value=value,
        color='rgba(150,150,150,0.3)'
    )
)])

fig2.update_layout(
    title_text="图书馆消费额流向桑基图<br><sub>年龄区间 → 借阅/消费项目 → 借阅/消费区域</sub>",
    font_size=12,
    width=1200,
    height=700
)

chart2_path = os.path.join(output_dir, '消费额流向桑基图.html')
pio.write_html(fig2, chart2_path)
print(f"   ✅ 已保存：{chart2_path}")

print(f"\n所有图表已保存至文件夹：{output_dir}")
print("=" * 80)
