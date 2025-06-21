import pandas as pd
import numpy as np
import re
import warnings

warnings.filterwarnings('ignore')


def load_and_clean_data(file1_path='附件1.xlsx', file2_path='附件2.xlsx'):
    """
    读取并清理Excel数据
    """
    print("正在读取Excel文件...")

    # 读取附件1
    try:
        land_df = pd.read_excel(file1_path, sheet_name='乡村的现有耕地')
        crops_df = pd.read_excel(file1_path, sheet_name='乡村种植的农作物')
        print(f"附件1读取成功: 地块数据{len(land_df)}条, 作物信息{len(crops_df)}条")
    except Exception as e:
        print(f"读取附件1失败: {e}")
        return None

    # 读取附件2
    try:
        planting_2023_df = pd.read_excel(file2_path, sheet_name='2023年的农作物种植情况')
        statistics_df = pd.read_excel(file2_path, sheet_name='2023年统计的相关数据')
        print(f"附件2读取成功: 2023年种植{len(planting_2023_df)}条, 统计数据{len(statistics_df)}条")
    except Exception as e:
        print(f"读取附件2失败: {e}")
        return None

    # 数据清洗
    land_df = land_df.dropna(subset=['地块名称']).copy()
    crops_df = crops_df.dropna(subset=['作物编号']).copy()
    planting_2023_df = planting_2023_df.dropna(subset=['种植地块']).copy()
    statistics_df = statistics_df.dropna(subset=['作物编号']).copy()

    return land_df, crops_df, planting_2023_df, statistics_df


def process_land_data(land_df):
    """
    处理地块数据
    """
    print("处理地块数据...")

    # 处理所有地块（包括大棚）
    land_info = {}
    for _, row in land_df.iterrows():
        land_info[row['地块名称']] = {
            'type': row['地块类型'].strip(),  # 去除可能的空格
            'area': float(row['地块面积/亩'])
        }

    print(f"地块数据处理完成，共{len(land_info)}个地块")

    # 统计各类型地块
    land_type_count = {}
    for info in land_info.values():
        land_type = info['type']
        if land_type not in land_type_count:
            land_type_count[land_type] = 0
        land_type_count[land_type] += 1

    print("地块类型分布:")
    for land_type, count in land_type_count.items():
        print(f"  {land_type}: {count}个")

    return land_info


def process_crop_data(crops_df):
    """
    处理作物基本信息
    """
    print("处理作物基本信息...")

    crop_info = {}
    for _, row in crops_df.iterrows():
        if pd.isna(row['作物编号']):
            continue

        crop_id = int(row['作物编号'])
        crop_info[crop_id] = {
            'name': row['作物名称'],
            'type': row['作物类型'],
            'is_legume': '豆类' in str(row['作物类型'])
        }

    print(f"作物信息处理完成，共{len(crop_info)}种作物")
    return crop_info


def parse_price_range(price_str):
    """
    解析价格区间字符串，返回最小值和最大值
    """
    try:
        price_str = str(price_str).replace('元', '').replace('/', '').strip()

        # 查找价格区间模式 (如 "2.50-4.00")
        match = re.search(r'(\d+\.?\d*)-(\d+\.?\d*)', price_str)
        if match:
            price_min = float(match.group(1))
            price_max = float(match.group(2))
            return price_min, price_max
        else:
            # 如果没有区间，尝试解析单个数字
            price = float(re.search(r'\d+\.?\d*', price_str).group())
            return price, price
    except:
        return 0.0, 0.0


def process_statistics_data(statistics_df):
    """
    处理作物统计数据，计算经济效益
    """
    print("处理作物统计数据...")

    statistics_data = []

    for _, row in statistics_df.iterrows():
        if pd.isna(row['作物编号']):
            continue

        crop_id = int(row['作物编号'])
        crop_name = row['作物名称']
        land_type = row['地块类型']
        season = row['种植季次']
        yield_per_mu = float(row['亩产量/斤'])
        cost_per_mu = float(row['种植成本/(元/亩)'])
        price_range = str(row['销售单价/(元/斤)'])

        # 解析价格
        price_min, price_max = parse_price_range(price_range)
        price_avg = (price_min + price_max) / 2

        # 计算经济指标
        revenue_per_mu = yield_per_mu * price_avg
        profit_per_mu = revenue_per_mu - cost_per_mu
        profit_rate = (profit_per_mu / cost_per_mu * 100) if cost_per_mu > 0 else 0

        statistics_data.append({
            'crop_id': crop_id,
            'crop_name': crop_name,
            'land_type': land_type,
            'season': season,
            'yield_per_mu': yield_per_mu,
            'cost_per_mu': cost_per_mu,
            'price_min': price_min,
            'price_max': price_max,
            'price_avg': price_avg,
            'revenue_per_mu': revenue_per_mu,
            'profit_per_mu': profit_per_mu,
            'profit_rate': profit_rate
        })

    # 补充智慧大棚第一季数据（与普通大棚相同）
    statistics_data = supplement_smart_greenhouse_data(statistics_data)

    print(f"统计数据处理完成，共{len(statistics_data)}条记录")
    return statistics_data


def supplement_smart_greenhouse_data(statistics_data):
    """
    补充智慧大棚第一季的数据（与普通大棚第一季相同）
    根据题目说明：智慧大棚第一季可种植的蔬菜作物及其亩产量、种植成本和销售价格均与普通大棚相同
    """
    print("补充智慧大棚第一季数据...")

    # 找出普通大棚第一季的蔬菜数据（注意处理空格问题）
    greenhouse_first_season = []
    for stat in statistics_data:
        land_type = str(stat['land_type']).strip()  # 去除空格
        season = str(stat['season']).strip()

        if (land_type == '普通大棚' and
                season == '第一季'):
            # 排除食用菌，只要蔬菜类作物
            crop_name = stat['crop_name']
            if '菇' not in crop_name and '菌' not in crop_name:
                greenhouse_first_season.append(stat)

    print(f"找到普通大棚第一季作物: {len(greenhouse_first_season)}种")

    # 为智慧大棚补充相同的数据
    supplemented_data = []
    for stat in greenhouse_first_season:
        smart_stat = stat.copy()
        smart_stat['land_type'] = '智慧大棚'  # 统一格式，不加空格
        supplemented_data.append(smart_stat)

    original_count = len(statistics_data)
    final_data = statistics_data + supplemented_data

    print(f"为智慧大棚补充了{len(supplemented_data)}条第一季数据")
    if supplemented_data:
        crop_names = [stat['crop_name'] for stat in supplemented_data]
        print(f"补充的作物: {', '.join(crop_names[:8])}{'...' if len(crop_names) > 8 else ''}")

    return final_data


def process_planting_2023(planting_2023_df):
    """
    处理2023年种植情况
    """
    print("处理2023年种植情况...")

    planting_data = []
    crop_total_area = {}

    for _, row in planting_2023_df.iterrows():
        block_name = row['种植地块']
        crop_id = int(row['作物编号'])
        crop_name = row['作物名称']
        crop_type = row['作物类型']
        area = float(row['种植面积/亩'])
        season = row['种植季次']

        planting_data.append({
            'block_name': block_name,
            'crop_id': crop_id,
            'crop_name': crop_name,
            'crop_type': crop_type,
            'area': area,
            'season': season
        })

        # 统计各作物总种植面积
        if crop_name not in crop_total_area:
            crop_total_area[crop_name] = 0
        crop_total_area[crop_name] += area

    print(f"2023年种植数据处理完成，共{len(planting_data)}条记录")
    return planting_data, crop_total_area


def calculate_expected_sales(planting_data, statistics_data, land_info):
    """
    基于2023年种植情况计算预期销售量
    """
    print("计算预期销售量...")

    # 创建统计数据查找字典
    stats_lookup = {}
    for stat in statistics_data:
        key = (stat['crop_id'], stat['land_type'], stat['season'])
        stats_lookup[key] = stat

    expected_sales = {}

    for planting in planting_data:
        crop_id = planting['crop_id']
        crop_name = planting['crop_name']
        block_name = planting['block_name']
        area = planting['area']
        season = planting['season']

        # 获取地块类型
        if block_name in land_info:
            land_type = land_info[block_name]['type']
        else:
            land_type = '平旱地'  # 默认值

        # 查找对应的亩产量
        key = (crop_id, land_type, season)
        if key in stats_lookup:
            yield_per_mu = stats_lookup[key]['yield_per_mu']
            total_yield = area * yield_per_mu

            if crop_id not in expected_sales:
                expected_sales[crop_id] = 0
            expected_sales[crop_id] += total_yield

    print(f"预期销售量计算完成，涉及{len(expected_sales)}种作物")
    return expected_sales


def save_processed_data(land_info, crop_info, statistics_data, planting_data,
                        crop_total_area, expected_sales, output_file='processed_data.xlsx'):
    """
    保存处理后的数据到Excel文件
    """
    print("保存处理后的数据...")

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

        # 1. 地块信息
        land_df = pd.DataFrame([
            {
                '地块名称': name,
                '地块类型': info['type'],
                '地块面积(亩)': info['area']
            }
            for name, info in land_info.items()
        ])
        land_df.to_excel(writer, sheet_name='地块信息', index=False)

        # 2. 作物信息
        crops_df = pd.DataFrame([
            {
                '作物编号': crop_id,
                '作物名称': info['name'],
                '作物类型': info['type'],
                '是否豆类': info['is_legume']
            }
            for crop_id, info in crop_info.items()
        ])
        crops_df.to_excel(writer, sheet_name='作物信息', index=False)

        # 3. 作物统计数据（包含经济指标）
        stats_df = pd.DataFrame(statistics_data)
        cols_order = ['crop_id', 'crop_name', 'land_type', 'season', 'yield_per_mu',
                      'cost_per_mu', 'price_min', 'price_max', 'price_avg',
                      'revenue_per_mu', 'profit_per_mu', 'profit_rate']
        stats_df = stats_df[cols_order]
        stats_df.to_excel(writer, sheet_name='作物统计数据', index=False)

        # 4. 2023年种植情况
        planting_df = pd.DataFrame(planting_data)
        planting_df.to_excel(writer, sheet_name='2023年种植情况', index=False)

        # 5. 各作物种植面积汇总
        area_summary = pd.DataFrame([
            {
                '作物名称': crop_name,
                '2023年种植面积(亩)': area
            }
            for crop_name, area in sorted(crop_total_area.items(), key=lambda x: x[1], reverse=True)
        ])
        area_summary.to_excel(writer, sheet_name='种植面积汇总', index=False)

        # 6. 预期销售量
        sales_df = pd.DataFrame([
            {
                '作物编号': crop_id,
                '作物名称': crop_info[crop_id]['name'] if crop_id in crop_info else f'作物{crop_id}',
                '预期销售量(斤)': sales
            }
            for crop_id, sales in expected_sales.items()
        ])
        sales_df.to_excel(writer, sheet_name='预期销售量', index=False)

        # 7. 地块类型统计
        land_type_stats = {}
        for name, info in land_info.items():
            land_type = info['type']
            if land_type not in land_type_stats:
                land_type_stats[land_type] = {'count': 0, 'total_area': 0}
            land_type_stats[land_type]['count'] += 1
            land_type_stats[land_type]['total_area'] += info['area']

        land_stats_df = pd.DataFrame([
            {
                '地块类型': land_type,
                '地块数量': stats['count'],
                '总面积(亩)': stats['total_area'],
                '平均面积(亩)': stats['total_area'] / stats['count']
            }
            for land_type, stats in land_type_stats.items()
        ])
        land_stats_df.to_excel(writer, sheet_name='地块类型统计', index=False)

        # 8. 数据补充说明
        supplement_info = pd.DataFrame([
            {
                '数据类型': '智慧大棚补充数据',
                '说明': '根据题目要求，智慧大棚第一季数据与普通大棚相同',
                '补充方法': '自动复制普通大棚第一季蔬菜作物数据',
                '数据来源': '基于普通大棚第一季统计数据'
            }
        ])
        supplement_info.to_excel(writer, sheet_name='数据补充说明', index=False)

    print(f"数据已保存到 {output_file}")


def main():
    """
    主函数：执行完整的数据预处理流程
    """
    print("=" * 60)
    print("农作物种植策略 - 数据预处理")
    print("=" * 60)

    # 1. 加载数据
    data = load_and_clean_data()
    if data is None:
        print("数据加载失败，程序退出")
        return

    land_df, crops_df, planting_2023_df, statistics_df = data

    # 2. 处理各部分数据
    land_info = process_land_data(land_df)
    crop_info = process_crop_data(crops_df)
    statistics_data = process_statistics_data(statistics_df)
    planting_data, crop_total_area = process_planting_2023(planting_2023_df)

    # 3. 计算预期销售量
    expected_sales = calculate_expected_sales(planting_data, statistics_data, land_info)

    # 4. 输出数据概览
    print("\n" + "=" * 60)
    print("数据处理结果概览")
    print("=" * 60)

    total_area = sum(info['area'] for info in land_info.values())
    print(f"地块总数: {len(land_info)} 个")
    print(f"总耕地面积: {total_area:.1f} 亩")
    print(f"作物种类: {len(crop_info)} 种")
    print(f"统计数据: {len(statistics_data)} 条")
    print(f"2023年种植记录: {len(planting_data)} 条")

    # 地块类型分布
    print(f"\n地块类型分布:")
    land_type_count = {}
    for info in land_info.values():
        land_type = info['type']
        if land_type not in land_type_count:
            land_type_count[land_type] = {'count': 0, 'area': 0}
        land_type_count[land_type]['count'] += 1
        land_type_count[land_type]['area'] += info['area']

    for land_type, stats in land_type_count.items():
        percentage = (stats['area'] / total_area) * 100
        print(f"  {land_type}: {stats['count']}个, {stats['area']:.1f}亩 ({percentage:.1f}%)")

    # 种植面积前10的作物
    print(f"\n2023年种植面积前10的作物:")
    sorted_crops = sorted(crop_total_area.items(), key=lambda x: x[1], reverse=True)
    for i, (crop_name, area) in enumerate(sorted_crops[:10]):
        percentage = (area / total_area) * 100
        print(f"  {i + 1:2d}. {crop_name}: {area:.1f}亩 ({percentage:.1f}%)")

    # 5. 保存数据
    save_processed_data(land_info, crop_info, statistics_data, planting_data,
                        crop_total_area, expected_sales)

    print(f"\n" + "=" * 60)
    print("数据预处理完成！")
    print("输出文件: processed_data.xlsx")
    print("=" * 60)


if __name__ == "__main__":
    main()