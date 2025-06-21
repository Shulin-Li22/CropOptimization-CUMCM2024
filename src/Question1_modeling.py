import pandas as pd
import numpy as np
from pulp import *
import warnings

warnings.filterwarnings('ignore')


class CorrectedCropOptimizer:
    def __init__(self, data_file='processed_data.xlsx'):
        """初始化优化器，读取预处理数据"""
        print("正在加载数据...")

        # 读取预处理数据
        self.land_df = pd.read_excel(data_file, sheet_name='地块信息')
        self.crop_df = pd.read_excel(data_file, sheet_name='作物信息')
        self.stats_df = pd.read_excel(data_file, sheet_name='作物统计数据')
        self.planting_2023_df = pd.read_excel(data_file, sheet_name='2023年种植情况')
        self.expected_sales_df = pd.read_excel(data_file, sheet_name='预期销售量')

        print(f"数据加载完成：{len(self.land_df)}个地块，{len(self.crop_df)}种作物")

        # 处理数据
        self._process_data()

    def _process_data(self):
        """处理和组织数据"""
        print("正在处理数据...")

        # 创建地块信息字典
        self.lands = {}
        for _, row in self.land_df.iterrows():
            self.lands[row['地块名称']] = {
                'type': row['地块类型'],
                'area': row['地块面积(亩)']
            }

        # 创建作物信息字典
        self.crops = {}
        for _, row in self.crop_df.iterrows():
            crop_name = row['作物名称']
            crop_type = row['作物类型']

            is_grain = '粮食' in crop_type and crop_name != '水稻'
            is_rice = crop_name == '水稻'
            is_vegetable = '蔬菜' in crop_type
            is_mushroom = crop_name in ['香菇', '羊肚菌', '白灵菇', '榆黄菇']
            is_winter_vegetable = crop_name in ['大白菜', '白萝卜', '红萝卜']
            is_legume = row['是否豆类']

            self.crops[row['作物编号']] = {
                'name': crop_name,
                'type': crop_type,
                'is_legume': is_legume,
                'is_grain': is_grain,
                'is_rice': is_rice,
                'is_vegetable': is_vegetable,
                'is_mushroom': is_mushroom,
                'is_winter_vegetable': is_winter_vegetable,
                'is_regular_vegetable': is_vegetable and not is_winter_vegetable
            }

        # 创建预期销售量字典
        self.expected_sales = {}
        for _, row in self.expected_sales_df.iterrows():
            self.expected_sales[row['作物编号']] = row['预期销售量(斤)']

        # 处理统计数据
        self.valid_planting_options = []

        for _, row in self.stats_df.iterrows():
            land_type = row['land_type']
            season = row['season']
            crop_id = row['crop_id']

            if crop_id not in self.crops:
                continue

            crop_info = self.crops[crop_id]

            if self._is_valid_planting_combination(land_type, season, crop_info):
                self.valid_planting_options.append({
                    'land_type': land_type,
                    'season': season,
                    'crop_id': crop_id,
                    'crop_name': crop_info['name'],
                    'yield_per_mu': row['yield_per_mu'],
                    'cost_per_mu': row['cost_per_mu'],
                    'price_avg': row['price_avg'],
                    'profit_per_mu': row['profit_per_mu'],
                    **crop_info
                })

        # 筛选有预期销售量的可行选项
        self.viable_options = [opt for opt in self.valid_planting_options
                               if opt['crop_id'] in self.expected_sales and
                               self.expected_sales[opt['crop_id']] > 0]

        self.viable_options.sort(key=lambda x: x['profit_per_mu'], reverse=True)

        # 获取2023年豆类种植情况
        self.legume_planted_2023 = {}
        for _, row in self.planting_2023_df.iterrows():
            land_name = row['block_name']
            crop_id = row['crop_id']

            if land_name not in self.legume_planted_2023:
                self.legume_planted_2023[land_name] = False

            if crop_id in self.crops and self.crops[crop_id]['is_legume']:
                self.legume_planted_2023[land_name] = True

        print(f"数据处理完成，找到{len(self.viable_options)}个符合规则的可行种植选项")

        # 计算理论最大种植面积
        total_land_area = sum(info['area'] for info in self.lands.values())
        total_seasons = total_land_area * 7  # 7年
        print(f"总耕地面积: {total_land_area}亩")
        print(f"7年可种植总面积: {total_seasons}亩")

    def _is_valid_planting_combination(self, land_type, season, crop_info):
        """检查地块类型、季次和作物的组合是否符合种植规则"""

        if land_type in ['平旱地', '梯田', '山坡地']:
            return season == '单季' and crop_info['is_grain']

        elif land_type == '水浇地':
            if season == '单季':
                return crop_info['is_rice']
            elif season == '第一季':
                return crop_info['is_regular_vegetable']
            elif season == '第二季':
                return crop_info['is_winter_vegetable']

        elif land_type in ['普通大棚', '普通大棚 ']:
            if season == '第一季':
                return crop_info['is_regular_vegetable']
            elif season == '第二季':
                return crop_info['is_mushroom']

        elif land_type == '智慧大棚':
            if season in ['第一季', '第二季']:
                return crop_info['is_regular_vegetable']

        return False

    def create_model(self, scenario=1):
        """创建优化模型"""
        print(f"创建优化模型 - 场景{scenario}")

        prob = LpProblem("Crop_Optimization_Fixed", LpMaximize)

        years = list(range(2024, 2031))

        # 决策变量
        var_keys = []
        for land_name, land_info in self.lands.items():
            land_type = land_info['type']
            for year in years:
                for opt in self.viable_options:
                    if opt['land_type'] == land_type:
                        var_keys.append((land_name, year, opt['season'], opt['crop_id']))

        x = LpVariable.dicts("plant", var_keys, lowBound=0, cat='Continuous')

        # 水浇地选择二进制变量
        water_land_choice = {}
        for land_name, land_info in self.lands.items():
            if land_info['type'] == '水浇地':
                for year in years:
                    water_land_choice[(land_name, year)] = LpVariable(
                        f"use_rice_{land_name}_{year}", cat='Binary'
                    )

        # 重茬控制二进制变量
        planting_binary = {}
        for land_name in self.lands.keys():
            for year in years:
                for season in ['单季', '第一季', '第二季']:
                    for crop_id in self.expected_sales.keys():
                        planting_binary[(land_name, year, season, crop_id)] = LpVariable(
                            f"plant_binary_{land_name}_{year}_{season}_{crop_id}", cat='Binary'
                        )

        # 目标函数
        total_profit = 0
        for (land_name, year, season, crop_id) in var_keys:
            opt = next((o for o in self.viable_options
                        if o['land_type'] == self.lands[land_name]['type'] and
                        o['season'] == season and o['crop_id'] == crop_id), None)

            if opt:
                base_profit = opt['profit_per_mu']

                # 给豆类作物更多激励
                if crop_id in self.crops and self.crops[crop_id]['is_legume']:
                    base_profit += 200  # 增加激励

                total_profit += x[(land_name, year, season, crop_id)] * base_profit

        prob += total_profit

        # 添加约束条件
        self._add_fixed_constraints(prob, x, water_land_choice, planting_binary, years, scenario)

        return prob, x, water_land_choice, planting_binary

    def _add_fixed_constraints(self, prob, x, water_land_choice, planting_binary, years, scenario):
        """添加约束条件"""
        print("添加约束条件...")

        # 1. 地块面积约束
        for land_name, land_info in self.lands.items():
            max_area = land_info['area']
            land_type = land_info['type']

            for year in years:
                if land_type != '水浇地':
                    for season in ['单季', '第一季', '第二季']:
                        season_vars = [x[(ln, y, s, crop_id)]
                                       for (ln, y, s, crop_id) in x.keys()
                                       if ln == land_name and y == year and s == season]
                        if season_vars:
                            prob += lpSum(season_vars) <= max_area
                else:
                    # 水浇地约束处理
                    single_vars = [x[(ln, y, s, crop_id)]
                                   for (ln, y, s, crop_id) in x.keys()
                                   if ln == land_name and y == year and s == '单季']

                    first_vars = [x[(ln, y, s, crop_id)]
                                  for (ln, y, s, crop_id) in x.keys()
                                  if ln == land_name and y == year and s == '第一季']

                    second_vars = [x[(ln, y, s, crop_id)]
                                   for (ln, y, s, crop_id) in x.keys()
                                   if ln == land_name and y == year and s == '第二季']

                    if single_vars:
                        prob += lpSum(single_vars) <= max_area
                    if first_vars:
                        prob += lpSum(first_vars) <= max_area
                    if second_vars:
                        prob += lpSum(second_vars) <= max_area

        # 2. 销售量约束 - 7年总量约束
        for crop_id, expected_sale in self.expected_sales.items():
            total_production_vars = []
            for year in years:
                for season in ['单季', '第一季', '第二季']:
                    for (land_name, y, s, c_id) in x.keys():
                        if y == year and s == season and c_id == crop_id:
                            opt = next((o for o in self.viable_options
                                        if o['land_type'] == self.lands[land_name]['type'] and
                                        o['season'] == s and o['crop_id'] == c_id), None)
                            if opt:
                                total_production_vars.append(x[(land_name, y, s, c_id)] * opt['yield_per_mu'])

            if total_production_vars:
                total_production = lpSum(total_production_vars)

                if scenario == 1:
                    # 场景1：7年总产量不超过7倍预期销售量
                    prob += total_production <= expected_sale * 7
                else:
                    # 场景2：允许适度超产
                    prob += total_production <= expected_sale * 10

        # 3. 重茬约束 - 使用二进制变量
        M = 1000  # 大M常数

        for land_name in self.lands.keys():
            for crop_id in self.expected_sales.keys():
                for year in years[:-1]:
                    for season in ['单季', '第一季', '第二季']:
                        current_key = (land_name, year, season, crop_id)
                        next_key = (land_name, year + 1, season, crop_id)

                        if current_key in planting_binary and next_key in planting_binary:
                            current_binary = planting_binary[current_key]
                            next_binary = planting_binary[next_key]

                            # 重茬约束：不能连续两年种植同一作物
                            prob += current_binary + next_binary <= 1

                            # 连接二进制变量和连续变量
                            current_vars = [x[(ln, y, s, c_id)]
                                            for (ln, y, s, c_id) in x.keys()
                                            if ln == land_name and y == year and s == season and c_id == crop_id]

                            next_vars = [x[(ln, y, s, c_id)]
                                         for (ln, y, s, c_id) in x.keys()
                                         if ln == land_name and y == year + 1 and s == season and c_id == crop_id]

                            if current_vars:
                                prob += lpSum(current_vars) <= M * current_binary
                                prob += lpSum(current_vars) >= 0.01 * current_binary

                            if next_vars:
                                prob += lpSum(next_vars) <= M * next_binary
                                prob += lpSum(next_vars) >= 0.01 * next_binary

        # 4. 水浇地选择约束
        for land_name, land_info in self.lands.items():
            if land_info['type'] == '水浇地':
                for year in years:
                    if (land_name, year) in water_land_choice:
                        use_rice = water_land_choice[(land_name, year)]
                        M_water = land_info['area']

                        single_vars = [x[(ln, y, s, crop_id)]
                                       for (ln, y, s, crop_id) in x.keys()
                                       if ln == land_name and y == year and s == '单季']

                        multi_vars = [x[(ln, y, s, crop_id)]
                                      for (ln, y, s, crop_id) in x.keys()
                                      if ln == land_name and y == year and s in ['第一季', '第二季']]

                        if single_vars:
                            prob += lpSum(single_vars) <= M_water * use_rice
                        if multi_vars:
                            prob += lpSum(multi_vars) <= M_water * (1 - use_rice)

                        # 两季蔬菜面积相等约束
                        first_vars = [x[(ln, y, s, crop_id)]
                                      for (ln, y, s, crop_id) in x.keys()
                                      if ln == land_name and y == year and s == '第一季']

                        second_vars = [x[(ln, y, s, crop_id)]
                                       for (ln, y, s, crop_id) in x.keys()
                                       if ln == land_name and y == year and s == '第二季']

                        if first_vars and second_vars:
                            prob += lpSum(first_vars) == lpSum(second_vars)

        # 5. 简化的豆类轮作约束
        for land_name, land_info in self.lands.items():
            # 如果2023年已种豆类，跳过
            if self.legume_planted_2023.get(land_name, False):
                continue

            # 2024-2026年必须种植豆类（按比例要求）
            legume_vars = []
            for year in [2024, 2025, 2026]:
                if year in years:
                    for (ln, y, s, crop_id) in x.keys():
                        if (ln == land_name and y == year and
                                crop_id in self.crops and self.crops[crop_id]['is_legume']):
                            legume_vars.append(x[(ln, y, s, crop_id)])

            if legume_vars:
                # 要求至少种植地块面积的5%
                min_legume_area = max(0.1, land_info['area'] * 0.05)
                prob += lpSum(legume_vars) >= min_legume_area

        # 6. 添加最小种植面积约束，避免碎片化
        for (land_name, year, season, crop_id) in x.keys():
            # 如果种植，至少种植0.1亩
            if (land_name, year, season, crop_id) in planting_binary:
                binary_var = planting_binary[(land_name, year, season, crop_id)]
                prob += x[(land_name, year, season, crop_id)] >= 0.1 * binary_var
                prob += x[(land_name, year, season, crop_id)] <= self.lands[land_name]['area'] * binary_var

        print("约束条件添加完成")

    def solve_and_save(self, scenario=1, output_file=None):
        """求解模型并保存结果"""
        if output_file is None:
            output_file = f'result1_{scenario}_fixed.xlsx'

        print(f"开始求解场景{scenario}...")

        years = list(range(2024, 2031))

        prob, x, water_land_choice, planting_binary = self.create_model(scenario)

        # 求解
        print("正在求解...")
        solver = PULP_CBC_CMD(msg=False, timeLimit=1200)  # 增加时间限制
        prob.solve(solver)

        status = LpStatus[prob.status]
        print(f"求解状态: {status}")

        if status not in ['Optimal', 'Feasible']:
            print(f"警告：求解状态为 {status}")
            # 尝试放松约束
            print("尝试放松约束重新求解...")
            return self._solve_relaxed_model(scenario, output_file)

        # 提取结果
        results = []
        total_profit = 0

        for (land_name, year, season, crop_id) in x.keys():
            area = x[(land_name, year, season, crop_id)].varValue
            if area and area > 0.01:
                land_type = self.lands[land_name]['type']
                crop_name = self.crops[crop_id]['name']

                opt = next((o for o in self.viable_options
                            if o['land_type'] == land_type and o['season'] == season and o['crop_id'] == crop_id),
                           None)

                if opt:
                    production = area * opt['yield_per_mu']
                    cost = area * opt['cost_per_mu']
                    expected_sale = self.expected_sales.get(crop_id, 0)

                    # 计算收入
                    if scenario == 1:
                        # 场景1：按年度销售量限制
                        yearly_expected = expected_sale  # 每年的预期销售量
                        sold_amount = min(production, yearly_expected)
                        revenue = sold_amount * opt['price_avg']
                    else:
                        # 场景2：超产部分按50%价格
                        yearly_expected = expected_sale
                        normal_sale = min(production, yearly_expected)
                        excess_sale = max(0, production - yearly_expected)
                        revenue = normal_sale * opt['price_avg'] + excess_sale * opt['price_avg'] * 0.5

                    profit = revenue - cost
                    total_profit += profit

                    results.append({
                        '年份': year,
                        '地块名称': land_name,
                        '地块类型': land_type,
                        '种植季次': season,
                        '作物编号': crop_id,
                        '作物名称': crop_name,
                        '作物分类': self._get_crop_category(crop_id),
                        '种植面积': round(area, 2),
                        '产量': round(production, 1),
                        '成本': round(cost, 1),
                        '收入': round(revenue, 1),
                        '利润': round(profit, 1)
                    })

        print(f"总利润: {total_profit:,.1f}元")
        print(f"找到{len(results)}个种植方案")

        if results:
            # 计算总种植面积
            total_area = sum(r['种植面积'] for r in results)
            available_area = sum(info['area'] for info in self.lands.values()) * 7
            utilization = total_area / available_area * 100

            print(f"总种植面积: {total_area:.1f}亩")
            print(f"可用面积: {available_area:.1f}亩")
            print(f"土地利用率: {utilization:.1f}%")

            self._save_results_with_validation(results, total_profit, scenario, output_file, water_land_choice, years)
            return results, total_profit
        else:
            print("未找到可行解")
            return None, 0

    def _solve_relaxed_model(self, scenario, output_file):
        """求解放松约束的模型"""
        print("创建放松约束的模型...")

        prob = LpProblem("Crop_Optimization_Relaxed", LpMaximize)
        years = list(range(2024, 2031))

        # 简化的决策变量
        var_keys = []
        for land_name, land_info in self.lands.items():
            land_type = land_info['type']
            for year in years:
                for opt in self.viable_options:
                    if opt['land_type'] == land_type:
                        var_keys.append((land_name, year, opt['season'], opt['crop_id']))

        x = LpVariable.dicts("plant", var_keys, lowBound=0, cat='Continuous')

        # 简化的目标函数
        total_profit = 0
        for (land_name, year, season, crop_id) in var_keys:
            opt = next((o for o in self.viable_options
                        if o['land_type'] == self.lands[land_name]['type'] and
                        o['season'] == season and o['crop_id'] == crop_id), None)

            if opt:
                base_profit = opt['profit_per_mu']
                if crop_id in self.crops and self.crops[crop_id]['is_legume']:
                    base_profit += 100
                total_profit += x[(land_name, year, season, crop_id)] * base_profit

        prob += total_profit

        # 只添加最基本的约束
        # 1. 地块面积约束
        for land_name, land_info in self.lands.items():
            max_area = land_info['area']
            land_type = land_info['type']

            for year in years:
                for season in ['单季', '第一季', '第二季']:
                    season_vars = [x[(ln, y, s, crop_id)]
                                   for (ln, y, s, crop_id) in x.keys()
                                   if ln == land_name and y == year and s == season]
                    if season_vars:
                        prob += lpSum(season_vars) <= max_area

        # 2. 放松的销售量约束
        for crop_id, expected_sale in self.expected_sales.items():
            total_production_vars = []
            for (land_name, year, season, c_id) in x.keys():
                if c_id == crop_id:
                    opt = next((o for o in self.viable_options
                                if o['land_type'] == self.lands[land_name]['type'] and
                                o['season'] == season and o['crop_id'] == c_id), None)
                    if opt:
                        total_production_vars.append(x[(land_name, year, season, c_id)] * opt['yield_per_mu'])

            if total_production_vars:
                # 非常宽松的约束
                prob += lpSum(total_production_vars) <= expected_sale * 20

        # 求解
        solver = PULP_CBC_CMD(msg=False, timeLimit=600)
        prob.solve(solver)

        status = LpStatus[prob.status]
        print(f"放松模型求解状态: {status}")

        if status in ['Optimal', 'Feasible']:
            # 提取并保存结果
            return self._extract_and_save_results(x, scenario, output_file, years)
        else:
            print("即使放松约束也无法求解")
            return None, 0

    def _extract_and_save_results(self, x, scenario, output_file, years):
        """提取并保存结果"""
        results = []
        total_profit = 0

        for (land_name, year, season, crop_id) in x.keys():
            area = x[(land_name, year, season, crop_id)].varValue
            if area and area > 0.01:
                land_type = self.lands[land_name]['type']
                crop_name = self.crops[crop_id]['name']

                opt = next((o for o in self.viable_options
                            if o['land_type'] == land_type and o['season'] == season and o['crop_id'] == crop_id),
                           None)

                if opt:
                    production = area * opt['yield_per_mu']
                    cost = area * opt['cost_per_mu']
                    revenue = production * opt['price_avg']
                    profit = revenue - cost
                    total_profit += profit

                    results.append({
                        '年份': year,
                        '地块名称': land_name,
                        '地块类型': land_type,
                        '种植季次': season,
                        '作物编号': crop_id,
                        '作物名称': crop_name,
                        '作物分类': self._get_crop_category(crop_id),
                        '种植面积': round(area, 2),
                        '产量': round(production, 1),
                        '成本': round(cost, 1),
                        '收入': round(revenue, 1),
                        '利润': round(profit, 1)
                    })

        if results:
            total_area = sum(r['种植面积'] for r in results)
            print(f"放松模型总种植面积: {total_area:.1f}亩")
            print(f"放松模型总利润: {total_profit:,.1f}元")

            self._save_simple_results(results, total_profit, scenario, output_file)
            return results, total_profit

        return None, 0

    def _get_crop_category(self, crop_id):
        """获取作物分类标签"""
        crop_info = self.crops[crop_id]
        if crop_info['is_rice']:
            return '水稻'
        elif crop_info['is_grain']:
            return '粮食类'
        elif crop_info['is_winter_vegetable']:
            return '冬季蔬菜'
        elif crop_info['is_mushroom']:
            return '食用菌'
        elif crop_info['is_regular_vegetable']:
            return '普通蔬菜'
        else:
            return '其他'

    def _save_results_with_validation(self, results, total_profit, scenario, output_file, water_land_choice, years):
        """保存结果并添加验证"""
        results_df = pd.DataFrame(results)

        # 创建汇总
        yearly_summary = results_df.groupby('年份').agg({
            '种植面积': 'sum',
            '产量': 'sum',
            '成本': 'sum',
            '收入': 'sum',
            '利润': 'sum'
        }).round(1).reset_index()

        crop_summary = results_df.groupby(['作物编号', '作物名称', '作物分类']).agg({
            '种植面积': 'sum',
            '产量': 'sum',
            '成本': 'sum',
            '收入': 'sum',
            '利润': 'sum'
        }).round(1).reset_index()
        crop_summary['利润率%'] = (crop_summary['利润'] / crop_summary['成本'] * 100).round(1)
        crop_summary = crop_summary.sort_values('利润', ascending=False)

        # 保存到Excel
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            results_df.to_excel(writer, sheet_name='种植方案', index=False)
            yearly_summary.to_excel(writer, sheet_name='年度汇总', index=False)
            crop_summary.to_excel(writer, sheet_name='作物汇总', index=False)

            # 总体信息
            info_df = pd.DataFrame([{
                '场景': f'场景{scenario}',
                '规划期间': '2024-2030年',
                '总利润': f'{total_profit:,.1f}元',
                '平均年利润': f'{total_profit / 7:,.1f}元',
                '总种植面积': f'{sum(r["种植面积"] for r in results):,.1f}亩',
                '土地利用率': f'{sum(r["种植面积"] for r in results) / (sum(info["area"] for info in self.lands.values()) * 7) * 100:.1f}%',
                '解法说明': '超产滞销' if scenario == 1 else '超产降价50%'
            }])
            info_df.to_excel(writer, sheet_name='总体信息', index=False)

        print(f"结果已保存到: {output_file}")

    def _save_simple_results(self, results, total_profit, scenario, output_file):
        """保存简化结果"""
        results_df = pd.DataFrame(results)

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            results_df.to_excel(writer, sheet_name='种植方案', index=False)

            info_df = pd.DataFrame([{
                '场景': f'场景{scenario}',
                '规划期间': '2024-2030年',
                '总利润': f'{total_profit:,.1f}元',
                '总种植面积': f'{sum(r["种植面积"] for r in results):,.1f}亩',
                '说明': '使用放松约束求解'
            }])
            info_df.to_excel(writer, sheet_name='总体信息', index=False)

        print(f"放松模型结果已保存到: {output_file}")

    def run_all_scenarios(self):
        """运行所有场景"""
        print("=" * 60)
        print("农作物种植优化 - 问题一求解")
        print("=" * 60)

        results = {}

        # 场景1
        print("\n" + "=" * 40)
        print("场景1：超产部分滞销浪费")
        print("=" * 40)
        result1, profit1 = self.solve_and_save(scenario=1, output_file='result1_1_fixed.xlsx')
        results['scenario1'] = {'result': result1, 'profit': profit1}

        # 场景2
        print("\n" + "=" * 40)
        print("场景2：超产部分降价50%销售")
        print("=" * 40)
        result2, profit2 = self.solve_and_save(scenario=2, output_file='result1_2_fixed.xlsx')
        results['scenario2'] = {'result': result2, 'profit': profit2}

        return results


def main():
    """主函数"""
    try:
        optimizer = CorrectedCropOptimizer('processed_data.xlsx')
        results = optimizer.run_all_scenarios()

        print("\n" + "=" * 60)
        print("问题一求解完成！")
        print("输出文件:")
        print("  - result1_1.xlsx: 场景1结果")
        print("  - result1_2.xlsx: 场景2结果")
        print("=" * 60)

        return results

    except Exception as e:
        print(f"求解过程中出现错误: {e}")
        import traceback
        traceback.print_exc()
        return None


if __name__ == "__main__":
    results = main()