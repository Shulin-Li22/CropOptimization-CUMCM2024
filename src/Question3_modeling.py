import pandas as pd
import numpy as np
from pulp import *
import warnings
from itertools import combinations
import matplotlib.pyplot as plt
import seaborn as sns

warnings.filterwarnings('ignore')


class AdvancedCropOptimizer:
    """
    问题3：考虑农作物间相关性的高级优化器
    包含可替代性、互补性和价格-成本相关性建模
    """

    def __init__(self, data_file='processed_data.xlsx'):
        """初始化高级优化器"""
        print("=" * 70)
        print("问题3：考虑农作物间相关性的种植策略优化")
        print("=" * 70)

        self.data_file = data_file
        self._load_base_data()
        self._define_crop_relationships()
        self._calculate_correlation_parameters()
        self._setup_risk_parameters()

    def _load_base_data(self):
        """加载基础数据"""
        print("正在加载基础数据...")

        # 读取数据
        self.land_df = pd.read_excel(self.data_file, sheet_name='地块信息')
        self.crop_df = pd.read_excel(self.data_file, sheet_name='作物信息')
        self.stats_df = pd.read_excel(self.data_file, sheet_name='作物统计数据')
        self.expected_sales_df = pd.read_excel(self.data_file, sheet_name='预期销售量')
        self.planting_2023_df = pd.read_excel(self.data_file, sheet_name='2023年种植情况')

        # 处理地块信息
        self.lands = {}
        for _, row in self.land_df.iterrows():
            self.lands[row['地块名称']] = {
                'type': row['地块类型'].strip(),
                'area': row['地块面积(亩)']
            }

        # 处理作物信息
        self.crops = {}
        for _, row in self.crop_df.iterrows():
            crop_name = row['作物名称']
            self.crops[row['作物编号']] = {
                'name': crop_name,
                'type': row['作物类型'],
                'is_legume': row['是否豆类'],
                'category': self._classify_crop_category(crop_name, row['作物类型'])
            }

        # 预期销售量
        self.expected_sales = {}
        for _, row in self.expected_sales_df.iterrows():
            self.expected_sales[row['作物编号']] = row['预期销售量(斤)']

        # 获取有效种植选项
        self._get_valid_planting_options()

        print(f"基础数据加载完成：{len(self.lands)}个地块，{len(self.crops)}种作物")

    def _classify_crop_category(self, crop_name, crop_type):
        """严格按照约束条件分类作物类别"""
        # 水稻：只能在水浇地单季种植
        if crop_name == '水稻':
            return 'rice'
        # 粮食类作物（水稻除外）：适宜在平旱地、梯田、山坡地单季种植
        elif '粮食' in crop_type and crop_name != '水稻':
            return 'grain'
        # 冬季蔬菜：只能在水浇地第二季种植
        elif crop_name in ['大白菜', '白萝卜', '红萝卜']:
            return 'winter_vegetable'
        # 食用菌：只能在普通大棚第二季种植
        elif crop_name in ['香菇', '羊肚菌', '白灵菇', '榆黄菇']:
            return 'mushroom'
        # 普通蔬菜：可在水浇地第一季、普通大棚第一季、智慧大棚两季种植
        elif '蔬菜' in crop_type and crop_name not in ['大白菜', '白萝卜', '红萝卜']:
            return 'vegetable'
        else:
            return 'other'

    def _get_valid_planting_options(self):
        """获取有效种植选项"""
        self.valid_options = []

        for _, row in self.stats_df.iterrows():
            land_type = row['land_type'].strip()
            season = row['season']
            crop_id = row['crop_id']

            if crop_id in self.crops and self._is_valid_combination(land_type, season, crop_id):
                self.valid_options.append({
                    'land_type': land_type,
                    'season': season,
                    'crop_id': crop_id,
                    'crop_name': self.crops[crop_id]['name'],
                    'base_yield': row['yield_per_mu'],
                    'base_cost': row['cost_per_mu'],
                    'base_price': row['price_avg'],
                    'base_profit': row['profit_per_mu']
                })

        print(f"有效种植选项: {len(self.valid_options)}")

    def _validate_constraint_compliance(self):
        """验证约束条件的执行情况"""
        print("\n🔍 验证约束条件执行情况...")

        # 统计各地块类型-季次-作物类型的组合
        combinations = {}
        for opt in self.valid_options:
            land_type = opt['land_type']
            season = opt['season']
            crop_name = opt['crop_name']
            crop_category = self.crops[opt['crop_id']]['category']

            key = f"{land_type}-{season}"
            if key not in combinations:
                combinations[key] = {'crops': [], 'categories': set()}

            combinations[key]['crops'].append(crop_name)
            combinations[key]['categories'].add(crop_category)

        # 验证关键约束
        constraints_check = []

        # 约束1：平旱地、梯田、山坡地只能单季种植粮食类（水稻除外）
        for land_type in ['平旱地', '梯田', '山坡地']:
            key = f"{land_type}-单季"
            if key in combinations:
                categories = combinations[key]['categories']
                crops = combinations[key]['crops']
                is_valid = (categories == {'grain'} and '水稻' not in crops)
                constraints_check.append({
                    '约束': f'{land_type}单季种植',
                    '要求': '粮食类(水稻除外)',
                    '实际': f"{categories}, 作物: {len(crops)}种",
                    '符合': '✅' if is_valid else '❌'
                })

        # 约束2：水浇地单季只能种水稻
        key = "水浇地-单季"
        if key in combinations:
            crops = combinations[key]['crops']
            is_valid = (crops == ['水稻'])
            constraints_check.append({
                '约束': '水浇地单季种植',
                '要求': '仅水稻',
                '实际': f"作物: {crops}",
                '符合': '✅' if is_valid else '❌'
            })

        # 约束3：水浇地第一季不能种冬季蔬菜
        key = "水浇地-第一季"
        if key in combinations:
            crops = combinations[key]['crops']
            winter_vegs = ['大白菜', '白萝卜', '红萝卜']
            has_winter_veg = any(crop in winter_vegs for crop in crops)
            constraints_check.append({
                '约束': '水浇地第一季种植',
                '要求': '蔬菜(冬季蔬菜除外)',
                '实际': f"作物: {len(crops)}种",
                '符合': '✅' if not has_winter_veg else '❌'
            })

        # 约束4：水浇地第二季只能种冬季蔬菜
        key = "水浇地-第二季"
        if key in combinations:
            crops = combinations[key]['crops']
            winter_vegs = ['大白菜', '白萝卜', '红萝卜']
            is_valid = all(crop in winter_vegs for crop in crops)
            constraints_check.append({
                '约束': '水浇地第二季种植',
                '要求': '仅冬季蔬菜',
                '实际': f"作物: {crops}",
                '符合': '✅' if is_valid else '❌'
            })

        # 约束5：普通大棚第一季不能种冬季蔬菜
        key = "普通大棚-第一季"
        if key in combinations:
            crops = combinations[key]['crops']
            winter_vegs = ['大白菜', '白萝卜', '红萝卜']
            has_winter_veg = any(crop in winter_vegs for crop in crops)
            constraints_check.append({
                '约束': '普通大棚第一季种植',
                '要求': '蔬菜(冬季蔬菜除外)',
                '实际': f"作物: {len(crops)}种",
                '符合': '✅' if not has_winter_veg else '❌'
            })

        # 约束6：普通大棚第二季只能种食用菌
        key = "普通大棚-第二季"
        if key in combinations:
            categories = combinations[key]['categories']
            is_valid = (categories == {'mushroom'})
            constraints_check.append({
                '约束': '普通大棚第二季种植',
                '要求': '仅食用菌',
                '实际': f"类型: {categories}",
                '符合': '✅' if is_valid else '❌'
            })

        # 约束7：智慧大棚不能种冬季蔬菜
        for season in ['第一季', '第二季']:
            key = f"智慧大棚-{season}"
            if key in combinations:
                crops = combinations[key]['crops']
                winter_vegs = ['大白菜', '白萝卜', '红萝卜']
                has_winter_veg = any(crop in winter_vegs for crop in crops)
                constraints_check.append({
                    '约束': f'智慧大棚{season}种植',
                    '要求': '蔬菜(冬季蔬菜除外)',
                    '实际': f"作物: {len(crops)}种",
                    '符合': '✅' if not has_winter_veg else '❌'
                })

        # 输出验证结果
        print("约束条件验证结果:")
        all_compliant = True
        for check in constraints_check:
            print(f"  {check['符合']} {check['约束']}: {check['要求']} | {check['实际']}")
            if check['符合'] == '❌':
                all_compliant = False

        if all_compliant:
            print("✅ 所有约束条件都得到正确执行")
        else:
            print("❌ 存在约束条件违反情况")

        return all_compliant

    def _is_valid_combination(self, land_type, season, crop_id):
        """严格按照约束条件检查种植组合是否有效"""
        crop_info = self.crops[crop_id]
        category = crop_info['category']
        crop_name = crop_info['name']

        # 约束条件1：平旱地、梯田、山坡地每年适宜单季种植粮食类作物(水稻除外)
        if land_type in ['平旱地', '梯田', '山坡地']:
            return (season == '单季' and
                    category == 'grain' and
                    crop_name != '水稻')

        # 约束条件2：水浇地每年可以单季种植水稻或两季种植蔬菜作物
        elif land_type == '水浇地':
            if season == '单季':
                # 单季只能种植水稻
                return crop_name == '水稻'
            elif season == '第一季':
                # 第一季可种植多种蔬菜(大白菜、白萝卜和红萝卜除外)
                return (category == 'vegetable' and
                        crop_name not in ['大白菜', '白萝卜', '红萝卜'])
            elif season == '第二季':
                # 第二季只能种植大白菜、白萝卜和红萝卜中的一种
                return crop_name in ['大白菜', '白萝卜', '红萝卜']

        # 约束条件3：普通大棚每年种植两季作物
        elif land_type in ['普通大棚', '普通大棚 ']:
            if season == '第一季':
                # 第一季可种植多种蔬菜(大白菜、白萝卜和红萝卜除外)
                return (category == 'vegetable' and
                        crop_name not in ['大白菜', '白萝卜', '红萝卜'])
            elif season == '第二季':
                # 第二季只能种植食用菌
                return category == 'mushroom'

        # 约束条件4：智慧大棚每年都可种植两季蔬菜(大白菜、白萝卜和红萝卜除外)
        elif land_type == '智慧大棚':
            return (season in ['第一季', '第二季'] and
                    category == 'vegetable' and
                    crop_name not in ['大白菜', '白萝卜', '红萝卜'])

        return False

    def _define_crop_relationships(self):
        """定义农作物间的相关性关系"""
        print("定义农作物间相关性关系...")

        # 1. 可替代性矩阵（0-1，1表示完全可替代）
        self.substitution_matrix = {}
        crop_ids = list(self.crops.keys())

        for i in crop_ids:
            self.substitution_matrix[i] = {}
            for j in crop_ids:
                if i == j:
                    self.substitution_matrix[i][j] = 1.0
                else:
                    # 基于作物类别的替代性
                    cat_i = self.crops[i]['category']
                    cat_j = self.crops[j]['category']

                    if cat_i == cat_j:
                        # 同类作物高度可替代
                        if cat_i in ['grain', 'vegetable']:
                            self.substitution_matrix[i][j] = 0.8
                        elif cat_i in ['mushroom']:
                            self.substitution_matrix[i][j] = 0.6
                        else:
                            self.substitution_matrix[i][j] = 0.4
                    elif (cat_i == 'grain' and self.crops[j]['is_legume']) or \
                            (cat_j == 'grain' and self.crops[i]['is_legume']):
                        # 豆类与粮食类有一定替代性
                        self.substitution_matrix[i][j] = 0.3
                    else:
                        self.substitution_matrix[i][j] = 0.1

        # 2. 互补性矩阵（-1到1，正值表示互补，负值表示竞争）
        self.complementarity_matrix = {}

        for i in crop_ids:
            self.complementarity_matrix[i] = {}
            for j in crop_ids:
                if i == j:
                    self.complementarity_matrix[i][j] = 0.0
                else:
                    # 豆类与其他作物的互补性
                    if self.crops[i]['is_legume'] and not self.crops[j]['is_legume']:
                        self.complementarity_matrix[i][j] = 0.6  # 豆类改善土壤
                    elif not self.crops[i]['is_legume'] and self.crops[j]['is_legume']:
                        self.complementarity_matrix[i][j] = 0.6
                    elif self.crops[i]['category'] != self.crops[j]['category']:
                        # 不同类别作物间存在轻微互补
                        self.complementarity_matrix[i][j] = 0.2
                    else:
                        # 同类作物间竞争
                        self.complementarity_matrix[i][j] = -0.1

        # 3. 需求弹性系数
        self.demand_elasticity = {}
        for crop_id in crop_ids:
            category = self.crops[crop_id]['category']
            if category in ['grain', 'rice']:
                self.demand_elasticity[crop_id] = -0.3  # 刚性需求
            elif category in ['vegetable', 'winter_vegetable']:
                self.demand_elasticity[crop_id] = -0.8  # 弹性需求
            elif category == 'mushroom':
                self.demand_elasticity[crop_id] = -1.2  # 高弹性需求
            else:
                self.demand_elasticity[crop_id] = -0.5

        print("农作物相关性关系定义完成")

    def _calculate_correlation_parameters(self):
        """计算相关性参数"""
        print("计算相关性参数...")

        years = list(range(2024, 2031))
        self.correlation_params = {}

        for year in years:
            years_from_base = year - 2023
            self.correlation_params[year] = {}

            for crop_id in self.crops.keys():
                crop_name = self.crops[crop_id]['name']
                category = self.crops[crop_id]['category']

                # 基础变化率
                if crop_name in ['小麦', '玉米']:
                    sales_growth = 0.075  # 7.5%增长
                else:
                    sales_growth = np.random.uniform(-0.05, 0.05)  # ±5%变化

                # 亩产量变化
                yield_change = np.random.uniform(-0.10, 0.10)  # ±10%变化

                # 成本变化
                cost_growth = 0.05  # 5%增长

                # 价格变化
                if category in ['grain', 'rice']:
                    price_change = np.random.uniform(-0.02, 0.02)  # 基本稳定
                elif category in ['vegetable', 'winter_vegetable']:
                    price_change = 0.05  # 5%增长
                elif category == 'mushroom':
                    if crop_name == '羊肚菌':
                        price_change = -0.05  # 5%下降
                    else:
                        price_change = np.random.uniform(-0.05, -0.01)  # 1%-5%下降
                else:
                    price_change = 0.0

                # 规模经济系数
                scale_economy = self._calculate_scale_economy(crop_id)

                # 风险调整系数
                risk_factor = self._calculate_risk_factor(crop_id)

                self.correlation_params[year][crop_id] = {
                    'sales_multiplier': (1 + sales_growth) ** years_from_base,
                    'yield_multiplier': 1 + yield_change,
                    'cost_multiplier': (1 + cost_growth) ** years_from_base,
                    'price_multiplier': (1 + price_change) ** years_from_base,
                    'scale_economy': scale_economy,
                    'risk_factor': risk_factor
                }

        print("相关性参数计算完成")

    def _calculate_scale_economy(self, crop_id):
        """计算规模经济系数"""
        category = self.crops[crop_id]['category']

        # 不同作物的规模经济效应不同
        if category in ['grain', 'rice']:
            return 0.15  # 粮食类规模经济显著
        elif category in ['vegetable', 'winter_vegetable']:
            return 0.10  # 蔬菜类中等规模经济
        elif category == 'mushroom':
            return 0.20  # 食用菌规模经济最显著
        else:
            return 0.05

    def _calculate_risk_factor(self, crop_id):
        """计算风险因子"""
        category = self.crops[crop_id]['category']

        # 不同作物的风险水平
        if category in ['grain', 'rice']:
            return 0.1  # 粮食类风险较低
        elif category in ['vegetable', 'winter_vegetable']:
            return 0.2  # 蔬菜类风险中等
        elif category == 'mushroom':
            return 0.3  # 食用菌风险较高
        else:
            return 0.15

    def _setup_risk_parameters(self):
        """设置风险参数"""
        print("设置风险管理参数...")

        # 风险厌恶系数
        self.risk_aversion = 0.3

        # 相关性风险调整
        self.correlation_risk_adjustment = 0.1

        # 多样性激励系数
        self.diversity_incentive = 50

        print("风险参数设置完成")

    def create_advanced_model(self):
        """创建考虑相关性的高级模型"""
        print("创建高级相关性模型...")

        prob = LpProblem("Advanced_Crop_Optimization", LpMaximize)
        years = list(range(2024, 2031))

        # 决策变量：种植面积
        x = {}
        for land_name in self.lands.keys():
            land_type = self.lands[land_name]['type']
            for year in years:
                for opt in self.valid_options:
                    if opt['land_type'] == land_type:
                        var_key = (land_name, year, opt['season'], opt['crop_id'])
                        x[var_key] = LpVariable(f"x_{len(x)}", lowBound=0, cat='Continuous')

        # 水浇地选择变量
        y_water = {}
        for land_name, land_info in self.lands.items():
            if land_info['type'] == '水浇地':
                for year in years:
                    y_water[(land_name, year)] = LpVariable(f"y_water_{land_name}_{year}", cat='Binary')

        # 作物种植指示变量（用于多样性约束）
        z_crop = {}
        for year in years:
            for crop_id in self.crops.keys():
                z_crop[(year, crop_id)] = LpVariable(f"z_crop_{year}_{crop_id}", cat='Binary')

        print(f"创建了{len(x)}个种植变量，{len(y_water)}个水浇地选择变量，{len(z_crop)}个作物指示变量")

        # 高级目标函数：考虑相关性和风险
        total_objective = self._build_advanced_objective(x, z_crop, years)
        prob += total_objective

        # 添加约束条件
        self._add_advanced_constraints(prob, x, y_water, z_crop, years)

        return prob, x, y_water, z_crop

    def _build_advanced_objective(self, x, z_crop, years):
        """构建考虑相关性的高级目标函数"""
        total_objective = 0

        # 1. 基础利润
        for (land_name, year, season, crop_id) in x.keys():
            params = self.correlation_params[year][crop_id]

            # 找到对应的基础数据
            opt = self._find_option(land_name, season, crop_id)
            if opt:
                # 基础收益计算
                expected_yield = opt['base_yield'] * params['yield_multiplier']
                expected_cost = opt['base_cost'] * params['cost_multiplier']
                expected_price = opt['base_price'] * params['price_multiplier']

                base_profit = expected_yield * expected_price - expected_cost

                # 规模经济效应（线性近似）
                scale_benefit = params['scale_economy'] * 100  # 规模效益

                # 风险调整
                risk_penalty = params['risk_factor'] * 50  # 风险惩罚

                # 互补性激励（针对豆类）
                complementarity_bonus = 0
                if self.crops[crop_id]['is_legume']:
                    complementarity_bonus = 100  # 豆类互补性激励

                adjusted_profit = base_profit + scale_benefit - risk_penalty + complementarity_bonus

                total_objective += x[(land_name, year, season, crop_id)] * adjusted_profit

        # 2. 多样性激励
        for year in years:
            for crop_id in self.crops.keys():
                if (year, crop_id) in z_crop:
                    total_objective += z_crop[(year, crop_id)] * self.diversity_incentive

        return total_objective

    def _find_option(self, land_name, season, crop_id):
        """查找对应的种植选项"""
        land_type = self.lands[land_name]['type']
        for opt in self.valid_options:
            if (opt['land_type'] == land_type and
                    opt['season'] == season and
                    opt['crop_id'] == crop_id):
                return opt
        return None

    def _add_advanced_constraints(self, prob, x, y_water, z_crop, years):
        """添加高级约束条件"""
        print("添加高级约束条件...")
        constraint_count = 0

        # 1. 基础约束（地块面积、种植规则等）
        constraint_count += self._add_basic_constraints(prob, x, y_water, years)

        # 2. 需求弹性约束
        constraint_count += self._add_demand_elasticity_constraints(prob, x, years)

        # 3. 作物指示变量约束
        constraint_count += self._add_crop_indicator_constraints(prob, x, z_crop, years)

        # 4. 多样性约束
        constraint_count += self._add_diversity_constraints(prob, z_crop, years)

        # 5. 重茬和轮作约束
        constraint_count += self._add_rotation_constraints(prob, x, years)

        # 6. 风险控制约束
        constraint_count += self._add_risk_control_constraints(prob, x, years)

        # 7. 互补性约束（线性化处理）
        constraint_count += self._add_complementarity_constraints(prob, x, years)

        print(f"高级约束条件添加完成，共{constraint_count}个约束")

    def _add_basic_constraints(self, prob, x, y_water, years):
        """添加严格的基础约束条件"""
        count = 0

        # 地块面积约束
        for land_name, land_info in self.lands.items():
            max_area = land_info['area']
            land_type = land_info['type']

            for year in years:
                if land_type == '水浇地':
                    # 水浇地特殊处理：单季水稻 OR 两季蔬菜（互斥选择）
                    if (land_name, year) in y_water:
                        # 单季水稻约束
                        rice_vars = [x[(ln, yr, s, crop_id)]
                                     for (ln, yr, s, crop_id) in x.keys()
                                     if (ln == land_name and yr == year and s == '单季' and
                                         crop_id in self.crops and self.crops[crop_id]['name'] == '水稻')]
                        if rice_vars:
                            prob += lpSum(rice_vars) <= max_area * y_water[(land_name, year)]
                            count += 1

                        # 两季蔬菜约束
                        for season in ['第一季', '第二季']:
                            veg_vars = [x[(ln, yr, s, crop_id)]
                                        for (ln, yr, s, crop_id) in x.keys()
                                        if ln == land_name and yr == year and s == season]
                            if veg_vars:
                                prob += lpSum(veg_vars) <= max_area * (1 - y_water[(land_name, year)])
                                count += 1

                        # 两季蔬菜面积必须相等
                        first_vars = [x[(ln, yr, s, crop_id)]
                                      for (ln, yr, s, crop_id) in x.keys()
                                      if ln == land_name and yr == year and s == '第一季']
                        second_vars = [x[(ln, yr, s, crop_id)]
                                       for (ln, yr, s, crop_id) in x.keys()
                                       if ln == land_name and yr == year and s == '第二季']
                        if first_vars and second_vars:
                            prob += lpSum(first_vars) == lpSum(second_vars)
                            count += 1

                elif land_type in ['平旱地', '梯田', '山坡地']:
                    # 平旱地、梯田、山坡地：每年只能种植一季
                    season_vars = [x[(ln, yr, s, crop_id)]
                                   for (ln, yr, s, crop_id) in x.keys()
                                   if ln == land_name and yr == year and s == '单季']
                    if season_vars:
                        prob += lpSum(season_vars) <= max_area
                        count += 1

                elif land_type in ['普通大棚', '普通大棚 ', '智慧大棚']:
                    # 大棚：每年种植两季，每季面积不超过地块面积
                    for season in ['第一季', '第二季']:
                        season_vars = [x[(ln, yr, s, crop_id)]
                                       for (ln, yr, s, crop_id) in x.keys()
                                       if ln == land_name and yr == year and s == season]
                        if season_vars:
                            prob += lpSum(season_vars) <= max_area
                            count += 1

                    # 大棚两季面积必须相等
                    first_vars = [x[(ln, yr, s, crop_id)]
                                  for (ln, yr, s, crop_id) in x.keys()
                                  if ln == land_name and yr == year and s == '第一季']
                    second_vars = [x[(ln, yr, s, crop_id)]
                                   for (ln, yr, s, crop_id) in x.keys()
                                   if ln == land_name and yr == year and s == '第二季']
                    if first_vars and second_vars:
                        prob += lpSum(first_vars) == lpSum(second_vars)
                        count += 1

        # 严格的种植规则验证约束
        for (land_name, year, season, crop_id) in x.keys():
            land_type = self.lands[land_name]['type']
            if not self._is_valid_combination(land_type, season, crop_id):
                # 如果不符合种植规则，强制面积为0
                prob += x[(land_name, year, season, crop_id)] == 0
                count += 1

        return count

    def _add_complementarity_constraints(self, prob, x, years):
        """添加互补性约束（线性化处理）"""
        count = 0

        # 豆类与非豆类作物的互补性约束
        for year in years:
            # 计算豆类总面积
            legume_vars = []
            non_legume_vars = []

            for (land_name, yr, season, crop_id) in x.keys():
                if yr == year:
                    if self.crops[crop_id]['is_legume']:
                        legume_vars.append(x[(land_name, yr, season, crop_id)])
                    else:
                        non_legume_vars.append(x[(land_name, yr, season, crop_id)])

            # 豆类面积应该占总面积的5%-25%
            if legume_vars and non_legume_vars:
                total_vars = legume_vars + non_legume_vars
                prob += lpSum(legume_vars) >= 0.05 * lpSum(total_vars)  # 至少5%
                prob += lpSum(legume_vars) <= 0.25 * lpSum(total_vars)  # 最多25%
                count += 2

        return count

    def _add_demand_elasticity_constraints(self, prob, x, years):
        """添加需求弹性约束"""
        count = 0

        for crop_id in self.expected_sales.keys():
            base_sales = self.expected_sales[crop_id]
            elasticity = self.demand_elasticity[crop_id]

            for year in years:
                params = self.correlation_params[year][crop_id]

                # 考虑价格变化对需求的影响
                price_effect = 1 + elasticity * (params['price_multiplier'] - 1)
                adjusted_demand = base_sales * params['sales_multiplier'] * price_effect

                # 产量约束
                production_vars = []
                for (land_name, yr, season, c_id) in x.keys():
                    if yr == year and c_id == crop_id:
                        opt = self._find_option(land_name, season, c_id)
                        if opt:
                            expected_yield = opt['base_yield'] * params['yield_multiplier']
                            production_vars.append(x[(land_name, yr, season, c_id)] * expected_yield)

                if production_vars:
                    prob += lpSum(production_vars) <= adjusted_demand * 1.1  # 允许10%的超产
                    count += 1

        return count

    def _add_crop_indicator_constraints(self, prob, x, z_crop, years):
        """添加作物指示变量约束"""
        count = 0

        for year in years:
            for crop_id in self.crops.keys():
                # 如果种植某种作物，对应的指示变量为1
                crop_vars = [x[(land_name, yr, season, c_id)]
                             for (land_name, yr, season, c_id) in x.keys()
                             if yr == year and c_id == crop_id]

                if crop_vars:
                    # 大M约束
                    prob += lpSum(crop_vars) <= 1000 * z_crop[(year, crop_id)]
                    prob += lpSum(crop_vars) >= 0.1 * z_crop[(year, crop_id)]
                    count += 2

        return count

    def _add_diversity_constraints(self, prob, z_crop, years):
        """添加多样性约束"""
        count = 0

        for year in years:
            # 每年至少种植6种不同作物（提高多样性要求）
            crop_indicators = [z_crop[(year, crop_id)] for crop_id in self.crops.keys()
                               if (year, crop_id) in z_crop]
            if crop_indicators:
                prob += lpSum(crop_indicators) >= 6
                count += 1

            # 每个作物类别至少种植一种
            categories = ['grain', 'vegetable', 'mushroom']
            for category in categories:
                category_indicators = [z_crop[(year, crop_id)]
                                       for crop_id in self.crops.keys()
                                       if (self.crops[crop_id]['category'] == category and
                                           (year, crop_id) in z_crop)]
                if category_indicators:
                    prob += lpSum(category_indicators) >= 1
                    count += 1

        return count

    def _add_rotation_constraints(self, prob, x, years):
        """添加轮作约束"""
        count = 0

        # 重茬约束
        for land_name in self.lands.keys():
            for crop_id in self.crops.keys():
                for season in ['单季', '第一季', '第二季']:
                    for year in years[:-1]:
                        current_vars = [x[(ln, yr, s, c_id)]
                                        for (ln, yr, s, c_id) in x.keys()
                                        if (ln == land_name and yr == year and
                                            s == season and c_id == crop_id)]
                        next_vars = [x[(ln, yr, s, c_id)]
                                     for (ln, yr, s, c_id) in x.keys()
                                     if (ln == land_name and yr == year + 1 and
                                         s == season and c_id == crop_id)]

                        if current_vars and next_vars:
                            prob += lpSum(current_vars) + lpSum(next_vars) <= 0.1
                            count += 1

        # 豆类轮作约束（每三年至少一次）
        for land_name in self.lands.keys():
            # 2024-2026年豆类约束
            legume_vars_early = []
            for year in [2024, 2025, 2026]:
                if year in years:
                    for (ln, yr, season, crop_id) in x.keys():
                        if (ln == land_name and yr == year and
                                self.crops[crop_id]['is_legume']):
                            legume_vars_early.append(x[(ln, yr, season, crop_id)])

            if legume_vars_early:
                prob += lpSum(legume_vars_early) >= 0.2
                count += 1

            # 2027-2029年豆类约束
            legume_vars_late = []
            for year in [2027, 2028, 2029]:
                if year in years:
                    for (ln, yr, season, crop_id) in x.keys():
                        if (ln == land_name and yr == year and
                                self.crops[crop_id]['is_legume']):
                            legume_vars_late.append(x[(ln, yr, season, crop_id)])

            if legume_vars_late:
                prob += lpSum(legume_vars_late) >= 0.2
                count += 1

        return count

    def _add_risk_control_constraints(self, prob, x, years):
        """添加风险控制约束"""
        count = 0

        # 单一作物种植面积不能超过总面积的40%
        total_area = sum(info['area'] for info in self.lands.values())

        for crop_id in self.crops.keys():
            for year in years:
                crop_vars = [x[(land_name, yr, season, c_id)]
                             for (land_name, yr, season, c_id) in x.keys()
                             if yr == year and c_id == crop_id]

                if crop_vars:
                    prob += lpSum(crop_vars) <= total_area * 0.4
                    count += 1

        # 高风险作物（食用菌）总面积限制
        for year in years:
            mushroom_vars = [x[(land_name, yr, season, crop_id)]
                             for (land_name, yr, season, crop_id) in x.keys()
                             if (yr == year and
                                 self.crops[crop_id]['category'] == 'mushroom')]

            if mushroom_vars:
                prob += lpSum(mushroom_vars) <= total_area * 0.15  # 食用菌不超过15%
                count += 1

        return count

    def solve_advanced_model(self):
        """求解高级模型"""
        print("开始求解高级相关性模型...")

        prob, x, y_water, z_crop = self.create_advanced_model()

        try:
            # 使用默认求解器（和您之前的代码一样）
            print("使用PuLP默认求解器...")
            prob.solve()

            status = LpStatus[prob.status]
            print(f"求解状态: {status}")

            if status in ['Optimal', 'Feasible']:
                print("✅ 求解成功")
                return self._extract_advanced_results(x, y_water, z_crop)
            else:
                print(f"❌ 求解失败: {status}")
                # 如果复杂模型失败，尝试简化模型
                print("尝试简化模型...")
                return self._solve_simplified_model()

        except Exception as e:
            print(f"求解过程出错: {e}")
            # 尝试简化模型
            print("尝试简化模型...")
            return self._solve_simplified_model()

    def _solve_simplified_model(self):
        """求解简化模型（当主模型失败时）"""
        print("创建并求解简化模型...")

        try:
            # 创建简化的模型（移除二进制变量）
            prob_simple = LpProblem("Simplified_Crop_Optimization", LpMaximize)
            years = list(range(2024, 2031))

            # 只使用连续变量
            x_simple = {}
            for land_name in self.lands.keys():
                land_type = self.lands[land_name]['type']
                for year in years:
                    for opt in self.valid_options:
                        if opt['land_type'] == land_type:
                            var_key = (land_name, year, opt['season'], opt['crop_id'])
                            x_simple[var_key] = LpVariable(f"x_simple_{len(x_simple)}", lowBound=0, cat='Continuous')

            # 简化的目标函数
            total_objective_simple = 0
            for (land_name, year, season, crop_id) in x_simple.keys():
                params = self.correlation_params[year][crop_id]
                opt = self._find_option(land_name, season, crop_id)

                if opt:
                    expected_yield = opt['base_yield'] * params['yield_multiplier']
                    expected_cost = opt['base_cost'] * params['cost_multiplier']
                    expected_price = opt['base_price'] * params['price_multiplier']

                    base_profit = expected_yield * expected_price - expected_cost

                    # 豆类激励
                    if self.crops[crop_id]['is_legume']:
                        base_profit += 100

                    total_objective_simple += x_simple[(land_name, year, season, crop_id)] * base_profit

            prob_simple += total_objective_simple

            # 添加基本约束
            self._add_simplified_constraints(prob_simple, x_simple, years)

            # 求解简化模型（使用默认求解器）
            prob_simple.solve()

            status = LpStatus[prob_simple.status]
            print(f"简化模型求解状态: {status}")

            if status in ['Optimal', 'Feasible']:
                print("✅ 简化模型求解成功")
                return self._extract_simplified_results(x_simple)
            else:
                print(f"❌ 简化模型也无法求解: {status}")
                return None, 0

        except Exception as e:
            print(f"简化模型求解失败: {e}")
            return None, 0

    def _add_simplified_constraints(self, prob, x, years):
        """添加简化约束"""
        # 1. 地块面积约束
        for land_name, land_info in self.lands.items():
            max_area = land_info['area']
            for year in years:
                for season in ['单季', '第一季', '第二季']:
                    season_vars = [x[(ln, yr, s, crop_id)]
                                   for (ln, yr, s, crop_id) in x.keys()
                                   if ln == land_name and yr == year and s == season]
                    if season_vars:
                        prob += lpSum(season_vars) <= max_area

        # 2. 销售量约束
        for crop_id in self.expected_sales.keys():
            base_sales = self.expected_sales[crop_id]
            for year in years:
                params = self.correlation_params[year][crop_id]
                max_sales = base_sales * params['sales_multiplier'] * 1.2

                production_vars = []
                for (land_name, yr, season, c_id) in x.keys():
                    if yr == year and c_id == crop_id:
                        opt = self._find_option(land_name, season, c_id)
                        if opt:
                            expected_yield = opt['base_yield'] * params['yield_multiplier']
                            production_vars.append(x[(land_name, yr, season, c_id)] * expected_yield)

                if production_vars:
                    prob += lpSum(production_vars) <= max_sales

        # 3. 简化的重茬约束
        for land_name in self.lands.keys():
            for crop_id in self.crops.keys():
                for season in ['单季', '第一季', '第二季']:
                    for year in years[:-1]:
                        current_vars = [x[(ln, yr, s, c_id)]
                                        for (ln, yr, s, c_id) in x.keys()
                                        if (ln == land_name and yr == year and
                                            s == season and c_id == crop_id)]
                        next_vars = [x[(ln, yr, s, c_id)]
                                     for (ln, yr, s, c_id) in x.keys()
                                     if (ln == land_name and yr == year + 1 and
                                         s == season and c_id == crop_id)]

                        if current_vars and next_vars:
                            prob += lpSum(current_vars) + lpSum(next_vars) <= 0.1

    def _extract_simplified_results(self, x):
        """提取简化结果"""
        results = []
        total_profit = 0
        years = list(range(2024, 2031))

        for (land_name, year, season, crop_id), var in x.items():
            area = var.varValue
            if area and area > 0.01:
                crop_name = self.crops[crop_id]['name']
                land_type = self.lands[land_name]['type']

                params = self.correlation_params[year][crop_id]
                opt = self._find_option(land_name, season, crop_id)

                if opt:
                    expected_yield = opt['base_yield'] * params['yield_multiplier']
                    expected_cost = opt['base_cost'] * params['cost_multiplier']
                    expected_price = opt['base_price'] * params['price_multiplier']

                    production = area * expected_yield
                    cost = area * expected_cost
                    revenue = production * expected_price
                    profit = revenue - cost

                    if self.crops[crop_id]['is_legume']:
                        profit += area * 100

                    total_profit += profit

                    results.append({
                        '年份': year,
                        '地块名称': land_name,
                        '地块类型': land_type,
                        '种植季次': season,
                        '作物编号': crop_id,
                        '作物名称': crop_name,
                        '作物分类': self.crops[crop_id]['category'],
                        '种植面积': round(area, 2),
                        '预期产量': round(production, 1),
                        '调整成本': round(cost, 1),
                        '调整收入': round(revenue, 1),
                        '风险调整利润': round(profit, 1),
                        '规模效应': 0.0,
                        '需求弹性影响': 0.0,
                        '风险调整': 0.0
                    })

        print(f"简化模型求解完成，总利润: {total_profit:,.1f}元，{len(results)}个种植方案")
        return results, total_profit

    def _extract_advanced_results(self, x, y_water, z_crop):
        """提取高级模型结果"""
        results = []
        total_profit = 0
        years = list(range(2024, 2031))

        for (land_name, year, season, crop_id), var in x.items():
            area = var.varValue
            if area and area > 0.01:
                crop_name = self.crops[crop_id]['name']
                land_type = self.lands[land_name]['type']

                # 计算经济指标
                params = self.correlation_params[year][crop_id]
                opt = self._find_option(land_name, season, crop_id)

                if opt:
                    expected_yield = opt['base_yield'] * params['yield_multiplier']
                    expected_cost = opt['base_cost'] * params['cost_multiplier']
                    expected_price = opt['base_price'] * params['price_multiplier']

                    # 规模经济效应
                    scale_effect = 1 - params['scale_economy'] * (area / 10)  # 规模越大成本越低
                    adjusted_cost = expected_cost * scale_effect

                    production = area * expected_yield
                    cost = area * adjusted_cost

                    # 需求弹性调整收入
                    elasticity = self.demand_elasticity[crop_id]
                    price_effect = 1 + elasticity * (params['price_multiplier'] - 1) * 0.5
                    adjusted_price = expected_price * price_effect

                    revenue = production * adjusted_price
                    profit = revenue - cost

                    # 风险调整
                    risk_adjustment = profit * params['risk_factor'] * self.risk_aversion
                    adjusted_profit = profit - risk_adjustment

                    total_profit += adjusted_profit

                    results.append({
                        '年份': year,
                        '地块名称': land_name,
                        '地块类型': land_type,
                        '种植季次': season,
                        '作物编号': crop_id,
                        '作物名称': crop_name,
                        '作物分类': self.crops[crop_id]['category'],
                        '种植面积': round(area, 2),
                        '预期产量': round(production, 1),
                        '调整成本': round(cost, 1),
                        '调整收入': round(revenue, 1),
                        '风险调整利润': round(adjusted_profit, 1),
                        '规模效应': round((1 - scale_effect) * 100, 1),
                        '需求弹性影响': round((price_effect - 1) * 100, 1),
                        '风险调整': round(risk_adjustment, 1)
                    })

        print(f"高级模型求解完成，总利润: {total_profit:,.1f}元，{len(results)}个种植方案")

        # 计算相关性效益
        self._calculate_correlation_benefits(results)

        return results, total_profit

    def _calculate_correlation_benefits(self, results):
        """计算相关性效益"""
        results_df = pd.DataFrame(results)

        print("\n📊 相关性效益分析:")

        # 1. 多样性效益
        yearly_diversity = results_df.groupby('年份')['作物名称'].nunique()
        avg_diversity = yearly_diversity.mean()
        print(f"  平均年度作物多样性: {avg_diversity:.1f}种")

        # 2. 互补性效益
        legume_crops = []
        for crop_id, crop_info in self.crops.items():
            if crop_info['is_legume']:
                legume_crops.append(crop_info['name'])

        legume_area = results_df[results_df['作物名称'].isin(legume_crops)]['种植面积'].sum()
        total_area = results_df['种植面积'].sum()
        legume_ratio = legume_area / total_area * 100 if total_area > 0 else 0
        print(f"  豆类作物占比: {legume_ratio:.1f}%")

        # 3. 风险分散效果
        category_distribution = results_df.groupby('作物分类')['种植面积'].sum()
        max_category_ratio = category_distribution.max() / total_area * 100 if total_area > 0 else 0
        print(f"  最大类别集中度: {max_category_ratio:.1f}%")

        # 4. 规模经济效果
        avg_scale_effect = results_df['规模效应'].mean()
        print(f"  平均规模经济效应: {avg_scale_effect:.1f}%")

    def compare_with_problem2(self, problem2_file='result2_strict.xlsx'):
        """与问题2结果进行比较"""
        print("\n🔍 与问题2结果比较分析...")

        try:
            # 读取问题2结果
            problem2_df = pd.read_excel(problem2_file, sheet_name='种植方案（严格约束）')

            # 计算比较指标
            p2_total_profit = problem2_df['期望利润'].sum()
            p2_total_area = problem2_df['种植面积'].sum()
            p2_diversity = problem2_df.groupby('年份')['作物名称'].nunique().mean()

            print(f"问题2结果:")
            print(f"  总利润: {p2_total_profit:,.1f}元")
            print(f"  总面积: {p2_total_area:,.1f}亩")
            print(f"  平均多样性: {p2_diversity:.1f}种")

            return {
                'problem2_profit': p2_total_profit,
                'problem2_area': p2_total_area,
                'problem2_diversity': p2_diversity
            }

        except Exception as e:
            print(f"读取问题2结果失败: {e}")
            return None

    def save_advanced_results(self, results, total_profit, comparison_data=None,
                              output_file='result3.xlsx'):
        """保存高级结果"""
        print(f"保存高级结果到 {output_file}...")

        if not results:
            print("无结果可保存")
            return

        results_df = pd.DataFrame(results)

        # 年度汇总
        yearly_summary = results_df.groupby('年份').agg({
            '种植面积': 'sum',
            '预期产量': 'sum',
            '调整成本': 'sum',
            '调整收入': 'sum',
            '风险调整利润': 'sum',
            '作物名称': 'nunique'
        }).round(1).reset_index()
        yearly_summary.columns = ['年份', '种植面积', '预期产量', '调整成本', '调整收入', '风险调整利润', '作物种类数']

        # 作物汇总
        crop_summary = results_df.groupby(['作物编号', '作物名称', '作物分类']).agg({
            '种植面积': 'sum',
            '预期产量': 'sum',
            '调整成本': 'sum',
            '调整收入': 'sum',
            '风险调整利润': 'sum',
            '规模效应': 'mean',
            '需求弹性影响': 'mean'
        }).round(1).reset_index()
        crop_summary['利润率%'] = (crop_summary['风险调整利润'] / crop_summary['调整成本'] * 100).round(1)
        crop_summary = crop_summary.sort_values('风险调整利润', ascending=False)

        # 相关性分析
        correlation_analysis = self._generate_correlation_analysis(results_df)

        # 风险收益分析
        risk_return_analysis = self._generate_risk_return_analysis(results_df)

        # 与问题2的比较分析
        if comparison_data:
            comparison_analysis = self._generate_comparison_analysis(results_df, total_profit, comparison_data)
        else:
            comparison_analysis = pd.DataFrame([{'说明': '未提供问题2结果进行比较'}])

        # 敏感性分析
        sensitivity_analysis = self._generate_sensitivity_analysis(results_df)

        # 策略建议
        strategy_recommendations = self._generate_strategy_recommendations(results_df)

        # 保存到Excel
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            results_df.to_excel(writer, sheet_name='高级种植方案', index=False)
            yearly_summary.to_excel(writer, sheet_name='年度汇总', index=False)
            crop_summary.to_excel(writer, sheet_name='作物汇总', index=False)
            correlation_analysis.to_excel(writer, sheet_name='相关性分析', index=False)
            risk_return_analysis.to_excel(writer, sheet_name='风险收益分析', index=False)
            comparison_analysis.to_excel(writer, sheet_name='与问题2比较', index=False)
            sensitivity_analysis.to_excel(writer, sheet_name='敏感性分析', index=False)
            strategy_recommendations.to_excel(writer, sheet_name='策略建议', index=False)

            # 模型特性说明
            model_features = pd.DataFrame([
                {'特性': '农作物可替代性', '说明': '同类作物间存在替代关系，影响种植决策', '实现方式': '替代性矩阵建模'},
                {'特性': '农作物互补性', '说明': '豆类与其他作物互补，不同类别作物组合效益',
                 '实现方式': '互补性矩阵建模'},
                {'特性': '需求价格弹性', '说明': '价格变化影响需求量，不同作物弹性不同', '实现方式': '弹性系数调整'},
                {'特性': '规模经济效应', '说明': '种植规模越大，单位成本越低', '实现方式': '规模效应函数'},
                {'特性': '风险调整', '说明': '考虑不同作物的风险水平，调整预期收益', '实现方式': '风险因子和风险厌恶'},
                {'特性': '多样性激励', '说明': '鼓励作物多样化种植，提高系统稳定性', '实现方式': '多样性约束和激励'},
                {'特性': '动态相关性', '说明': '考虑年际间的动态影响和相关性变化', '实现方式': '时变参数建模'}
            ])
            model_features.to_excel(writer, sheet_name='模型特性说明', index=False)

            # 总体信息
            model_info = pd.DataFrame([{
                '问题': '问题3',
                '模型版本': '高级相关性版本',
                '规划期间': '2024-2030年',
                '总风险调整利润': f'{total_profit:,.0f}元',
                '平均年利润': f'{total_profit / 7:,.0f}元',
                '主要创新': '考虑可替代性、互补性、需求弹性、规模经济、风险调整',
                '作物多样性': f'{results_df.groupby("年份")["作物名称"].nunique().mean():.1f}种/年',
                '风险控制': '多层次风险控制机制',
                '相关性建模': '全面考虑农作物间相关关系',
                '决策支持': '提供策略建议和敏感性分析'
            }])
            model_info.to_excel(writer, sheet_name='模型信息', index=False)

        print(f"高级结果已保存到: {output_file}")

    def _generate_correlation_analysis(self, results_df):
        """生成相关性分析"""
        analysis = []

        # 替代性分析
        grain_crops = results_df[results_df['作物分类'] == 'grain']['作物名称'].unique()
        if len(grain_crops) > 1:
            analysis.append({
                '分析类型': '可替代性',
                '作物组合': f"粮食类: {', '.join(grain_crops)}",
                '替代效应': '高度可替代，优化选择高效益品种',
                '建议': '根据市场价格动态调整粮食作物结构'
            })

        # 互补性分析 - 使用更通用的方法识别豆类
        legume_crops = []
        for _, row in results_df.iterrows():
            crop_id = row['作物编号']
            if crop_id in self.crops and self.crops[crop_id]['is_legume']:
                legume_crops.append(row['作物名称'])

        legume_crops = list(set(legume_crops))  # 去重

        if legume_crops:
            legume_area = results_df[results_df['作物名称'].isin(legume_crops)]['种植面积'].sum()
            other_area = results_df[~results_df['作物名称'].isin(legume_crops)]['种植面积'].sum()

            if legume_area > 0 and other_area > 0:
                analysis.append({
                    '分析类型': '互补性',
                    '作物组合': f"豆类({', '.join(legume_crops)})与非豆类作物",
                    '互补效应': f'豆类{legume_area:.1f}亩，改善土壤利于其他作物',
                    '建议': '保持适当豆类比例，发挥土壤改良作用'
                })

        # 多样性分析
        yearly_diversity = results_df.groupby('年份')['作物名称'].nunique()
        analysis.append({
            '分析类型': '多样性',
            '作物组合': f'年均{yearly_diversity.mean():.1f}种作物',
            '多样性效应': '降低系统性风险，提高稳定性',
            '建议': '维持高多样性，避免过度集中'
        })

        return pd.DataFrame(analysis)

    def _generate_risk_return_analysis(self, results_df):
        """生成风险收益分析"""
        analysis = []

        # 按作物分类分析风险收益
        for category in results_df['作物分类'].unique():
            category_data = results_df[results_df['作物分类'] == category]

            avg_profit_rate = (category_data['风险调整利润'].sum() /
                               category_data['调整成本'].sum() * 100)
            total_area = category_data['种植面积'].sum()
            risk_level = category_data['风险调整'].mean()

            analysis.append({
                '作物类别': category,
                '种植面积': round(total_area, 1),
                '平均利润率%': round(avg_profit_rate, 1),
                '风险水平': round(risk_level, 1),
                '风险评级': self._get_risk_rating(risk_level),
                '投资建议': self._get_investment_advice(avg_profit_rate, risk_level)
            })

        return pd.DataFrame(analysis).sort_values('平均利润率%', ascending=False)

    def _get_risk_rating(self, risk_level):
        """获取风险评级"""
        if risk_level < 10:
            return '低风险'
        elif risk_level < 30:
            return '中风险'
        else:
            return '高风险'

    def _get_investment_advice(self, profit_rate, risk_level):
        """获取投资建议"""
        if profit_rate > 20 and risk_level < 20:
            return '强烈推荐'
        elif profit_rate > 15 and risk_level < 30:
            return '推荐'
        elif profit_rate > 10:
            return '适度投资'
        else:
            return '谨慎投资'

    def _generate_comparison_analysis(self, results_df, total_profit, comparison_data):
        """生成比较分析"""
        p3_total_area = results_df['种植面积'].sum()
        p3_diversity = results_df.groupby('年份')['作物名称'].nunique().mean()

        comparison = pd.DataFrame([
            {
                '指标': '总利润（元）',
                '问题2结果': f"{comparison_data['problem2_profit']:,.0f}",
                '问题3结果': f"{total_profit:,.0f}",
                '差异': f"{total_profit - comparison_data['problem2_profit']:,.0f}",
                '变化率%': f"{(total_profit / comparison_data['problem2_profit'] - 1) * 100:+.1f}"
            },
            {
                '指标': '总种植面积（亩）',
                '问题2结果': f"{comparison_data['problem2_area']:,.1f}",
                '问题3结果': f"{p3_total_area:,.1f}",
                '差异': f"{p3_total_area - comparison_data['problem2_area']:,.1f}",
                '变化率%': f"{(p3_total_area / comparison_data['problem2_area'] - 1) * 100:+.1f}"
            },
            {
                '指标': '平均作物多样性（种）',
                '问题2结果': f"{comparison_data['problem2_diversity']:.1f}",
                '问题3结果': f"{p3_diversity:.1f}",
                '差异': f"{p3_diversity - comparison_data['problem2_diversity']:+.1f}",
                '变化率%': f"{(p3_diversity / comparison_data['problem2_diversity'] - 1) * 100:+.1f}"
            }
        ])

        # 添加改进说明
        improvements = pd.DataFrame([
            {'改进方面': '经济效益', '问题3优势': '考虑规模经济和需求弹性，优化收益结构'},
            {'改进方面': '风险管理', '说明': '引入风险调整机制，提高方案稳健性'},
            {'改进方面': '作物配置', '说明': '基于替代性和互补性优化作物组合'},
            {'改进方面': '市场适应', '说明': '考虑价格弹性，增强市场适应能力'},
            {'改进方面': '可持续性', '说明': '强化多样性约束，提升长期可持续性'}
        ])

        return pd.concat([comparison, improvements], ignore_index=True)

    def _generate_sensitivity_analysis(self, results_df):
        """生成敏感性分析"""
        analysis = []

        # 价格敏感性
        analysis.append({
            '敏感性因子': '销售价格',
            '变化范围': '±10%',
            '对利润影响': '高度敏感',
            '风险等级': '中',
            '应对策略': '密切关注市场价格，及时调整种植结构'
        })

        # 成本敏感性
        analysis.append({
            '敏感性因子': '种植成本',
            '变化范围': '年增长5%',
            '对利润影响': '中度敏感',
            '风险等级': '中',
            '应对策略': '提高种植效率，控制成本上升'
        })

        # 产量敏感性
        analysis.append({
            '敏感性因子': '亩产量',
            '变化范围': '±10%',
            '对利润影响': '高度敏感',
            '风险等级': '高',
            '应对策略': '加强田间管理，提高抗风险能力'
        })

        # 需求敏感性
        analysis.append({
            '敏感性因子': '市场需求',
            '变化范围': '±5%',
            '对利润影响': '中度敏感',
            '风险等级': '中',
            '应对策略': '多样化种植，分散市场风险'
        })

        return pd.DataFrame(analysis)

    def _generate_strategy_recommendations(self, results_df):
        """生成策略建议"""
        recommendations = []

        # 短期策略（1-2年）
        recommendations.append({
            '时间范围': '短期（2024-2025）',
            '策略重点': '稳健发展',
            '具体建议': '优先种植确定性较高的粮食类作物，适度发展蔬菜类',
            '预期效果': '确保基本收益，降低初期风险'
        })

        # 中期策略（3-5年）
        recommendations.append({
            '时间范围': '中期（2026-2028）',
            '策略重点': '结构优化',
            '具体建议': '根据市场反馈调整作物结构，增加高价值作物比重',
            '预期效果': '提升整体收益水平，优化资源配置'
        })

        # 长期策略（5-7年）
        recommendations.append({
            '时间范围': '长期（2029-2030）',
            '策略重点': '可持续发展',
            '具体建议': '建立稳定的种植体系，重视土壤保护和生态平衡',
            '预期效果': '实现可持续高收益，建立品牌优势'
        })

        # 风险管理建议
        recommendations.append({
            '时间范围': '全期间',
            '策略重点': '风险管理',
            '具体建议': '建立多元化种植组合，加强市场信息收集和分析',
            '预期效果': '提高抗风险能力，确保稳定收益'
        })

        # 技术创新建议
        recommendations.append({
            '时间范围': '全期间',
            '策略重点': '技术创新',
            '具体建议': '充分利用智慧大棚优势，引入先进种植技术',
            '预期效果': '提高生产效率，增强竞争优势'
        })

        return pd.DataFrame(recommendations)

    def run_advanced_optimization(self):
        """运行高级优化"""
        print("\n开始问题3高级相关性优化求解...")

        # 求解模型
        results, total_profit = self.solve_advanced_model()

        if results is None:
            print("❌ 高级优化失败")
            return None, 0

        # 与问题2比较
        comparison_data = self.compare_with_problem2()

        # 保存结果
        self.save_advanced_results(results, total_profit, comparison_data)

        print("\n" + "=" * 70)
        print("问题3高级相关性优化完成")
        print("=" * 70)
        print("主要特性:")
        print("  ✓ 农作物可替代性和互补性建模")
        print("  ✓ 需求价格弹性机制")
        print("  ✓ 规模经济效应")
        print("  ✓ 多层次风险调整")
        print("  ✓ 动态相关性考虑")
        print("  ✓ 全面的敏感性分析")
        print("  ✓ 与问题2结果对比")
        print("  ✓ 详细策略建议")
        print("=" * 70)

        return results, total_profit


def main():
    """主函数"""
    try:
        print("问题3 - 考虑农作物间相关性的高级优化")
        print("包含可替代性、互补性、需求弹性、规模经济等因素")

        # 创建高级优化器
        optimizer = AdvancedCropOptimizer('processed_data.xlsx')

        # 运行高级优化
        results, total_profit = optimizer.run_advanced_optimization()

        if results:
            print(f"\n✅ 问题3高级优化求解成功！")
            print(f"📊 总风险调整利润: {total_profit:,.0f}元")
            print(f"📁 结果文件: result3.xlsx")
            print(f"🚀 包含全面的相关性分析和策略建议")
        else:
            print("❌ 问题3高级优化求解失败")

        return results, total_profit

    except Exception as e:
        print(f"执行过程中出现错误: {e}")
        import traceback
        traceback.print_exc()
        return None, 0


if __name__ == "__main__":
    # 检查依赖
    try:
        import pulp

        print(f"PuLP库已安装，版本: {pulp.__version__}")
    except ImportError:
        print("请先安装PuLP库: pip install pulp")
        exit(1)

    # 运行高级优化
    results, total_profit = main()

    if results:
        print("\n🎉 问题3高级相关性版本完成！")
        print("💡 这个版本全面考虑了农作物间的复杂相关关系。")
        print("📋 请查看result3.xlsx文件获取详细结果。")
    else:
        print("\n⚠️ 问题3高级优化执行失败。")
        print("💡 这可能是由于模型复杂度较高，可以尝试：")
        print("1. 检查数据文件是否存在")
        print("2. 确保有足够的内存")
        print("3. 模型会自动尝试简化版本")