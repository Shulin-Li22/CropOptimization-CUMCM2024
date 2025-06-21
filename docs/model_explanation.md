# 农作物种植策略优化模型详细说明

## 项目概述

本项目针对2024年全国大学生数学建模竞赛C题"农作物种植策略"，建立了一套完整的农作物种植优化体系。项目包含数据预处理模块和三个递进式的优化模型，从基础确定性优化逐步发展到考虑不确定性和作物间相关性的高级优化模型。

## 模型架构

```
数据预处理 → 问题1(基础优化) → 问题2(不确定性优化) → 问题3(相关性优化)
```

---

## 1. 数据预处理模块 (`data_preprocessing.py`)

### 1.1 功能概述
数据预处理模块负责读取、清洗和标准化原始Excel数据，为后续优化模型提供统一的数据接口。

### 1.2 主要功能

#### 数据读取与清洗
- **输入文件**：`附件1.xlsx`（地块和作物信息）、`附件2.xlsx`（2023年种植数据）
- **清洗策略**：移除空值记录，标准化数据格式
- **异常处理**：完善的错误捕获和提示机制

#### 地块数据处理
```python
def process_land_data(land_df):
    land_info = {}
    for _, row in land_df.iterrows():
        land_info[row['地块名称']] = {
            'type': row['地块类型'].strip(),
            'area': float(row['地块面积/亩'])
        }
```

#### 作物信息处理
- **作物分类**：自动识别豆类作物（含"豆类"关键字）
- **经济指标计算**：基于产量、成本、价格计算收益和利润率
- **价格区间解析**：智能解析价格范围字符串（如"2.50-4.00"）

#### 智慧大棚数据补充
根据题目要求，智慧大棚第一季数据与普通大棚相同：
```python
def supplement_smart_greenhouse_data(statistics_data):
    # 复制普通大棚第一季蔬菜数据到智慧大棚
    # 排除食用菌，只包含蔬菜类作物
```

### 1.3 输出数据结构
生成`processed_data.xlsx`，包含8个工作表：
1. **地块信息**：标准化的地块类型和面积
2. **作物信息**：作物编号、名称、类型、豆类标识
3. **作物统计数据**：包含经济指标的完整统计信息
4. **2023年种植情况**：历史种植记录
5. **种植面积汇总**：按作物汇总的种植面积
6. **预期销售量**：基于2023年数据计算的预期销售量
7. **地块类型统计**：各类型地块的数量和面积统计
8. **数据补充说明**：数据处理方法说明

---

## 2. 问题1：基础优化模型 (`Question1_modeling.py`)

### 2.1 模型概述
问题1建立了基础的线性规划模型，在假定市场参数稳定的前提下，优化2024-2030年的种植策略。

### 2.2 数学模型

#### 决策变量
- $x_{i,t,s,j}$：第$t$年在地块$i$的第$s$季种植作物$j$的面积（亩）
- $y_{i,t}$：水浇地$i$在第$t$年的选择（1=单季水稻，0=两季蔬菜）
- $z_{i,t,s,j}$：二进制变量，表示是否在地块$i$第$t$年第$s$季种植作物$j$

#### 目标函数
$$\max Z = \sum_{i,t,s,j} x_{i,t,s,j} \cdot \text{profit}_{j,s,\text{type}(i)} + \sum_{i,t,s,j} \text{legume\_bonus}_{j} \cdot x_{i,t,s,j}$$

其中：
- $\text{profit}_{j,s,\text{type}(i)}$：作物$j$在地块类型$\text{type}(i)$第$s$季的单位利润
- $\text{legume\_bonus}_{j}$：豆类作物激励（200元/亩）

#### 约束条件

**1. 地块面积约束**
$$\sum_{j} x_{i,t,s,j} \leq A_i, \quad \forall i,t,s$$

**2. 水浇地互斥选择约束**
$$\sum_{j \in \text{rice}} x_{i,t,\text{单季},j} \leq A_i \cdot y_{i,t}$$
$$\sum_{s \in \{\text{第一季,第二季}\}} \sum_{j} x_{i,t,s,j} \leq A_i \cdot (1-y_{i,t})$$

**3. 销售量约束**
- **场景1**：$\sum_{i,t,s} x_{i,t,s,j} \cdot \text{yield}_{j,s,\text{type}(i)} \leq 7 \times \text{ExpectedSale}_j$
- **场景2**：$\sum_{i,t,s} x_{i,t,s,j} \cdot \text{yield}_{j,s,\text{type}(i)} \leq 10 \times \text{ExpectedSale}_j$

**4. 重茬约束**
$$z_{i,t,s,j} + z_{i,t+1,s,j} \leq 1, \quad \forall i,t,s,j$$

**5. 豆类轮作约束**
$$\sum_{t \in [2024,2026]} \sum_{s,j \in \text{legume}} x_{i,t,s,j} \geq 0.2, \quad \forall i \text{ (if no legume in 2023)}$$

### 2.3 种植规则约束

严格按照题目要求实现地块类型与作物的匹配规则：

| 地块类型 | 季次 | 允许作物 |
|---------|------|---------|
| 平旱地/梯田/山坡地 | 单季 | 粮食类（水稻除外） |
| 水浇地 | 单季 | 水稻 |
| 水浇地 | 第一季 | 蔬菜（冬季蔬菜除外） |
| 水浇地 | 第二季 | 冬季蔬菜（大白菜、白萝卜、红萝卜） |
| 普通大棚 | 第一季 | 蔬菜（冬季蔬菜除外） |
| 普通大棚 | 第二季 | 食用菌 |
| 智慧大棚 | 第一季/第二季 | 蔬菜（冬季蔬菜除外） |

### 2.4 求解策略
- **主求解器**：使用PuLP的CBC求解器
- **备用策略**：如果主模型不可行，自动简化约束重新求解
- **时间限制**：设置1200秒求解时间限制

---

## 3. 问题2：不确定性优化模型 (`Question2_modeling.py`)

### 3.1 模型概述
问题2在问题1的基础上，考虑市场参数的不确定性，建立了鲁棒优化模型。

### 3.2 不确定性参数建模

#### 销售量变化
- **小麦、玉米**：年增长率5%-10%（模型采用7.5%）
- **其他作物**：年变化±5%（模型采用相对稳定）

#### 产量变化
- **变化范围**：±10%
- **建模方法**：采用保守估计（-5%）

#### 成本变化
- **增长率**：每年5%
- **公式**：$\text{cost}_{t} = \text{cost}_{2023} \times (1.05)^{t-2023}$

#### 价格变化
- **粮食类**：基本稳定
- **蔬菜类**：每年增长5%
- **食用菌**：下降1%-5%（羊肚菌下降5%，其他3%）

### 3.3 期望参数计算
```python
def _calculate_expected_parameters(self):
    for year in years:
        years_from_base = year - 2023
        for crop_id in self.crops.keys():
            self.expected_params[year][crop_id] = {
                'sales_multiplier': (1 + sales_growth) ** years_from_base,
                'yield_multiplier': yield_change,
                'cost_multiplier': (1.05) ** years_from_base,
                'price_multiplier': (1 + price_change) ** years_from_base
            }
```

### 3.4 增强约束
- **多样性约束**：每年至少种植5种不同作物
- **风险分散**：单一作物种植面积限制
- **豆类轮作**：严格执行三年轮作要求

---

## 4. 问题3：相关性优化模型 (`Question3_modeling.py`)

### 4.1 模型概述
问题3是最高级的优化模型，全面考虑农作物间的可替代性、互补性和市场相关性。

### 4.2 相关性建模

#### 可替代性矩阵
$$S_{ij} = \begin{cases}
1.0 & \text{if } i = j \\
0.8 & \text{if } \text{category}(i) = \text{category}(j) \text{ and category} \in \{\text{grain, vegetable}\} \\
0.6 & \text{if } \text{category}(i) = \text{category}(j) \text{ and category} = \text{mushroom} \\
0.3 & \text{if legume-grain substitution} \\
0.1 & \text{otherwise}
\end{cases}$$

#### 互补性矩阵
$$C_{ij} = \begin{cases}
0.6 & \text{if legume-nonlegume complementarity} \\
0.2 & \text{if different categories} \\
-0.1 & \text{if same category competition} \\
0.0 & \text{if } i = j
\end{cases}$$

#### 需求弹性系数
- **粮食类**：$\epsilon = -0.3$（刚性需求）
- **蔬菜类**：$\epsilon = -0.8$（弹性需求）
- **食用菌**：$\epsilon = -1.2$（高弹性需求）

### 4.3 高级目标函数
$$\max Z = \sum_{i,t,s,j} x_{i,t,s,j} \cdot \text{AdjustedProfit}_{i,t,s,j} + \sum_{t,j} z_{t,j} \cdot \text{DiversityIncentive}$$

其中调整利润包含：
```python
adjusted_profit = base_profit + scale_benefit - risk_penalty + complementarity_bonus
```

### 4.4 规模经济建模
$$\text{ScaleEffect}_{j} = 1 - \alpha_j \times \frac{\text{area}}{10}$$

其中$\alpha_j$为规模经济系数：
- 粮食类：0.15
- 蔬菜类：0.10  
- 食用菌：0.20

### 4.5 风险调整机制
$$\text{RiskAdjustedProfit} = \text{Profit} - \text{Profit} \times \beta_j \times \gamma$$

其中：
- $\beta_j$：作物$j$的风险因子
- $\gamma$：风险厌恶系数（0.3）

### 4.6 需求弹性约束
$$\text{AdjustedDemand}_{j,t} = \text{BaseSales}_j \times \text{SalesMultiplier}_{j,t} \times (1 + \epsilon_j \times (\text{PriceMultiplier}_{j,t} - 1))$$

---

## 5. 模型求解与验证

### 5.1 求解策略
1. **主求解器**：PuLP默认求解器（通常为CBC）
2. **求解时间**：问题1设置1200秒，问题2/3设置600秒
3. **容错机制**：模型不可行时自动降级到简化版本

### 5.2 约束验证系统
每个模型都包含完整的约束验证功能：

```python
def _validate_solution(self, results):
    # 1. 种植规则验证
    # 2. 面积约束验证  
    # 3. 销售量约束验证
    # 4. 重茬约束验证
    # 5. 豆类轮作验证
```

### 5.3 结果输出结构
每个问题生成详细的Excel报告，包含：
- **种植方案**：详细的年度种植计划
- **年度汇总**：按年度的经济指标汇总
- **作物汇总**：按作物类型的种植面积和效益
- **约束验证报告**：所有约束条件的执行情况
- **敏感性分析**：关键参数的敏感性测试
- **策略建议**：基于结果的决策建议

---

## 6. 模型特色与创新

### 6.1 严格的约束实现
- 100%遵循题目给出的种植规则
- 完整的约束验证体系
- 详细的违规检测和报告

### 6.2 递进式建模策略
- 从简单到复杂的模型设计
- 每个模型都可独立运行
- 便于理解和调试

### 6.3 全面的不确定性处理
- 多种不确定性源的综合考虑
- 保守与乐观策略的平衡
- 鲁棒性优化方法

### 6.4 创新的相关性建模
- 农作物间相关关系的数学建模
- 替代性和互补性的量化分析
- 市场弹性的动态调整

### 6.5 智能求解机制
- 自动模型简化策略
- 多层次容错处理
- 计算效率优化

---

## 7. 使用说明

### 7.1 环境要求
```python
pandas>=1.5.0
numpy>=1.21.0
openpyxl>=3.0.0
PuLP>=2.7.0
matplotlib>=3.5.0
seaborn>=0.11.0
```

### 7.2 运行流程
```bash
# 1. 数据预处理
python src/data_preprocessing.py

# 2. 问题1求解
python src/Question1_modeling.py

# 3. 问题2求解  
python src/Question2_modeling.py

# 4. 问题3求解
python src/Question3_modeling.py
```

### 7.3 输出文件
- `processed_data.xlsx`：预处理数据
- `result1_1_fixed.xlsx`：问题1场景1结果
- `result1_2_fixed.xlsx`：问题1场景2结果
- `result2_strict.xlsx`：问题2结果
- `result3.xlsx`：问题3结果

---

## 8. 技术亮点

### 8.1 代码质量
- 模块化设计，职责清晰
- 完善的异常处理机制
- 详细的日志和进度提示
- 标准的文档字符串

### 8.2 算法优化
- 高效的约束生成算法
- 智能的变量筛选策略
- 内存优化的数据结构

### 8.3 结果分析
- 多维度的结果验证
- 直观的图表和报表
- 实用的决策建议

### 8.4 扩展性
- 易于添加新的约束条件
- 支持更多的不确定性因素
- 可扩展的相关性模型

---

## 9. 模型局限性与改进方向

### 9.1 当前局限性
1. **计算复杂度**：问题3的二进制变量较多，可能影响求解效率
2. **参数估计**：部分相关性参数基于假设，缺乏实际数据验证
3. **动态调整**：模型为静态优化，未考虑期间内的动态调整

### 9.2 改进方向
1. **启发式算法**：引入遗传算法、模拟退火等元启发式方法
2. **机器学习**：利用历史数据训练预测模型
3. **随机规划**：引入随机规划方法处理不确定性
4. **多目标优化**：考虑经济、环境、社会多重目标

---

## 10. 总结

本项目建立了一套完整的农作物种植策略优化体系，从基础的线性规划模型逐步发展到考虑复杂相关性的高级优化模型。模型严格遵循题目要求，具有良好的实用性和可扩展性，为农业种植决策提供了科学的量化分析工具。

通过递进式的建模策略，项目不仅解决了具体的优化问题，还展示了数学建模中从简单到复杂、从确定到不确定的建模思路，具有良好的教学和参考价值。