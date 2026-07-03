#!/usr/bin/python
# -*- coding: UTF-8 -*-

import pandas as pd
import numpy as np
from typing import Tuple, Optional


class FertilizerRecommend:
    """
    基于测土数据和目标产量，推荐氮磷钾施肥量。
    读取 Excel 中的六张表进行查表和线性插值。
    """

    def __init__(self, excel_path: str, sheet_names: Optional[dict] = None):
        """
        参数:
            excel_path: Excel 文件路径
            sheet_names: 自定义 Sheet 名称映射，例如:
                {'N_thr':'氮阈值', 'P_thr':'磷阈值', 'K_thr':'钾阈值',
                 'N_rec':'氮推荐', 'P_rec':'磷推荐', 'K_rec':'钾推荐'}
        """
        self.excel_path = excel_path
        # 默认 Sheet 名称
        default_names = {
            'N_thr': '氮阈值', 'P_thr': '磷阈值', 'K_thr': '钾阈值',
            'N_rec': '氮推荐', 'P_rec': '磷推荐', 'K_rec': '钾推荐'
        }
        if sheet_names is not None:
            default_names.update(sheet_names)
        self.sheets = default_names

        # 读取数据
        self._load_data()

    def _load_data(self):
        """从 Excel 读取所有表，并预处理"""
        # 读取阈值表
        self.df_n_thr = pd.read_excel(self.excel_path, sheet_name=self.sheets['N_thr'])
        self.df_p_thr = pd.read_excel(self.excel_path, sheet_name=self.sheets['P_thr'])
        self.df_k_thr = pd.read_excel(self.excel_path, sheet_name=self.sheets['K_thr'])

        # 读取推荐表（注意可能有重复列名，我们按位置取）
        self.df_n_rec = self._read_recommend(self.sheets['N_rec'])
        self.df_p_rec = self._read_recommend(self.sheets['P_rec'])
        self.df_k_rec = self._read_recommend(self.sheets['K_rec'])

        # 将阈值表转换为地区→级别下限字典
        self.n_thr_dict = self._build_threshold_dict(self.df_n_thr)
        self.p_thr_dict = self._build_threshold_dict(self.df_p_thr)
        self.k_thr_dict = self._build_threshold_dict(self.df_k_thr)

        # 推荐表按类型分组（产量和各级别值）
        self.n_rec_dict = self._build_recommend_dict(self.df_n_rec)
        self.p_rec_dict = self._build_recommend_dict(self.df_p_rec)
        self.k_rec_dict = self._build_recommend_dict(self.df_k_rec)

    def _read_recommend(self, sheet_name: str) -> pd.DataFrame:
        """
        读取推荐表，处理可能的重复列名。
        我们要求第一列为目标产量，最后一列为类型，中间7列为级别7~1。
        """
        df = pd.read_excel(self.excel_path, sheet_name=sheet_name, header=0)
        # 若列名重复，pandas会自动添加.1,.2，我们重命名
        # 获取列名列表
        cols = df.columns.tolist()
        # 找到目标产量列（包含"产量"或第一个列）
        yield_col = None
        for c in cols:
            if '产量' in c or '目标' in c:
                yield_col = c
                break
        if yield_col is None:
            yield_col = cols[0]  # 默认第一列

        # 找到类型列（包含"类型"）
        type_col = None
        for c in cols:
            if '类型' in c:
                type_col = c
                break
        if type_col is None:
            type_col = cols[-1]  # 默认最后一列

        # 级别列：从yield_col之后到type_col之前，应该是7个级别列
        # 但为了保险，我们根据列名包含"级别"来筛选
        level_cols = [c for c in cols if '级别' in c and c != type_col]
        # 如果没找到，就按位置取中间的7列
        if len(level_cols) != 7:
            # 找到yield_col和type_col的位置
            try:
                yield_idx = cols.index(yield_col)
                type_idx = cols.index(type_col)
                # 两者之间的所有列（不包括两端）
                level_cols = cols[yield_idx + 1:type_idx]
                if len(level_cols) != 7:
                    # 如果还不是7列，可能顺序不对，就按数字提取
                    level_cols = [c for c in cols if any(str(i) in c for i in range(1, 8))]
            except:
                # 默认取第2到第8列
                level_cols = cols[1:8]

        # 现在构建新的DataFrame，只保留需要的列
        # 注意：如果列名重复，我们使用原始列名
        # 但我们会重命名标准列
        new_cols = ['产量', '级别7', '级别6', '级别5', '级别4', '级别3', '级别2', '级别1', '类型']
        # 提取数据
        data = {}
        data['产量'] = df[yield_col]
        for i, col in enumerate(level_cols[:7]):  # 只取前7个
            data[f'级别{7 - i}'] = df[col]  # 注意顺序：第一个级别列可能是级别7或级别1，需检查
        # 但实际顺序未知，我们根据数值判断？简单起见，我们假定列的顺序就是级别7到级别1（从低到高）
        # 因为用户数据中级别7在前，级别1在后，所以直接映射
        data['类型'] = df[type_col]

        return pd.DataFrame(data)

    def _build_threshold_dict(self, df: pd.DataFrame) -> dict:
        """将阈值DataFrame转换为地区→级别下限字典"""
        # 假设第一列为地区，其余7列对应级别7~1
        thr_dict = {}
        # 提取列名
        cols = df.columns.tolist()
        region_col = cols[0]
        level_cols = [c for c in cols if '级别' in c]  # 可能顺序不一定
        if len(level_cols) != 7:
            # 按位置取后7列
            level_cols = cols[1:9]

        for idx, row in df.iterrows():
            region = str(row[region_col]).strip()
            if region == 'nan' or not region:
                continue
            lower_bounds = []
            for col in level_cols:
                val = row[col]
                lower = self._parse_threshold(val)
                lower_bounds.append(lower)
            # 将级别号与下限对应：级别7对应第一个，级别1对应最后一个
            # 但可能顺序是级别7到级别1，我们确保顺序
            # 假设顺序就是级别7到级别1（由低到高）
            thr_dict[region] = lower_bounds  # 索引0->级别7, 6->级别1
        return thr_dict

    def _parse_threshold(self, val) -> float:
        """
        解析阈值单元格，返回该级别的下限值。
        支持格式：'<50', '50~60', '≥60', '>60', '-', 纯数字
        对于 '<50' 返回 -inf, '50~60' 返回 50, '≥60' 返回 60, '-' 返回 None
        """
        if pd.isna(val):
            return None
        s = str(val).strip()
        if s == '-' or s == '':
            return None
        if s.startswith('<'):
            return -np.inf
        if s.startswith('≥') or s.startswith('>'):
            num = s[1:].strip()
            try:
                return float(num)
            except:
                return None
        if '~' in s:
            parts = s.split('~')
            try:
                return float(parts[0].strip())
            except:
                return None
        # 纯数字
        try:
            return float(s)
        except:
            return None

    def _build_recommend_dict(self, df: pd.DataFrame) -> dict:
        """构建推荐字典，按类型分组，返回 (产量列表, 各级别值列表)"""
        # 假设df有列: 产量, 级别7,级别6,...级别1, 类型
        rec_dict = {}
        for crop_type in df['类型'].unique():
            if pd.isna(crop_type):
                continue
            sub = df[df['类型'] == crop_type].copy()
            # 按产量排序
            sub = sub.sort_values('产量')
            yields = sub['产量'].values.tolist()
            level_values = {}
            for lv in ['级别7', '级别6', '级别5', '级别4', '级别3', '级别2', '级别1']:
                level_values[lv] = sub[lv].values.tolist()
            rec_dict[crop_type] = (yields, level_values)
        return rec_dict

    def _get_level(self, region: str, soil_value: float, thr_dict: dict) -> int:
        """
        根据地区、测土值和阈值字典，返回丰缺级别(1~7)。
        级别1最高，级别7最低。
        """
        if region not in thr_dict:
            raise ValueError(f"地区 '{region}' 未在阈值表中找到")
        lower_bounds = thr_dict[region]  # 索引0->级别7, 6->级别1
        # 从级别1（最高）向级别7（最低）查找
        # 测土值 >= 该级别下限 且 该级别存在
        # 但要注意，级别7的下限可能是 -inf，所以肯定满足
        # 我们找到第一个满足 soil_value >= lower_bound 的级别（从高到低）
        # 因为级别1下限最高，级别7下限最低
        for idx in range(6, -1, -1):  # 从级别1到级别7
            lb = lower_bounds[idx]
            if lb is None:
                continue
            if soil_value >= lb:
                return 7 - idx  # idx=6 -> 1, idx=0 -> 7
        # 如果都不满足（通常不会），返回最低级别7
        return 7

    def _interpolate(self, x, xp, fp):
        """对给定点x，在离散点(xp, fp)上进行线性插值（不进行外推）"""
        if x <= xp[0]:
            return fp[0]
        if x >= xp[-1]:
            return fp[-1]
        # 找到区间
        for i in range(len(xp) - 1):
            if xp[i] <= x <= xp[i + 1]:
                slope = (fp[i + 1] - fp[i]) / (xp[i + 1] - xp[i])
                return fp[i] + slope * (x - xp[i])
        return fp[-1]  # fallback

    def recommend(self, region: str, soil_n: float, soil_p: float, soil_k: float,
                  target_yield: float, crop_type: str = '籽实') -> Tuple[float, float, float]:
        """
        主函数：返回 (N_recom, P2O5_recom, K2O_recom) 单位 kg/hm²
        """
        # 1. 确定级别
        n_level = self._get_level(region, soil_n, self.n_thr_dict)
        p_level = self._get_level(region, soil_p, self.p_thr_dict)
        k_level = self._get_level(region, soil_k, self.k_thr_dict)

        # 2. 查推荐表并插值
        n_rec = self._get_recommendation(self.n_rec_dict, crop_type, target_yield, n_level)
        p_rec = self._get_recommendation(self.p_rec_dict, crop_type, target_yield, p_level)
        k_rec = self._get_recommendation(self.k_rec_dict, crop_type, target_yield, k_level)

        return n_rec, p_rec, k_rec

    def _get_recommendation(self, rec_dict: dict, crop_type: str, target_yield: float, level: int) -> float:
        """从推荐字典中插值获取特定级别的推荐量"""
        if crop_type not in rec_dict:
            raise ValueError(f"类型 '{crop_type}' 未在推荐表中找到，可用类型: {list(rec_dict.keys())}")
        yields, level_vals = rec_dict[crop_type]
        level_key = f'级别{level}'
        if level_key not in level_vals:
            raise ValueError(f"级别 {level} 不存在")
        fp = level_vals[level_key]
        # 过滤掉可能为None的值（如果有）
        # 但我们的推荐表里都有值
        return self._interpolate(target_yield, yields, fp)


# ====================== 使用示例 ======================
if __name__ == "__main__":
    # 请替换为您的Excel文件路径
    excel_file = "玉米施肥推荐.xlsx"

    # 初始化
    fr = FertilizerRecommend(excel_file)

    # 输入参数
    region = "吉林中部"
    soil_n = 72
    soil_p = 15
    soil_k = 120
    target_yield = 30
    crop_type = "青贮"  # 或 "青贮"

    # 获取推荐
    n, p, k = fr.recommend(region, soil_n, soil_p, soil_k, target_yield, crop_type)

    print(f"地区: {region}")
    print(f"测土氮: {soil_n} mg/kg, 测土磷: {soil_p} mg/kg, 测土钾: {soil_k} mg/kg")
    print(f"目标产量: {target_yield} t/hm², 类型: {crop_type}")
    print(f"推荐施肥量: N = {n:.2f} kg/hm², P₂O₅ = {p:.2f} kg/hm², K₂O = {k:.2f} kg/hm²")