#!/usr/bin/python
# -*- coding: UTF-8 -*-

import pandas as pd
import numpy as np
from typing import Tuple, Optional


class FertilizerRecommend:
    """
    基于测土数据和目标产量，推荐氮磷钾施肥量。
    读取 Excel 中的六张表进行查表和线性插值。
    增加级别间线性插值：根据测土值在同一级别内的相对位置，
    在当前级别和上一级别推荐量之间插值，实现更精细的推荐。
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

        # ---------- 新增：地区氮素指标类型映射 ----------
        self.nitrogen_indicator_map = {
            "河套灌区": ("全氮", "g/kg"),
            "安徽淮北": ("有机质", "g/kg"),
            "川中丘陵区": ("有机质", "g/kg"),
        }
        # 默认其它地区为碱解氮
        self.default_nitrogen_indicator = ("碱解氮", "mg/kg")

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
        cols = df.columns.tolist()
        # 找到目标产量列（包含"产量"或第一个列）
        yield_col = None
        for c in cols:
            if '产量' in c or '目标' in c:
                yield_col = c
                break
        if yield_col is None:
            yield_col = cols[0]

        # 找到类型列（包含"类型"）
        type_col = None
        for c in cols:
            if '类型' in c:
                type_col = c
                break
        if type_col is None:
            type_col = cols[-1]

        # 级别列
        level_cols = [c for c in cols if '级别' in c and c != type_col]
        if len(level_cols) != 7:
            try:
                yield_idx = cols.index(yield_col)
                type_idx = cols.index(type_col)
                level_cols = cols[yield_idx + 1:type_idx]
                if len(level_cols) != 7:
                    level_cols = [c for c in cols if any(str(i) in c for i in range(1, 8))]
            except:
                level_cols = cols[1:8]

        data = {}
        data['产量'] = df[yield_col]
        for i, col in enumerate(level_cols[:7]):
            data[f'级别{7 - i}'] = df[col]
        data['类型'] = df[type_col]
        return pd.DataFrame(data)

    def _build_threshold_dict(self, df: pd.DataFrame) -> dict:
        """将阈值DataFrame转换为地区→级别下限字典"""
        thr_dict = {}
        cols = df.columns.tolist()
        region_col = cols[0]
        level_cols = [c for c in cols if '级别' in c]
        if len(level_cols) != 7:
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
            thr_dict[region] = lower_bounds
        return thr_dict

    def _parse_threshold(self, val) -> float:
        """解析阈值单元格，返回该级别的下限值。"""
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
        try:
            return float(s)
        except:
            return None

    def _build_recommend_dict(self, df: pd.DataFrame) -> dict:
        """构建推荐字典，按类型分组，返回 (产量列表, 各级别值列表)"""
        rec_dict = {}
        for crop_type in df['类型'].unique():
            if pd.isna(crop_type):
                continue
            sub = df[df['类型'] == crop_type].copy()
            sub = sub.sort_values('产量')
            yields = sub['产量'].values.tolist()
            level_values = {}
            for lv in ['级别7', '级别6', '级别5', '级别4', '级别3', '级别2', '级别1']:
                level_values[lv] = sub[lv].values.tolist()
            rec_dict[crop_type] = (yields, level_values)
        return rec_dict

    def _get_level(self, region: str, soil_value: float, thr_dict: dict) -> int:
        """根据地区、测土值和阈值字典，返回丰缺级别(1~7)。"""
        if region not in thr_dict:
            raise ValueError(f"地区 '{region}' 未在阈值表中找到")
        lower_bounds = thr_dict[region]
        for idx in range(6, -1, -1):
            lb = lower_bounds[idx]
            if lb is None:
                continue
            if soil_value >= lb:
                return 7 - idx
        return 7

    def _interpolate(self, x, xp, fp):
        """一维线性插值（不进行外推）"""
        if x <= xp[0]:
            return fp[0]
        if x >= xp[-1]:
            return fp[-1]
        for i in range(len(xp) - 1):
            if xp[i] <= x <= xp[i + 1]:
                slope = (fp[i + 1] - fp[i]) / (xp[i + 1] - xp[i])
                return fp[i] + slope * (x - xp[i])
        return fp[-1]

    def _get_recommendation(self, rec_dict: dict, crop_type: str, target_yield: float, level: int) -> float:
        """从推荐字典中插值获取特定级别的推荐量"""
        if crop_type not in rec_dict:
            raise ValueError(f"类型 '{crop_type}' 未在推荐表中找到，可用类型: {list(rec_dict.keys())}")
        yields, level_vals = rec_dict[crop_type]
        level_key = f'级别{level}'
        if level_key not in level_vals:
            raise ValueError(f"级别 {level} 不存在")
        fp = level_vals[level_key]
        return self._interpolate(target_yield, yields, fp)

    def _get_interpolated_recommendation(self, rec_dict: dict, crop_type: str, target_yield: float,
                                         level: int, region: str, soil_val: float, thr_dict: dict) -> float:
        """
        根据测土值在级别内的相对位置，在当前级别和上一级别推荐量之间进行线性插值。
        若级别为1（最高）或区间无效，则直接返回当前级别推荐量。
        """
        if level == 1:
            return self._get_recommendation(rec_dict, crop_type, target_yield, 1)

        rec_lower = self._get_recommendation(rec_dict, crop_type, target_yield, level)
        rec_upper = self._get_recommendation(rec_dict, crop_type, target_yield, level - 1)

        bounds = thr_dict[region]
        idx = 7 - level
        lower = bounds[idx]
        upper = bounds[idx + 1]

        if lower is None or upper is None or not np.isfinite(lower) or not np.isfinite(upper) or upper <= lower:
            return rec_lower

        ratio = (soil_val - lower) / (upper - lower)
        ratio = max(0, min(1, ratio))
        interpolated = rec_lower * (1 - ratio) + rec_upper * ratio
        return interpolated

    def get_nitrogen_indicator(self, region: str) -> Tuple[str, str]:
        """
        根据地区返回氮素指标名称和单位。
        返回 (名称, 单位)
        """
        if region in self.nitrogen_indicator_map:
            return self.nitrogen_indicator_map[region]
        else:
            return self.default_nitrogen_indicator

    def recommend(self, region: str, soil_n: float, soil_p: float, soil_k: float,
                  target_yield: float, crop_type: str = '籽实') -> Tuple[float, float, float]:
        """
        主函数：返回 (N_recom, P2O5_recom, K2O_recom) 单位 kg/hm²
        采用级别间线性插值，无需额外调整参数。
        """
        n_level = self._get_level(region, soil_n, self.n_thr_dict)
        p_level = self._get_level(region, soil_p, self.p_thr_dict)
        k_level = self._get_level(region, soil_k, self.k_thr_dict)

        n_rec = self._get_interpolated_recommendation(
            self.n_rec_dict, crop_type, target_yield, n_level, region, soil_n, self.n_thr_dict)
        p_rec = self._get_interpolated_recommendation(
            self.p_rec_dict, crop_type, target_yield, p_level, region, soil_p, self.p_thr_dict)
        k_rec = self._get_interpolated_recommendation(
            self.k_rec_dict, crop_type, target_yield, k_level, region, soil_k, self.k_thr_dict)

        return n_rec, p_rec, k_rec


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
    crop_type = "青贮"

    # 获取推荐
    n, p, k = fr.recommend(region, soil_n, soil_p, soil_k, target_yield, crop_type)

    # ----- 获取氮素指标信息 -----
    n_indicator, n_unit = fr.get_nitrogen_indicator(region)

    print(f"地区: {region}")
    # 根据地区动态显示氮素指标名称和单位
    print(f"测土{n_indicator}: {soil_n} {n_unit}, 测土速效磷: {soil_p} mg/kg, 测土有效钾: {soil_k} mg/kg")
    print(f"目标产量: {target_yield} t/hm², 类型: {crop_type}")
    print(f"推荐施肥量: N = {n:.2f} kg/hm², P₂O₅ = {p:.2f} kg/hm², K₂O = {k:.2f} kg/hm²")