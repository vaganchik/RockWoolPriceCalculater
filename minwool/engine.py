"""
Калькулятор себестоимости производства минеральной ваты (минваты)
=================================================================

НАЗНАЧЕНИЕ
----------
Инструмент для расчёта производственной себестоимости плит из минеральной ваты
с учётом упаковки и логистики. Позволяет быстро оценить стоимость продукции
для разных плотностей (35–200 кг/м³) при изменении технологических параметров.

СТРУКТУРА РАСЧЁТА
-----------------

1. СЕБЕСТОИМОСТЬ ПРОИЗВОДСТВА 1 ТОННЫ (без упаковки)
   Формула:
       cost_t = (fix_h / Q / Y) + (var_t / Y) + resin_t

   Где:
       fix_h   — постоянные затраты (руб/час): ФОТ + энергия + амортизация
       Q       — производительность линии (т/час, по расплаву)
       Y       — выход годного (0–1): доля превращённого в товар расплава
       var_t   — переменные затраты (руб/т сырья): камень + энергия плавки + прочее
       resin_t — стоимость смолы (руб/т ГОТОВОЙ продукции)

   ВАЖНО: var_t делится на Y (затраты на тонну входного сырья → тонну готовой).
          resin_t НЕ делится на Y: LOI% задаётся от веса ГОТОВОЙ плиты, поэтому
          стоимость смолы уже выражена в руб/т готовой продукции.

   Расход жидкой смолы:
       resin_kg_t = 1000 × (LOI% / 100) / (solid_content × efficiency)
   Где solid_content — сухой остаток концентрата, efficiency — КПД напыления.

2. ГЕОМЕТРИЯ ПАЧКИ И ПАЛЛЕТА
   Оптимальное число плит в пачке (optimize_pack):
       a) Ограничение по целевой высоте пачки: n_h = target_h // thickness
       b) Ограничение по максимальному весу: n_w = max_kg // (V_плиты × density)
       c) n_start = min(n_h, n_w)
       d) Из диапазона [n_start..1] берётся максимальное n, при котором
          высота_паллета кратна высоте пачки (n × thickness), что обеспечивает
          ровную укладку без подрезки.

   Число пачек на паллете (calc_packs_on_pallet):
       Проверяются две ориентации плит на паллете (0° и 90°).
       Берётся лучшая: пачек_в_слое × (высота_паллета // высота_пачки).

3. СТОИМОСТЬ УПАКОВКИ (calc_packaging_per_pack)
   Плёнка оборачивает пачку по торцу (поперечное сечение).
   Плёнка закупается в погонных метрах (ширина рулона 1.4 м):
       периметр = (2 × высота_пачки + 2 × ширина_плиты) / 1000  [п.м.]
       стоимость_плёнки = периметр × цена_пленки_за_пм

   Стоимость упаковки паллета:
       стоимость_паллета = (плёнка_пачки × пачек_на_паллете) + худ + стрейч

   Распределённая стоимость на 1 пачку:
       стоимость_пачки = стоимость_паллета / пачек_на_паллете

4. ПЕРЕСЧЁТ В ОБЪЁМ И ИТОГ
   Для каждой плотности rho (кг/м³):
       cost_wool_m3  = cost_t × rho / 1000            [руб/м³ ваты]
       cost_pkg_m3   = стоимость_паллета / V_паллета  [руб/м³ упаковки]
       total_m3      = cost_wool_m3 + cost_pkg_m3     [итого руб/м³]
       cost_t_pkg    = total_m3 / (rho / 1000)        [руб/т с упаковкой]
       price_pack    = total_m3 × V_пачки             [руб за пачку]

АРХИТЕКТУРА МОДУЛЕЙ
-------------------
MinwoolEngine — расчётный движок (бизнес-логика, без UI)
    .calc_production_cost_t()   — себестоимость тонны
    .optimize_pack()            — оптимизация пачки
    .calc_packs_on_pallet()     — количество пачек на паллете
    .calc_packaging_per_pack()  — стоимость упаковки
    .run()                      — основной цикл, возвращает DataFrame
    .get_detailed_report()      — текстовый отчёт с расшифровкой формул
    .save_results()             — экспорт в Excel

MinwoolGUI  — графический интерфейс (Tkinter, вкладки)
ToolTip     — всплывающие подсказки с формулами для таблицы
"""
from typing import Dict, List, Any, Optional, Union

import pandas as pd

from .io import save_results_to_excel

class MinwoolEngine:
    """
    Класс для выполнения инженерных и экономических расчетов производства минваты.
    Содержит бизнес-логику: формулы себестоимости, расчет геометрии упаковки
    и агрегацию затрат.
    """
    def __init__(self):
        # Конфигурация по умолчанию (технические и стоимостные параметры)
        self.fixed_costs = {
            'ФОТ': 40000.0,
            'Энергия': 20000.0,
            'Амортизация': 20000.0
        }
        self.config = {
            'throughput_t_h': 4.0,           # Производительность линии
            'yield_rate': 0.97,              # Выход годного
            'loi_percent': 4.5,              # Содержание связующего (%)
            'resin_solid_content': 0.5,      # Сухой остаток смолы
            'resin_efficiency': 0.95,        # Эффективность смолы
            'resin_price_per_ton': 60000,    # Цена смолы (руб/т)
            'var_stone_t': 15000,            # Камень/Сырье (руб/т)
            'var_melting_energy_t': 15000,   # Энергия плавки (руб/т)
            'var_other_t': 3500,             # Прочие переменные (руб/т)
            'slab_length_mm': 1200,          # Длина плиты (мм)
            'slab_width_mm': 600,            # Ширина плиты (мм)
            'slab_thickness_mm': 50,         # Толщина плиты (мм)
            'target_pack_height_mm': 600,    # Целевая высота пачки (мм)
            'target_pallet_height_mm': 2400, # Высота паллета без поддона (мм)
            'pallet_length_mm': 2400,        # Длина паллета (мм)
            'pallet_width_mm': 1200,         # Ширина паллета (мм)
            'max_pack_weight_kg': 30,        # Максимальный вес пачки (кг)
            'film_price_per_lm': 15,         # Цена пленки (руб/п.м)
            'film_width_m': 1.4,             # Ширина пленки (м), справочный параметр методики
            'pallet_price': 1500,            # Цена поддона (руб)
            'hood_price': 500,               # Цена худа (руб)
            'stretch_price_pallet': 150,     # Цена стрейч-пленки (руб)
            'pallets_per_truck': 22          # Паллет в фуре
        }
        self.densities = [35, 50, 75, 100, 125, 150, 175, 200]
        self.pack_settings = {rho: {'mode': 'auto', 'manual_n': 1} for rho in self.densities}

    def calc_production_cost_t(self) -> float:
        """Расчет себестоимости производства одной тонны ваты (без упаковки)."""
        c = self.config
        # Расчет расхода жидкой смолы на 1 тонну готовой продукции с учетом потерь и сухого остатка
        resin_kg_t = 1000 * (c['loi_percent']/100) / (c['resin_solid_content'] * c['resin_efficiency'])
        # Стоимость смолы в составе тонны продукции
        resin_cost_t = (resin_kg_t / 1000) * c['resin_price_per_ton']
        # Агрегация затрат
        fixed_costs_h = sum(self.fixed_costs.values())
        var_costs_t_ex = c['var_stone_t'] + c['var_melting_energy_t'] + c['var_other_t']
        # Итоговая себестоимость тонны:
        # - Постоянные затраты и переменные (камень, энергия) делятся на yield_rate, т.к. это затраты
        #   на тонну ВХОДНОГО сырья/времени, а выход годного < 1.
        # - resin_cost_t НЕ делится на yield_rate: LOI% задаётся от веса ГОТОВОЙ плиты,
        #   поэтому стоимость смолы уже выражена в руб/т готовой продукции.
        cost_t = (fixed_costs_h / c['throughput_t_h'] / c['yield_rate']) + (var_costs_t_ex / c['yield_rate']) + resin_cost_t
        return round(cost_t, 2)

    def get_calc_context(self) -> Dict[str, Any]:
        """Возвращает контекст расчетов для подстановки в формулы подсказок."""
        c = self.config
        resin_kg_t = 1000 * (c['loi_percent']/100) / (c['resin_solid_content'] * c['resin_efficiency'])
        resin_cost_t = (resin_kg_t / 1000) * c['resin_price_per_ton']
        fixed_h = sum(self.fixed_costs.values())
        var_t_ex = c['var_stone_t'] + c['var_melting_energy_t'] + c['var_other_t']
        cost_t = self.calc_production_cost_t()
        
        # Расчет репрезентативного количества пачек (для 50 мм плиты и целевой высоты)
        n_rep = self.optimize_pack(c['slab_thickness_mm'], 50)
        pks_pal_rep = self.calc_packs_on_pallet(n_rep * c['slab_thickness_mm'])
        
        return {
            'resin_kg_t': round(resin_kg_t, 2),
            'resin_cost_t': round(resin_cost_t, 2),
            'fixed_h': fixed_h,
            'var_t_ex': var_t_ex,
            'cost_t': cost_t,
            'q': c['throughput_t_h'],
            'y': c['yield_rate'],
            'loi': c['loi_percent'],
            'resin_s': c['resin_solid_content'],
            'resin_e': c['resin_efficiency'],
            'resin_p': c['resin_price_per_ton'],
            'slab_t': c['slab_thickness_mm'],
            'target_h': c['target_pack_height_mm'],
            'pks_pal': pks_pal_rep,
            'pals_truck': c['pallets_per_truck'],
            'slab_length_mm': c['slab_length_mm'],
            'slab_width_mm': c['slab_width_mm'],
            'target_pallet_height_mm': c['target_pallet_height_mm'],
            'hood_price': c['hood_price'],
            'stretch_price_pallet': c['stretch_price_pallet'],
            'film_price_per_lm': c['film_price_per_lm'],
            'film_width_m': c['film_width_m'],
            'pallet_price': c['pallet_price']
        }

    def calc_packs_on_pallet(self, h_pack_mm: float) -> int:
        """Расчет количества пачек на паллете исходя из геометрии."""
        c = self.config
        # 1. Пачек в слое (проверка двух ориентаций)
        # Ориентация 1
        l1 = c['pallet_length_mm'] // c['slab_length_mm']
        w1 = c['pallet_width_mm'] // c['slab_width_mm']
        per_layer_1 = l1 * w1
        
        # Ориентация 2 (поворот на 90 град)
        l2 = c['pallet_length_mm'] // c['slab_width_mm']
        w2 = c['pallet_width_mm'] // c['slab_length_mm']
        per_layer_2 = l2 * w2
        
        per_layer = max(per_layer_1, per_layer_2)
        if per_layer < 1: per_layer = 1
        
        # 2. Количество слоев по высоте
        n_layers = int(c['target_pallet_height_mm'] // h_pack_mm)
        if n_layers < 1: n_layers = 1
        
        return int(per_layer * n_layers)

    def optimize_pack(self, slab_t: float, density: float = 50.0) -> int:
        """
        Определение оптимального количества плит в пачке.
        Учитывает: целевую высоту, макс. вес и кратность высоте паллета.
        """
        c = self.config
        # 1. Ограничение по высоте (целевая)
        n_height = int(c['target_pack_height_mm'] // slab_t)
        if n_height < 1: n_height = 1

        # 2. Ограничение по весу
        v_slab = (c['slab_length_mm'] * c['slab_width_mm'] * slab_t) / 1e9
        w_slab = v_slab * density
        n_weight = int(c['max_pack_weight_kg'] // w_slab) if w_slab > 0 else 999

        # Начальное приближение (минимум из ограничений)
        n_start = min(n_height, n_weight)
        if n_start < 1: n_start = 1

        # 3. Проверка кратности высоте паллета (без поддона)
        pallet_h = c.get('target_pallet_height_mm', 2400)
        best_n = n_start
        
        # Ищем максимальное n <= n_start, которое делит высоту паллета без остатка.
        # Используем допуск 0.001 мм для корректной работы с нецелыми толщинами плит.
        for n in range(n_start, 0, -1):
            pack_h = n * slab_t
            if pack_h > 0 and abs(pallet_h % pack_h) < 0.001:
                return n
        
        return best_n

    def _calc_pack_height_mm(self, n_slabs: int) -> float:
        """Высота пачки в мм."""
        return n_slabs * self.config['slab_thickness_mm']

    def _calc_pack_perimeter_m(self, h_pack_mm: float) -> float:
        """Периметр поперечного сечения пачки в погонных метрах."""
        return (2 * h_pack_mm + 2 * self.config['slab_width_mm']) / 1000

    def _calc_film_cost_per_pack_rub(self, perimeter_m: float) -> float:
        """Стоимость плёнки на одну пачку."""
        return perimeter_m * self.config['film_price_per_lm']

    def _calc_total_packaging_per_pallet_rub(self, film_cost_per_pack: float, packs_per_pallet: int) -> float:
        """Стоимость упаковки паллета: плёнка пачек + худ + стрейч."""
        c = self.config
        return (film_cost_per_pack * packs_per_pallet) + c['hood_price'] + c['stretch_price_pallet']

    def _calc_total_packaging_per_pack_rub(self, total_packaging_per_pallet: float, packs_per_pallet: int) -> float:
        """Распределенная стоимость упаковки на одну пачку."""
        return total_packaging_per_pallet / packs_per_pallet if packs_per_pallet > 0 else 0

    def _calc_pack_volume_m3(self, h_pack_mm: float) -> float:
        """Объем одной пачки в м3."""
        c = self.config
        return (c['slab_length_mm'] * c['slab_width_mm'] * h_pack_mm) / 1e9

    def _calc_pallet_volume_m3(self, pack_volume_m3: float, packs_per_pallet: int) -> float:
        """Объем продукции на паллете в м3."""
        return pack_volume_m3 * packs_per_pallet

    def _calc_wool_cost_m3(self, cost_t: float, density: float) -> float:
        """Себестоимость ваты в 1 м3 без упаковки."""
        return cost_t * density / 1000

    def _calc_packaging_cost_m3(self, packaging_pallet_cost_rub: float, pallet_volume_m3: float) -> float:
        """С/С упаковки в 1 м3."""
        return packaging_pallet_cost_rub / pallet_volume_m3 if pallet_volume_m3 > 0 else 0

    def _calc_wool_pallet_cost_rub(self, wool_cost_m3: float, pallet_volume_m3: float) -> float:
        """Стоимость ваты в одном паллете."""
        return wool_cost_m3 * pallet_volume_m3

    def _calc_total_pallet_cost_with_packaging_rub(self, wool_pallet_cost_rub: float, packaging_pallet_cost_rub: float) -> float:
        """Полная стоимость паллета с упаковкой."""
        return wool_pallet_cost_rub + packaging_pallet_cost_rub

    def _calc_total_cost_m3(self, total_pallet_cost_with_packaging_rub: float, pallet_volume_m3: float) -> float:
        """С/С 1м3 с упаковкой."""
        return total_pallet_cost_with_packaging_rub / pallet_volume_m3 if pallet_volume_m3 > 0 else 0

    def _calc_total_cost_t_with_packaging(self, total_cost_m3: float, density: float) -> float:
        """С/С 1т с упаковкой."""
        return total_cost_m3 / (density / 1000) if density > 0 else 0

    def _calc_pack_weight_kg(self, pack_volume_m3: float, density: float) -> float:
        """Вес пачки в кг."""
        return pack_volume_m3 * density

    def _calc_pallet_weight_kg(self, pack_weight_kg: float, packs_per_pallet: int) -> float:
        """Вес продукции на паллете в кг."""
        return pack_weight_kg * packs_per_pallet

    def _calc_truck_weight_kg(self, pallet_weight_kg: float) -> float:
        """Вес продукции в фуре в кг."""
        return pallet_weight_kg * self.config['pallets_per_truck']

    def _calc_truck_volume_m3(self, pallet_volume_m3: float) -> float:
        """Объем продукции в фуре в м3."""
        return pallet_volume_m3 * self.config['pallets_per_truck']

    def _calc_truck_cost_rub(self, pallet_cost_rub: float) -> float:
        """Полная стоимость продукции в фуре."""
        return pallet_cost_rub * self.config['pallets_per_truck']

    def _calc_real_pallet_height_mm(self, pack_height_mm: float) -> float:
        """Реальная высота паллета (продукт) в мм."""
        layers = int(self.config['target_pallet_height_mm'] // pack_height_mm) if pack_height_mm > 0 else 0
        return layers * pack_height_mm

    def _calc_units_per_ton(self, unit_weight_kg: float) -> float:
        """Количество единиц (пачек/паллета) в 1 тонне."""
        return 1000 / unit_weight_kg if unit_weight_kg > 0 else 0

    def _calc_pack_price_rub(self, total_cost_m3: float, pack_volume_m3: float) -> float:
        """Себестоимость одной пачки."""
        return total_cost_m3 * pack_volume_m3

    def calc_packaging_per_pack(self, n_slabs: int, packs_per_pallet: int) -> Dict[str, float]:
        """
        Расчет затрат на упаковку одной пачки (пленка, худ и стрейч).
        Возвращает словарь с геометрическими параметрами и стоимостью компонентов.
        """
        h_pack = self._calc_pack_height_mm(n_slabs)
        perimeter_m = self._calc_pack_perimeter_m(h_pack)
        pack_cost = self._calc_film_cost_per_pack_rub(perimeter_m)
        total_pallet_cost = self._calc_total_packaging_per_pallet_rub(pack_cost, packs_per_pallet)
        total_pack_cost = self._calc_total_packaging_per_pack_rub(total_pallet_cost, packs_per_pallet)
        pack_vol_m3 = self._calc_pack_volume_m3(h_pack)

        return {
            'n_slabs': n_slabs,
            'h_pack_mm': h_pack,
            'film_cost_rub': round(pack_cost, 2),
            'total_pallet_cost_rub': round(total_pallet_cost, 2),
            'total_pack_cost_rub': round(total_pack_cost, 2),
            'vol_m3': pack_vol_m3
        }

    def run(self) -> pd.DataFrame:
        """Основной цикл расчета для стандартного набора плотностей."""
        # 1. Базовая себестоимость 1 тонны (не зависит от плотности, только от рецептуры и производительности)
        cost_t = self.calc_production_cost_t()
        data = []
        for rho in self.densities:
            setting = self.pack_settings.get(rho, {'mode': 'auto', 'manual_n': 1})
            # Определение количества плит в пачке (автоматически или вручную)
            if setting['mode'] == 'manual':
                n = int(setting['manual_n'])
                if n < 1: n = 1
            else:
                n = self.optimize_pack(self.config['slab_thickness_mm'], rho)
            
            h_pack_mm = self._calc_pack_height_mm(n)
            pks_pal = self.calc_packs_on_pallet(h_pack_mm)
            pkg = self.calc_packaging_per_pack(n, pks_pal)
            
            pallet_vol = self._calc_pallet_volume_m3(pkg['vol_m3'], pks_pal)
            
            # Пересчет стоимости из тонн в кубические метры
            cost_wool_m3 = self._calc_wool_cost_m3(cost_t, rho)
            # Стоимость упаковки на м3 = Стоимость упаковки паллета / Объем паллета
            cost_pkg_m3 = self._calc_packaging_cost_m3(pkg['total_pallet_cost_rub'], pallet_vol)
            # С/С 1м3 с упаковкой = Стоимость 1 паллета с упаковкой / Объем 1 паллета
            wool_pallet_cost = self._calc_wool_pallet_cost_rub(cost_wool_m3, pallet_vol)
            total_pallet_cost_with_pkg = self._calc_total_pallet_cost_with_packaging_rub(wool_pallet_cost, pkg['total_pallet_cost_rub'])
            total_m3 = self._calc_total_cost_m3(total_pallet_cost_with_pkg, pallet_vol)
            cost_t_with_pkg = self._calc_total_cost_t_with_packaging(total_m3, rho)
            
            pack_weight = self._calc_pack_weight_kg(pkg['vol_m3'], rho)
            pallet_weight = self._calc_pallet_weight_kg(pack_weight, pks_pal)
            
            truck_weight = self._calc_truck_weight_kg(pallet_weight)
            truck_vol = self._calc_truck_volume_m3(pallet_vol)
            truck_cost = self._calc_truck_cost_rub(total_pallet_cost_with_pkg)
            
            # Расчет реальной высоты паллета (продукта)
            real_pallet_h = self._calc_real_pallet_height_mm(pkg['h_pack_mm'])
            
            # Логистические параметры (сколько упаковок/поддонов влезает в 1 тонну веса)
            packs_per_ton = self._calc_units_per_ton(pack_weight)
            pallets_per_ton = self._calc_units_per_ton(pallet_weight)

            data.append({
                'Плотность кг/м3': rho,
                'С/С 1т без упаковки': cost_t,
                'С/С 1т с упаковкой': round(cost_t_with_pkg, 2),
                'С/С 1м3 без упаковки': round(cost_wool_m3, 2),
                'С/С 1м3 с упаковкой': round(total_m3, 2),
                'Плит в пачке': n,
                'Высота пачки мм': pkg['h_pack_mm'],
                'V пачки м3': round(pkg['vol_m3'], 4),
                'Пачек на паллете': pks_pal,
                'Высота паллета мм': real_pallet_h,
                'Вес пачки кг': round(pack_weight, 2),
                'Вес поддона кг': round(pallet_weight, 2),
                'V паллета м3': round(pallet_vol, 4),
                'Вес фуры кг': round(truck_weight, 2),
                'V фуры м3': round(truck_vol, 4),
                'Упаковок в 1т': round(packs_per_ton, 2),
                'Поддонов в 1т': round(pallets_per_ton, 2),
                'Упаковка пачки руб': pkg['total_pack_cost_rub'],
                'Упаковка паллета руб': pkg['total_pallet_cost_rub'],
                'С/С упаковки руб/м3': round(cost_pkg_m3, 2),
                'Стоимость паллета руб': round(total_pallet_cost_with_pkg, 2),
                'Стоимость 1 фуры руб': round(truck_cost, 2),
                'Цена пачки руб': round(self._calc_pack_price_rub(total_m3, pkg['vol_m3']), 2)
            })
        return pd.DataFrame(data)

    def get_detailed_report(self) -> str:
        """Генерирует подробный текстовый отчет с описанием логики расчетов."""
        c = self.config
        cost_t = self.calc_production_cost_t()
        
        # Промежуточные расчеты для отчета
        resin_kg_t = 1000 * (c['loi_percent']/100) / (c['resin_solid_content'] * c['resin_efficiency'])
        resin_cost_t = (resin_kg_t / 1000) * c['resin_price_per_ton']
        fixed_h = sum(self.fixed_costs.values())
        var_t = c['var_stone_t'] + c['var_melting_energy_t'] + c['var_other_t']
        
        report = []
        report.append("=== ЭТАП 1: СЕБЕСТОИМОСТЬ ПРОИЗВОДСТВА 1 ТОННЫ (БЕЗ УПАКОВКИ) ===")
        report.append(f"1. Расход смолы: 1000 * ({c['loi_percent']}% / 100) / ({c['resin_solid_content']} * {c['resin_efficiency']}) = {resin_kg_t:.2f} кг/т")
        report.append(f"2. Стоимость смолы на 1т: {resin_kg_t:.2f} кг * {c['resin_price_per_ton']/1000:.2f} руб/кг = {resin_cost_t:.2f} руб")
        report.append(f"3. Постоянные затраты (час): Итого {fixed_h} руб/час")
        for name, val in self.fixed_costs.items():
            report.append(f"   - {name}: {val:.2f} руб")
        report.append(f"4. Переменные затраты (тонна): {c['var_stone_t']} + {c['var_melting_energy_t']} + {c['var_other_t']} = {var_t} руб/т")
        report.append(f"5. Итоговая формула С/С 1т:")
        report.append(f"   ({fixed_h} / {c['throughput_t_h']} / {c['yield_rate']}) + ({var_t} / {c['yield_rate']}) + {resin_cost_t} = {cost_t} руб/т")
        
        report.append("\n=== ЭТАП 2: ГЕОМЕТРИЯ И УПАКОВКА (на примере 50 кг/м3) ===")

        setting = self.pack_settings.get(50, {'mode': 'auto', 'manual_n': 1})
        if setting['mode'] == 'manual':
            n = int(setting['manual_n'])
            report.append(f"Режим: Ручной ввод ({n} плит)")
        else:
            n = self.optimize_pack(c['slab_thickness_mm'], 50)
            report.append(f"1. Оптимизация пачки (для 50 кг/м3):")
            report.append(f"   - Целевая высота: {c['target_pack_height_mm']} мм")
            report.append(f"   - Макс. вес: {c['max_pack_weight_kg']} кг")
            report.append(f"   - Высота паллета (продукт): {c['target_pallet_height_mm']} мм")
            report.append(f"   -> Результат: {n} плит (Высота {n*c['slab_thickness_mm']} мм)")

        # Исправлено: передаём packs_per_pallet как второй обязательный аргумент
        h_pack_mm = n * c['slab_thickness_mm']
        pks_pal = self.calc_packs_on_pallet(h_pack_mm)
        pkg = self.calc_packaging_per_pack(n, pks_pal)

        # Периметр поперечного сечения пачки (как в calc_packaging_per_pack):
        # плёнка оборачивает пачку по торцу: 2*высота_пачки + 2*ширина_плиты
        perimeter_m = (2 * h_pack_mm + 2 * c['slab_width_mm']) / 1000

        report.append(f"2. Расход плёнки на пачку (п.м.): (2*{h_pack_mm} + 2*{c['slab_width_mm']}) / 1000 = {perimeter_m:.3f} п.м.")
        report.append(f"   Принято: плёнка шириной {c['film_width_m']} м, стоимость задаётся за погонный метр.")
        report.append(f"3. Стоимость плёнки на пачку: {perimeter_m:.3f} п.м. * {c['film_price_per_lm']} руб/п.м. = {pkg['film_cost_rub']} руб")
        report.append(f"4. Доля худа: {c['hood_price']} руб / {pks_pal} пачек = {c['hood_price']/pks_pal:.2f} руб")
        report.append(f"5. Доля стрейча: {c['stretch_price_pallet']} руб / {pks_pal} пачек = {c['stretch_price_pallet']/pks_pal:.2f} руб")
        report.append(f"6. Итого упаковка на 1 пачку (распред.): {pkg['total_pack_cost_rub']} руб")
        report.append(f"7. Итого упаковка на 1 паллет: ({pkg['film_cost_rub']} * {pks_pal}) + {c['hood_price']} + {c['stretch_price_pallet']} = {pkg['total_pallet_cost_rub']} руб")
        
        report.append("\n=== ЭТАП 3: ПЕРЕСЧЕТ В ОБЪЕМ (м3) ===")
        rho = 50
        cost_wool_m3 = cost_t * rho / 1000
        pallet_vol = pkg['vol_m3'] * pks_pal
        cost_pkg_m3 = pkg['total_pallet_cost_rub'] / pallet_vol if pallet_vol > 0 else 0
        wool_pallet_cost = cost_wool_m3 * pallet_vol
        total_pallet_cost_with_pkg = wool_pallet_cost + pkg['total_pallet_cost_rub']
        total_m3 = total_pallet_cost_with_pkg / pallet_vol if pallet_vol > 0 else 0
        report.append(f"1. Объем паллета: {pkg['vol_m3']:.4f} * {pks_pal} = {pallet_vol:.4f} м3")
        report.append(f"2. Стоимость ваты в паллете: {cost_wool_m3:.2f} * {pallet_vol:.4f} = {wool_pallet_cost:.2f} руб")
        report.append(f"3. Стоимость 1 паллета с упаковкой: {wool_pallet_cost:.2f} + {pkg['total_pallet_cost_rub']} = {total_pallet_cost_with_pkg:.2f} руб")
        report.append(f"4. С/С 1м3 с упаковкой: {total_pallet_cost_with_pkg:.2f} / {pallet_vol:.4f} = {total_m3:.2f} руб")
        report.append(f"5. С/С упаковки в 1 м3: {pkg['total_pallet_cost_rub']} / {pallet_vol:.4f} = {cost_pkg_m3:.2f} руб")
        
        report.append("\n* Примечание: Расчеты для других плотностей выполняются аналогично с изменением параметра плотности.")
        
        return "\n".join(report)

    def save_results(self, df, filename='minwool_python_model_v1.xlsx'):
        """Сохранение результатов в Excel."""
        save_results_to_excel(df=df, config=self.config, fixed_costs=self.fixed_costs, filename=filename)
