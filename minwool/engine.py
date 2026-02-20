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

    def calc_packaging_per_pack(self, n_slabs: int, packs_per_pallet: int) -> Dict[str, float]:
        """
        Расчет затрат на упаковку одной пачки (пленка, доля поддона и чехла).
        Возвращает словарь с геометрическими параметрами и стоимостью компонентов.
        """
        c = self.config
        h_pack = n_slabs * c['slab_thickness_mm']
        
        # 1. Стоимость упаковки одной пачки: (2*высота + 2*ширина) * цена_п.м.
        perimeter_m = (2 * h_pack + 2 * c['slab_width_mm']) / 1000
        pack_cost = perimeter_m * c['film_price_per_lm']
        
        # 2. Стоимость упаковки паллета: стоимость_пачки * кол-во + худ + стрейч
        total_pallet_cost = (pack_cost * packs_per_pallet) + c['hood_price'] + c['stretch_price_pallet']
        
        # 2. Стоимость упаковки одной пачки (распределенная)
        total_pack_cost = total_pallet_cost / packs_per_pallet if packs_per_pallet > 0 else 0
        
        pack_vol_m3 = (c['slab_length_mm'] * c['slab_width_mm'] * h_pack) / 1e9
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
            
            h_pack_mm = n * self.config['slab_thickness_mm']
            pks_pal = self.calc_packs_on_pallet(h_pack_mm)
            pkg = self.calc_packaging_per_pack(n, pks_pal)
            
            pallet_vol = pkg['vol_m3'] * pks_pal
            
            # Пересчет стоимости из тонн в кубические метры
            cost_wool_m3 = cost_t * rho / 1000 # Перевод стоимости тонны в м3 через плотность
            # 3. Стоимость упаковки на м3 = Стоимость упаковки паллета / Объем паллета
            cost_pkg_m3 = pkg['total_pallet_cost_rub'] / pallet_vol if pallet_vol > 0 else 0
            total_m3 = cost_wool_m3 + cost_pkg_m3
            cost_t_with_pkg = total_m3 / (rho / 1000) # Стоимость тонны с учетом упаковки
            
            pack_weight = pkg['vol_m3'] * rho
            pallet_weight = pack_weight * pks_pal
            
            truck_weight = pallet_weight * self.config['pallets_per_truck']
            truck_vol = pallet_vol * self.config['pallets_per_truck']
            
            # Расчет реальной высоты паллета (продукта)
            layers = int(self.config['target_pallet_height_mm'] // pkg['h_pack_mm'])
            real_pallet_h = layers * pkg['h_pack_mm']
            
            # Логистические параметры (сколько упаковок/поддонов влезает в 1 тонну веса)
            packs_per_ton = 1000 / pack_weight if pack_weight > 0 else 0
            pallets_per_ton = 1000 / pallet_weight if pallet_weight > 0 else 0

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
                'V паллета м3': round(pallet_vol, 2),
                'Вес фуры кг': round(truck_weight, 2),
                'V фуры м3': round(truck_vol, 2),
                'Упаковок в 1т': round(packs_per_ton, 2),
                'Поддонов в 1т': round(pallets_per_ton, 2),
                'С/С упаковки руб/м3': round(cost_pkg_m3, 2),
                'Упаковка пачки руб': pkg['total_pack_cost_rub'],
                'Упаковка паллета руб': pkg['total_pallet_cost_rub'],
                'Цена пачки руб': round(total_m3 * pkg['vol_m3'], 2)
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
        report.append(f"4. Доля поддона: {c['pallet_price']} руб / {pks_pal} пачек = {c['pallet_price']/pks_pal:.2f} руб")
        report.append(f"5. Доля худа: {c['hood_price']} руб / {pks_pal} пачек = {c['hood_price']/pks_pal:.2f} руб")
        report.append(f"6. Доля стрейча: {c['stretch_price_pallet']} руб / {pks_pal} пачек = {c['stretch_price_pallet']/pks_pal:.2f} руб")
        report.append(f"7. Итого упаковка на 1 пачку (распред.): {pkg['total_pack_cost_rub']} руб")
        
        report.append("\n=== ЭТАП 3: ПЕРЕСЧЕТ В ОБЪЕМ (м3) ===")
        rho = 50
        cost_wool_m3 = cost_t * rho / 1000
        cost_pkg_m3 = pkg['total_pack_cost_rub'] / pkg['vol_m3']
        report.append(f"1. С/С ваты в 1 м3 (при {rho} кг/м3): {cost_t} * {rho} / 1000 = {cost_wool_m3:.2f} руб")
        report.append(f"2. С/С упаковки в 1 м3: {pkg['total_pack_cost_rub']} руб / {pkg['vol_m3']:.4f} м3 = {cost_pkg_m3:.2f} руб")
        report.append(f"3. ИТОГО за 1 м3: {cost_wool_m3:.2f} + {cost_pkg_m3:.2f} = {cost_wool_m3 + cost_pkg_m3:.2f} руб")
        
        report.append("\n* Примечание: Расчеты для других плотностей выполняются аналогично с изменением параметра плотности.")
        
        return "\n".join(report)

    def save_results(self, df, filename='minwool_python_model_v1.xlsx'):
        """Сохранение результатов в Excel."""
        save_results_to_excel(df=df, config=self.config, fixed_costs=self.fixed_costs, filename=filename)
