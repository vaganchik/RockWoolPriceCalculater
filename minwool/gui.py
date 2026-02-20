"""Графический интерфейс калькулятора себестоимости минваты."""

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Optional

from .engine import MinwoolEngine
from .output import CalculationOutput, OutputAdapter, TkOutputAdapter


class ToolTip:
    """
    Класс для создания всплывающих подсказок при наведении мыши.
    Поддерживает динамическое обновление текста (для таблиц).
    """
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text # Может быть строкой или функцией
        self.tip_window = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)
        # Для Treeview отслеживаем движение мыши для смены подсказок по колонкам
        if isinstance(widget, ttk.Treeview):
            widget.bind("<Motion>", self.update_tip)

    def get_text(self, event=None):
        if callable(self.text):
            return self.text(event)
        return self.text

    def show_tip(self, event=None):
        if self.tip_window: return
        text = self.get_text(event)
        if not text: return
        
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        
        # Позиционирование
        if event:
            x = event.x_root + 20
            y = event.y_root + 10
        else:
            x = self.widget.winfo_rootx() + 20
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 2
            
        tw.wm_geometry(f"+{x}+{y}")
        self.label = tk.Label(tw, text=text, justify=tk.LEFT,
                      background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                      font=("tahoma", "9", "normal"), padx=5, pady=2)
        self.label.pack()

    def update_tip(self, event=None):
        if self.tip_window:
            text = self.get_text(event)
            if not text:
                self.hide_tip()
            else:
                self.label.configure(text=text)
                self.tip_window.wm_geometry(f"+{event.x_root + 20}+{event.y_root + 10}")

    def hide_tip(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

class MinwoolGUI:
    """
    Главный класс графического интерфейса (Tkinter).
    Связывает пользовательский ввод с расчетным движком MinwoolEngine.
    """
    def __init__(self, root, output_adapter: Optional[OutputAdapter] = None):
        self.root = root
        self.root.title("Расчет себестоимости минваты")
        self.engine = MinwoolEngine()
        self.output_adapter = output_adapter or TkOutputAdapter()
        self.entries = {}
        self.last_results = None
        self.total_fixed_cost_var = tk.StringVar()
        self._init_constants()
        self._setup_ui()

    def _init_constants(self):
        """Инициализация констант интерфейса (подписи, группы, колонки)."""
        # Словарь для перевода технических ключей в понятные названия и подсказки
        self.labels_map = {
            'throughput_t_h': ("Производительность (т/ч)", "Количество продукции, выпускаемое линией в час (по расплаву)"),
            'yield_rate': ("Выход годного (0-1)", "Коэффициент полезного использования сырья (учитывает отходы и обрезь)"),
            'loi_percent': ("Содержание связующего (%)", "Потери при прокаливании (LOI) - целевой процент смолы в плите"),
            'resin_solid_content': ("Сухой остаток смолы (0-1)", "Доля сухого вещества в закупаемом концентрате смолы"),
            'resin_efficiency': ("Эффективность смолы (0-1)", "Коэффициент удержания связующего на волокне при напылении"),
            'resin_price_per_ton': ("Цена смолы (руб/т)", "Стоимость одной тонны жидкого связующего (концентрата)"),
            'var_stone_t': ("Камень/Сырье (руб/т)", "Стоимость основного сырья (базальт, доломит) на тонну расплава"),
            'var_melting_energy_t': ("Энергия плавки (руб/т)", "Кокс или газ, расходуемые на плавление 1 тонны сырья"),
            'var_other_t': ("Прочие перем. (руб/т)", "Прочие расходные материалы, обеспыливание, вода и т.д."),
            'slab_length_mm': ("Длина плиты (мм)", "Длина одной плиты утеплителя (стандарт 1200 мм)"),
            'slab_width_mm': ("Ширина плиты (мм)", "Ширина одной плиты утеплителя (стандарт 600 мм)"),
            'slab_thickness_mm': ("Толщина плиты (мм)", "Толщина одной плиты утеплителя"),
            'target_pack_height_mm': ("Целевая высота пачки (мм)", "Желаемая высота упаковки для оптимизации загрузки фуры"),
            'target_pallet_height_mm': ("Высота паллета (мм)", "Высота продукции на поддоне (без учета самого поддона). Используется для кратности."),
            'max_pack_weight_kg': ("Макс. вес пачки (кг)", "Ограничение веса одной упаковки для ручной разгрузки"),
            'film_price_per_lm': ("Цена пленки (руб/п.м)", "Стоимость погонного метра пленки"),
            'film_width_m': ("Ширина пленки (м)", "Технологическая ширина термоусадочной пленки (справочно)."),
            'pallet_price': ("Цена поддона (руб)", "Стоимость одного деревянного поддона"),
            'hood_price': ("Цена худа (руб)", "Стоимость упаковки stretch hood"),
            'stretch_price_pallet': ("Цена стрейч (руб)", "Стоимость стрейч-пленки на один поддон"),
            'pallet_length_mm': ("Длина паллета (мм)", "Длина основания поддона (обычно 2400 или 1200)"),
            'pallet_width_mm': ("Ширина паллета (мм)", "Ширина основания поддона (обычно 1200 или 1000)"),
            'pallets_per_truck': ("Паллет в фуре", "Количество поддонов, загружаемых в одну фуру")
        }
        
        # Группировка полей для логического разделения в интерфейсе (Grid Layout)
        self.groups = {
            "Производительность": ['throughput_t_h', 'yield_rate'],
            "Переменные затраты (тонна)": ['var_stone_t', 'var_melting_energy_t', 'var_other_t'],
            "Параметры связующего": ['loi_percent', 'resin_solid_content', 'resin_efficiency', 'resin_price_per_ton'],
            "Геометрия плиты": ['slab_length_mm', 'slab_width_mm', 'slab_thickness_mm', 'target_pack_height_mm'],
            "Упаковка и логистика": ['target_pallet_height_mm', 'pallet_length_mm', 'pallet_width_mm', 'max_pack_weight_kg', 'film_price_per_lm', 'film_width_m', 'pallet_price', 'hood_price', 'stretch_price_pallet', 'pallets_per_truck']
        }

        # Определение колонок таблицы результатов: ID -> (Заголовок, Ширина)
        self.tree_cols = {
            'rho': ("Плотность", 70),
            't_no': ("1т без упак", 90),
            't_yes': ("1т с упак", 90),
            'm3_no': ("1м3 без упак", 95),
            'm3_yes': ("1м3 с упак", 95),
            'n': ("Плит", 45),
            'h': ("Высота", 55),
            'v_pack': ("V пачки", 70),
            'cnt_pal': ("Пачек/пал", 70),
            'h_pal': ("H паллет", 70),
            'w_pack': ("Вес пачки", 80),
            'w_pal': ("Вес паллет", 80),
            'v_pal': ("V паллет", 80),
            'w_truck': ("Вес фуры", 80),
            'v_truck': ("V фуры", 80),
            'pks_t': ("Упак/1т", 70),
            'pal_t': ("Подд/1т", 70),
            'pkg_pack': ("Упак. пачки", 80),
            'pkg_pal': ("Упак. паллет", 90),
            'pkg_m3': ("Упак/м3", 80),
            'price': ("Цена пачки", 90)
        }

    def _setup_ui(self):
        """Создание основной структуры интерфейса."""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill='both', padx=5, pady=5)

        # Вкладка 1: Калькулятор
        calc_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(calc_tab, text="Калькулятор")
        self._init_calculator_tab(calc_tab)
        
        # Вкладка 2: Постоянные затраты (новая)
        fixed_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(fixed_tab, text="Постоянные затраты")
        self._init_fixed_costs_tab(fixed_tab)

        # Вкладка 2: Настройка пачек
        pack_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(pack_tab, text="Настройка пачек")
        self.setup_pack_tab(pack_tab)

        # Вкладка 3: Проверка расчетов
        debug_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(debug_tab, text="Проверка расчетов")
        self._init_debug_tab(debug_tab)

    def _init_calculator_tab(self, parent: ttk.Frame):
        """Инициализация главной вкладки калькулятора (ввод параметров и таблица результатов)."""
        container = ttk.Frame(parent)
        container.pack(expand=True, fill='both')

        # Настройка весов сетки для корректного изменения размера
        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(3, weight=1)

        # Создание полей ввода по группам
        self._create_input_groups(container)

        # Группа отображения итоговых постоянных затрат
        fc_frame = ttk.LabelFrame(container, text="Постоянные затраты (Итого)", padding="10")
        fc_frame.grid(row=2, column=1, sticky=(tk.N, tk.S, tk.E, tk.W), padx=5, pady=5)
        
        ttk.Label(fc_frame, text="Сумма (руб/час):").grid(row=0, column=0, sticky=tk.W, pady=2)
        fc_entry = ttk.Entry(fc_frame, textvariable=self.total_fixed_cost_var, state="readonly", width=15)
        fc_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2, padx=(10, 0))
        ToolTip(fc_entry, "Сумма всех статей из вкладки 'Постоянные затраты'")

        # Блок предварительного просмотра результатов
        result_frame = ttk.LabelFrame(container, text="Результаты расчета (предпросмотр)", padding="10")
        result_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)

        # Настройка весов внутри фрейма результатов
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(result_frame, columns=list(self.tree_cols.keys()), show='headings', height=8)
        
        for col_id, (name, width) in self.tree_cols.items():
            self.tree.heading(col_id, text=name)
            self.tree.column(col_id, width=width, anchor=tk.CENTER, stretch=False)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Подсказка для таблицы (разъяснение С/С)
        ToolTip(self.tree, self.get_tree_tip)
        
        # Скроллбары для таблицы
        scrollbar_y = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        
        self.tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)
        
        scrollbar_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        scrollbar_x.grid(row=1, column=0, sticky=(tk.W, tk.E))

        # Кнопки управления
        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=10)

        calc_btn = ttk.Button(btn_frame, text="Рассчитать", command=self.perform_calculation)
        calc_btn.pack(side=tk.LEFT, padx=5)

        save_btn = ttk.Button(btn_frame, text="Сохранить в Excel", command=self.save_to_excel)
        save_btn.pack(side=tk.LEFT, padx=5)

        # Статусная строка
        self.status_var = tk.StringVar(value="Готов к работе")
        status_label = ttk.Label(container, textvariable=self.status_var, foreground="blue", font=("Segoe UI", 9, "italic"))
        status_label.grid(row=5, column=0, columnspan=2)

    def _create_input_groups(self, parent: ttk.Frame):
        """Создает группы полей ввода."""
        for g_idx, (group_name, keys) in enumerate(self.groups.items()):
            frame = ttk.LabelFrame(parent, text=group_name, padding="10")
            frame.grid(row=g_idx // 2, column=g_idx % 2, sticky=(tk.N, tk.S, tk.E, tk.W), padx=5, pady=5)
            
            for i, key in enumerate(keys):
                value = self.engine.config[key]
                label_text, help_text = self.labels_map.get(key, (key, ""))
                
                lbl = ttk.Label(frame, text=f"{label_text}:")
                lbl.grid(row=i, column=0, sticky=tk.W, pady=2)
                ToolTip(lbl, help_text)
                
                entry = ttk.Entry(frame, width=15)
                entry.insert(0, str(value))
                entry.grid(row=i, column=1, sticky=(tk.W, tk.E), pady=2, padx=(10, 0))
                ToolTip(entry, help_text)
                self.entries[key] = entry

    def _init_fixed_costs_tab(self, parent: ttk.Frame):
        """Инициализация вкладки управления постоянными затратами."""
        # Layout: Left - List, Right - Controls
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill='both', expand=True)

        # Treeview
        columns = ('name', 'value')
        self.fc_tree = ttk.Treeview(main_frame, columns=columns, show='headings', height=15)
        self.fc_tree.heading('name', text='Статья затрат')
        self.fc_tree.heading('value', text='Сумма (руб/час)')
        self.fc_tree.column('name', width=200)
        self.fc_tree.column('value', width=100, anchor='e')
        self.fc_tree.pack(side=tk.LEFT, fill='both', expand=True, padx=(0, 10))

        # Controls
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(side=tk.RIGHT, fill='y')

        ttk.Label(control_frame, text="Название:").pack(anchor='w')
        self.fc_name_var = tk.StringVar()
        ttk.Entry(control_frame, textvariable=self.fc_name_var).pack(fill='x', pady=(0, 5))

        ttk.Label(control_frame, text="Сумма:").pack(anchor='w')
        self.fc_val_var = tk.StringVar()
        ttk.Entry(control_frame, textvariable=self.fc_val_var).pack(fill='x', pady=(0, 10))

        ttk.Button(control_frame, text="Добавить/Обновить", command=self.add_fixed_cost).pack(fill='x', pady=2)
        ttk.Button(control_frame, text="Удалить выбранное", command=self.delete_fixed_cost).pack(fill='x', pady=2)
        
        # Total label
        self.fc_total_var = tk.StringVar()
        ttk.Label(control_frame, textvariable=self.fc_total_var, font=('Segoe UI', 10, 'bold')).pack(pady=20)

        self.refresh_fixed_costs_ui()

    def refresh_fixed_costs_ui(self):
        """Обновляет таблицу постоянных затрат и пересчитывает итоговую сумму."""
        for item in self.fc_tree.get_children():
            self.fc_tree.delete(item)
        
        total = 0
        for name, val in self.engine.fixed_costs.items():
            self.fc_tree.insert('', tk.END, values=(name, val))
            total += val
        
        self.fc_total_var.set(f"Итого: {total:,.0f}")
        self.total_fixed_cost_var.set(f"{total:,.2f}")

    def add_fixed_cost(self):
        """Добавляет или обновляет статью постоянных затрат."""
        name = self.fc_name_var.get().strip()
        val_str = self.fc_val_var.get().replace(',', '.')
        if not name:
            messagebox.showwarning("Ошибка", "Введите название статьи затрат")
            return
        try:
            val = float(val_str)
            self.engine.fixed_costs[name] = val
            self.refresh_fixed_costs_ui()
            self.fc_name_var.set("")
            self.fc_val_var.set("")
        except ValueError:
            messagebox.showerror("Ошибка", "Сумма должна быть числом")

    def delete_fixed_cost(self):
        """Удаляет выбранные статьи затрат."""
        selected = self.fc_tree.selection()
        if not selected: return
        for item in selected:
            vals = self.fc_tree.item(item, 'values')
            name = vals[0]
            if name in self.engine.fixed_costs:
                del self.engine.fixed_costs[name]
        self.refresh_fixed_costs_ui()

    def _init_debug_tab(self, parent: ttk.Frame):
        """Инициализация вкладки отладки."""
        self.debug_text = tk.Text(parent, wrap=tk.WORD, font=("Consolas", 10), bg="#f8f9fa")
        self.debug_text.pack(side=tk.LEFT, expand=True, fill='both')
        
        debug_scroll = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.debug_text.yview)
        debug_scroll.pack(side=tk.RIGHT, fill='y')
        self.debug_text.configure(yscrollcommand=debug_scroll.set)
        
        self.debug_text.insert(tk.END, "Нажмите 'Рассчитать', чтобы увидеть детализацию формул...")

    def setup_pack_tab(self, parent: ttk.Frame):
        """Создает интерфейс для настройки количества плит в пачке."""
        self.pack_tab_frame = parent
        # Очистка перед перерисовкой (для обновления списка)
        for widget in parent.winfo_children():
            widget.destroy()

        # Панель добавления новой плотности
        control_frame = ttk.Frame(parent)
        control_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(control_frame, text="Новая плотность:").pack(side=tk.LEFT, padx=5)
        self.new_density_var = tk.StringVar()
        entry_new = ttk.Entry(control_frame, textvariable=self.new_density_var, width=10)
        entry_new.pack(side=tk.LEFT, padx=5)
        entry_new.bind('<Return>', lambda e: self.add_density()) # Добавление по Enter
        
        ttk.Button(control_frame, text="Добавить", command=self.add_density).pack(side=tk.LEFT, padx=5)

        # Область с прокруткой для списка плотностей
        container = ttk.Frame(parent)
        container.pack(fill='both', expand=True)
        
        canvas = tk.Canvas(container)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Заголовки таблицы
        ttk.Label(scrollable_frame, text="Плотность", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        ttk.Label(scrollable_frame, text="Режим расчета", font=("Segoe UI", 9, "bold")).grid(row=0, column=1, padx=10, pady=5, sticky=tk.W)
        ttk.Label(scrollable_frame, text="Кол-во плит (шт)", font=("Segoe UI", 9, "bold")).grid(row=0, column=2, padx=10, pady=5, sticky=tk.W)
        ttk.Label(scrollable_frame, text="Удалить", font=("Segoe UI", 9, "bold")).grid(row=0, column=3, padx=10, pady=5, sticky=tk.W)
        
        self.pack_vars = {}
        self.engine.densities.sort() # Сортируем для удобства
        
        for i, rho in enumerate(self.engine.densities):
            row = i + 1
            ttk.Label(scrollable_frame, text=str(rho)).grid(row=row, column=0, padx=10, pady=2)
            
            # Гарантируем наличие настроек для этой плотности
            if rho not in self.engine.pack_settings:
                self.engine.pack_settings[rho] = {'mode': 'auto', 'manual_n': 1}
            setting = self.engine.pack_settings[rho]
            
            mode_var = tk.StringVar(value=setting['mode'])
            n_var = tk.StringVar(value=str(setting['manual_n']))
            
            cb = ttk.Combobox(scrollable_frame, textvariable=mode_var, values=["auto", "manual"], state="readonly", width=10)
            cb.grid(row=row, column=1, padx=10, pady=2)
            
            entry = ttk.Entry(scrollable_frame, textvariable=n_var, width=10)
            entry.grid(row=row, column=2, padx=10, pady=2)
            
            if mode_var.get() == 'auto':
                entry.configure(state='disabled')
            
            # Кнопка удаления
            del_btn = ttk.Button(scrollable_frame, text="X", width=3, command=lambda r=rho: self.remove_density(r))
            del_btn.grid(row=row, column=3, padx=10, pady=2)
            
            self.pack_vars[rho] = {'mode': mode_var, 'n': n_var, 'entry': entry}
            
            # Привязываем событие смены режима
            cb.bind("<<ComboboxSelected>>", lambda e, r=rho: self.on_pack_mode_change(r))

    def save_pack_settings_from_ui(self):
        """Сохраняет текущие значения из UI в конфиг движка."""
        for rho, vars in self.pack_vars.items():
            try:
                n_val = int(vars['n'].get())
            except ValueError:
                n_val = 1
            self.engine.pack_settings[rho] = {
                'mode': vars['mode'].get(),
                'manual_n': n_val
            }

    def add_density(self):
        """Добавляет новую плотность в список."""
        val_str = self.new_density_var.get().replace(',', '.')
        if not val_str: return
        try:
            val = float(val_str)
            if val <= 0:
                messagebox.showwarning("Ошибка", "Плотность должна быть положительной.")
                return
            # Приводим к int, если число целое (50.0 -> 50)
            if val.is_integer():
                val = int(val)
                
            if val in self.engine.densities:
                messagebox.showwarning("Ошибка", "Такая плотность уже есть в списке.")
                return
            
            # Сохраняем текущее состояние UI перед перерисовкой
            self.save_pack_settings_from_ui()
            
            self.engine.densities.append(val)
            self.engine.pack_settings[val] = {'mode': 'auto', 'manual_n': 1}
            self.new_density_var.set("") # Очищаем поле ввода
            self.setup_pack_tab(self.pack_tab_frame) # Перерисовываем вкладку
            
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректное число.")

    def remove_density(self, rho):
        """Удаляет плотность из списка."""
        if messagebox.askyesno("Удаление", f"Удалить плотность {rho}?"):
            self.save_pack_settings_from_ui()
            if rho in self.engine.densities:
                self.engine.densities.remove(rho)
            if rho in self.engine.pack_settings:
                del self.engine.pack_settings[rho]
            self.setup_pack_tab(self.pack_tab_frame)

    def on_pack_mode_change(self, rho):
        vars = self.pack_vars[rho]
        if vars['mode'].get() == 'manual':
            vars['entry'].configure(state='normal')
        else:
            vars['entry'].configure(state='disabled')

    def get_tree_tip(self, event):
        """
        Обработчик движения мыши над таблицей.
        Определяет колонку и строку под курсором, чтобы показать контекстную формулу.
        """
        region = self.tree.identify_region(event.x, event.y)
        if region not in ("heading", "cell"):
            return None
            
        col = self.tree.identify_column(event.x)
        density = None
        
        # Если навели на ячейку с данными, пытаемся узнать плотность в этой строке
        if region == "cell":
            row_id = self.tree.identify_row(event.y)
            if row_id:
                try:
                    # Первая колонка (индекс 0) - это плотность
                    vals = self.tree.item(row_id, 'values')
                    if vals: density = float(vals[0])
                except (ValueError, IndexError):
                    pass

        try:
            col_index = int(col[1:]) - 1
            col_id = list(self.tree_cols.keys())[col_index]
            return self.get_col_formula(col_id, density)
        except (ValueError, IndexError):
            pass
        return None

    def get_col_formula(self, col_id, density=None):
        """
        Генерирует текст подсказки с формулой и подстановкой значений.
        Использует текущий контекст расчетов (self.engine.get_calc_context).
        """
        ctx = self.engine.get_calc_context()
        rho_str = str(density) if density else "Плотность"
        
        # Предварительные расчеты упаковки для подстановки чисел
        n = self.engine.optimize_pack(ctx['slab_t'], density if density else 50)
        h_pack_mm = n * ctx['slab_t']
        pks_pal = self.engine.calc_packs_on_pallet(h_pack_mm)
        pkg = self.engine.calc_packaging_per_pack(n, pks_pal)
        v_pack = pkg['vol_m3']
        v_pal = v_pack * pks_pal
        cost_pkg_m3 = pkg['total_pallet_cost_rub'] / v_pal if v_pal > 0 else 0
        
        # Расчеты, зависящие от плотности
        cost_wool_m3 = 0
        total_m3 = 0
        w_pack = 0
        w_pal = 0
        
        if density:
            cost_wool_m3 = ctx['cost_t'] * density / 1000
            total_m3 = cost_wool_m3 + cost_pkg_m3
            w_pack = v_pack * density
            w_pal = w_pack * pks_pal
        
        formulas = {
            'rho': {
                'desc': "Заданная плотность продукции.",
                'gen': "Плотность (кг/м3)",
                'sub': f"{rho_str}"
            },
            't_no': {
                'desc': "Себестоимость 1 тонны ваты без учета упаковки.\nПостоянные и переменные (сырье/энергия) делятся на выход годного.\nСмола НЕ делится — LOI% задан от веса готовой плиты.",
                'gen': "(Fix_h / Q / Yield) + (Var_t / Yield) + Resin_cost_t",
                'sub': f"({ctx['fixed_h']} / {ctx['q']} / {ctx['y']}) + ({ctx['var_t_ex']} / {ctx['y']}) + {ctx['resin_cost_t']} = {ctx['cost_t']} руб/т"
            },
            't_yes': {
                'desc': "Себестоимость 1 тонны ваты с учетом всех упаковочных материалов.",
                'gen': "С/С_1м3_с_упак / (Плотность / 1000)",
                'sub': f"{round(total_m3, 2) if density else 'ИТОГО_м3'} / ({rho_str} / 1000) = {round(total_m3 / (density/1000), 2) if density else '...'}"
            },
            'm3_no': {
                'desc': "Стоимость ваты в одном кубическом метре.",
                'gen': "С/С_1т_без_упак * Плотность / 1000",
                'sub': f"{ctx['cost_t']} * {rho_str} / 1000 = {round(cost_wool_m3, 2) if density else '...'}"
            },
            'm3_yes': {
                'desc': "Полная себестоимость одного кубического метра (вата + упаковка).",
                'gen': "С/С_1м3_без_упак + С/С_упаковки_м3",
                'sub': f"{round(cost_wool_m3, 2) if density else 'С/С_1м3_без_упак'} + {round(cost_pkg_m3, 2)} = {round(total_m3, 2) if density else '...'}"
            },
            'n': {
                'desc': "Количество плит в одной упаковке.",
                'gen': "Мин(Целевая_высота, Макс_вес) с учетом кратности паллету",
                'sub': f"Оптимизация для {rho_str} кг/м3 -> {n} шт"
            },
            'h': {
                'desc': "Фактическая высота сформированной пачки.",
                'gen': "Плит_в_пачке * Толщина_плиты",
                'sub': f"{n} * {ctx['slab_t']} = {pkg['h_pack_mm']} мм"
            },
            'v_pack': {
                'desc': "Геометрический объем одной упаковки.",
                'gen': "Длина * Ширина * Высота_пачки",
                'sub': f"{ctx['slab_length_mm']/1000} * {ctx['slab_width_mm']/1000} * {pkg['h_pack_mm']/1000} = {v_pack:.4f} м3"
            },
            'cnt_pal': {
                'desc': "Количество упаковок на одном поддоне (расчетное).",
                'gen': "(Пачек_в_слое) * (Высота_паллета // Высота_пачки)",
                'sub': f"{pks_pal} шт (для высоты пачки {pkg['h_pack_mm']} мм)"
            },
            'h_pal': {
                'desc': "Реальная высота продукции на поддоне (кратная высоте пачки).",
                'gen': "Высота_пачки * (Целевая_высота_паллета // Высота_пачки)",
                'sub': f"{pkg['h_pack_mm']} * ({ctx['target_pallet_height_mm']} // {pkg['h_pack_mm']}) = {pkg['h_pack_mm'] * int(ctx['target_pallet_height_mm'] // pkg['h_pack_mm'])} мм"
            },
            'w_pack': {
                'desc': "Вес одной упаковки готовой продукции.",
                'gen': "Объем_пачки * Плотность",
                'sub': f"{v_pack:.4f} * {rho_str} = {round(w_pack, 2) if density else '...'}"
            },
            'w_pal': {
                'desc': "Вес всей ваты на одном поддоне (без веса дерева).",
                'gen': "Вес_пачки * Пачек_на_поддоне (динамически)", 
                'sub': f"{round(w_pack, 2) if density else 'W_пачки'} * {pks_pal} = {round(w_pal, 2) if density else '...'}"
            },
            'v_pal': { 'desc': "Объем продукции на одном поддоне.", 'gen': "Объем_пачки * Пачек_на_поддоне", 'sub': f"{v_pack:.4f} * {pks_pal} = {round(v_pack * pks_pal, 2)}" },
            'w_truck': { 'desc': "Вес продукции в одной фуре.", 'gen': "Вес_поддона * Паллет_в_фуре", 'sub': f"{round(w_pal, 2) if density else 'W_поддона'} * {ctx['pals_truck']} = {round(w_pal * ctx['pals_truck'], 2) if density else '...'}" },
            'v_truck': { 'desc': "Объем продукции в одной фуре.", 'gen': "V_поддона * Паллет_в_фуре", 'sub': f"{round(v_pack * pks_pal, 2)} * {ctx['pals_truck']} = {round(v_pack * pks_pal * ctx['pals_truck'], 2)}" },
            'pks_t': { 'desc': "Упаковок из 1т.", 'gen': "1000 / Вес_пачки", 'sub': f"1000 / {round(w_pack, 2) if density else 'W_пачки'} = {round(1000 / w_pack, 2) if density and w_pack > 0 else '...'}" },
            'pal_t': { 'desc': "Поддонов из 1т.", 'gen': "1000 / Вес_поддона", 'sub': f"1000 / {round(w_pal, 2) if density else 'W_поддона'} = {round(1000 / w_pal, 2) if density and w_pal > 0 else '...'}" },
            'pkg_pack': { 'desc': "Стоимость упаковки одной пачки.", 'gen': "Стоимость_упак_паллета / Пачек_на_паллете", 'sub': f"{pkg['total_pallet_cost_rub']} / {pks_pal} = {pkg['total_pack_cost_rub']}" },
            'pkg_pal': { 'desc': "Стоимость упаковки всего паллета.", 'gen': "(Упак_пачки * Кол-во) + Худ + Стрейч", 'sub': f"({pkg['film_cost_rub']} * {pks_pal}) + {ctx['hood_price']} + {ctx['stretch_price_pallet']} = {pkg['total_pallet_cost_rub']}" },
            'pkg_m3': { 'desc': "Затраты на упаковку на 1 м3.", 'gen': "Стоимость_упак_паллета / Объем_паллета", 'sub': f"{pkg['total_pallet_cost_rub']} / {v_pal:.4f} = {round(cost_pkg_m3, 2)}" },
            'price': { 'desc': "Себестоимость одной пачки.", 'gen': "С/С_1м3_с_упак * Объем_пачки", 'sub': f"{round(total_m3, 2) if density else 'ИТОГО_м3'} * {v_pack:.4f} = {round(total_m3 * v_pack, 2) if density else '...'}" }
        }
        
        f = formulas.get(col_id)
        if not f: return ""
        return f"{f['desc']}\n\nФормула:\n{f['gen']}\n\nРасчет:\n{f['sub']}"

    def perform_calculation(self):
        """Считывает данные из полей ввода, запускает расчет и обновляет таблицу."""
        try:
            # 1. Считывание конфигурации из UI
            for key in self.engine.config:
                val = self.entries[key].get().replace(',', '.')
                # Целые числа для счетчиков
                self.engine.config[key] = float(val) if '.' in val or key not in ['pallets_per_truck'] else int(val)

            # Обновляем настройки пачек через вспомогательный метод
            self.save_pack_settings_from_ui()

            self.last_results = self.engine.run()
            self.output_adapter.render(
                self,
                CalculationOutput(
                    results=self.last_results,
                    report=self.engine.get_detailed_report(),
                ),
            )
            self.output_adapter.set_status(self, "Расчет выполнен.")
        except ValueError:
            self.output_adapter.error("Ошибка ввода", "Пожалуйста, введите корректные числовые значения.")
            self.last_results = None
        except Exception as e:
            self.output_adapter.error("Ошибка", f"Произошла ошибка при расчете: {str(e)}")
            self.last_results = None

    def save_to_excel(self):
        """Сохраняет текущие результаты расчета в файл Excel."""
        if self.last_results is None:
            self.perform_calculation()
            if self.last_results is None: return

        try:
            filename = 'minwool_python_model_v1.xlsx'
            self.engine.save_results(self.last_results, filename)
            self.output_adapter.set_status(self, f"Файл '{filename}' успешно сохранен.")
            self.output_adapter.info("Успех", f"Файл {filename} сохранен.")
        except Exception as e:
            self.output_adapter.error("Ошибка сохранения", f"Не удалось сохранить файл: {str(e)}")

def launch_app():
    """Запускает GUI-приложение."""
    try:
        root = tk.Tk()
        root.title("Загрузка калькулятора...")
        
        # Устанавливаем минимальный размер, чтобы окно не схлопывалось
        root.minsize(1100, 850)
        
        # Настройка стиля
        style = ttk.Style()
        if 'vista' in style.theme_names():
            style.theme_use('vista')
        else:
            style.theme_use('clam')

        MinwoolGUI(root)
        
        # Центрирование окна на экране
        root.update_idletasks()
        width = root.winfo_reqwidth()
        height = root.winfo_reqheight()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f'{width}x{height}+{x}+{y}')
        
        # Вывод на передний план
        root.lift()
        root.focus_force()
        
        root.mainloop()
    except Exception as e:
        print(f"Критическая ошибка при запуске: {e}")
        # Если tk успел инициализироваться, покажем ошибку в окне
        messagebox.showerror("Ошибка запуска", f"Приложение не смогло запуститься:\n{e}")


if __name__ == "__main__":
    launch_app()
