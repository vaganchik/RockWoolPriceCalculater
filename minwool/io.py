from typing import Any, Dict

import pandas as pd
from xlsxwriter.utility import xl_col_to_name


def save_results_to_excel(
    df: pd.DataFrame,
    config: Dict[str, Any],
    fixed_costs: Dict[str, float],
    filename: str = "minwool_python_model_v1.xlsx",
) -> None:
    """Сохраняет результаты расчета и входные параметры в Excel."""
    with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Расчет_Цен", index=False)

        params_list = list(config.items()) + [("--- Постоянные затраты ---", "")] + list(fixed_costs.items())
        params_df = pd.DataFrame(params_list, columns=["Параметр", "Значение"])
        params_df.to_excel(writer, sheet_name="Входные_Данные", index=False)

        workbook = writer.book
        worksheet = writer.sheets["Расчет_Цен"]
        header_format = workbook.add_format({"bold": True, "bg_color": "#D7E4BC", "border": 1})

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        last_col = xl_col_to_name(max(len(df.columns) - 1, 0))
        worksheet.set_column(f"A:{last_col}", 18)
