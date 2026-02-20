import tempfile
import unittest
from pathlib import Path

import pandas as pd

from minwool.engine import MinwoolEngine
from minwool.io import save_results_to_excel


class TestIOModule(unittest.TestCase):
    def test_save_results_to_excel_creates_valid_workbook(self):
        engine = MinwoolEngine()
        df = engine.run()

        with tempfile.TemporaryDirectory() as tmpdir:
            out_file = Path(tmpdir) / "calc.xlsx"
            save_results_to_excel(df, engine.config, engine.fixed_costs, str(out_file))

            self.assertTrue(out_file.exists())
            self.assertGreater(out_file.stat().st_size, 0)

            with pd.ExcelFile(out_file) as xls:
                self.assertIn("Расчет_Цен", xls.sheet_names)
                self.assertIn("Входные_Данные", xls.sheet_names)

            saved_df = pd.read_excel(out_file, sheet_name="Расчет_Цен")
            self.assertEqual(len(saved_df), len(df))


if __name__ == "__main__":
    unittest.main()
