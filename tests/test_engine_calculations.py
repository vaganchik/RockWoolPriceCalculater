import unittest

from minwool.engine import MinwoolEngine


class TestMinwoolEngineCalculations(unittest.TestCase):
    def setUp(self):
        self.engine = MinwoolEngine()

    def test_calc_production_cost_t_default(self):
        self.assertAlmostEqual(self.engine.calc_production_cost_t(), 60838.85, places=2)

    def test_get_calc_context_contains_expected_fields(self):
        ctx = self.engine.get_calc_context()
        for key in ["cost_t", "resin_kg_t", "resin_cost_t", "film_price_per_lm", "film_width_m"]:
            self.assertIn(key, ctx)

    def test_calc_packs_on_pallet_default_pack_height(self):
        self.assertEqual(self.engine.calc_packs_on_pallet(600), 16)

    def test_optimize_pack_default(self):
        self.assertEqual(self.engine.optimize_pack(50, 50), 12)

    def test_optimize_pack_respects_weight_constraint(self):
        self.engine.config["max_pack_weight_kg"] = 10
        self.assertEqual(self.engine.optimize_pack(50, 200), 1)

    def test_calc_packaging_per_pack_uses_linear_meters(self):
        pkg = self.engine.calc_packaging_per_pack(n_slabs=12, packs_per_pallet=16)
        self.assertAlmostEqual(pkg["film_cost_rub"], 36.0, places=2)
        self.assertAlmostEqual(pkg["total_pallet_cost_rub"], 1226.0, places=2)
        self.assertAlmostEqual(pkg["total_pack_cost_rub"], 76.62, places=2)

    def test_run_contains_default_density_values(self):
        df = self.engine.run()
        row = df[df["Плотность кг/м3"] == 50].iloc[0]
        self.assertAlmostEqual(row["С/С 1т без упаковки"], 60838.85, places=2)
        self.assertAlmostEqual(row["С/С упаковки руб/м3"], 177.37, places=2)
        self.assertEqual(int(row["Плит в пачке"]), 12)
        self.assertEqual(int(row["Пачек на паллете"]), 16)
        self.assertAlmostEqual(
            row["Стоимость 1 фуры руб"],
            row["Стоимость паллета руб"] * self.engine.config["pallets_per_truck"],
            delta=0.1,
        )

    def test_pallet_volume_and_m3_with_packaging_formulas(self):
        df = self.engine.run()
        for _, row in df.iterrows():
            self.assertAlmostEqual(
                row["V паллета м3"],
                row["V пачки м3"] * row["Пачек на паллете"],
                places=4,
            )
            total_pallet_cost_with_pkg = (
                (row["С/С 1т без упаковки"] * row["Плотность кг/м3"] / 1000) * row["V паллета м3"]
            ) + row["Упаковка паллета руб"]
            self.assertAlmostEqual(
                row["С/С 1м3 с упаковкой"],
                total_pallet_cost_with_pkg / row["V паллета м3"],
                places=2,
            )

    def test_detailed_report_formula_is_aligned(self):
        report = self.engine.get_detailed_report()
        self.assertIn("(33500 / 0.97) + 5684.21052631579 = 60838.85", report)
        self.assertNotIn("(33500 + 5684.21052631579) / 0.97", report)
        self.assertIn("стоимость задаётся за погонный метр", report)


if __name__ == "__main__":
    unittest.main()
