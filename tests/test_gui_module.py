import unittest
from unittest.mock import Mock

import tkinter as tk

from minwool.gui import MinwoolGUI


class DummyOutputAdapter:
    def __init__(self):
        self.render_called = False
        self.status_messages = []
        self.info_messages = []
        self.error_messages = []

    def render(self, _gui, _payload):
        self.render_called = True

    def set_status(self, _gui, status):
        self.status_messages.append(status)

    def info(self, title, message):
        self.info_messages.append((title, message))

    def error(self, title, message):
        self.error_messages.append((title, message))


class TestGUIModule(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
            self.root.withdraw()
        except tk.TclError as exc:
            self.skipTest(f"Tkinter недоступен в окружении: {exc}")
        self.adapter = DummyOutputAdapter()
        self.app = MinwoolGUI(self.root, output_adapter=self.adapter)

    def tearDown(self):
        if hasattr(self, "root"):
            self.root.destroy()

    def test_perform_calculation_uses_output_adapter(self):
        self.app.perform_calculation()
        self.assertTrue(self.adapter.render_called)
        self.assertIn("Расчет выполнен.", self.adapter.status_messages)

    def test_save_to_excel_uses_adapter_notifications(self):
        self.app.last_results = self.app.engine.run()
        self.app.engine.save_results = Mock()

        self.app.save_to_excel()

        self.app.engine.save_results.assert_called_once()
        self.assertTrue(any("успешно сохранен" in msg for msg in self.adapter.status_messages))
        self.assertEqual(len(self.adapter.error_messages), 0)


if __name__ == "__main__":
    unittest.main()
