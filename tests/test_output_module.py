import unittest

import pandas as pd

from minwool.output import CalculationOutput, TkOutputAdapter


class FakeTree:
    def __init__(self):
        self.rows = []

    def get_children(self):
        return list(range(len(self.rows)))

    def delete(self, item):
        if self.rows:
            self.rows.pop(0)

    def insert(self, _parent, _index, values):
        self.rows.append(tuple(values))


class FakeText:
    def __init__(self):
        self.content = ""

    def delete(self, _start, _end):
        self.content = ""

    def insert(self, _end, text):
        self.content = text


class FakeStatus:
    def __init__(self):
        self.value = ""

    def set(self, value):
        self.value = value


class FakeGUI:
    def __init__(self):
        self.tree = FakeTree()
        self.debug_text = FakeText()
        self.status_var = FakeStatus()


class TestOutputModule(unittest.TestCase):
    def test_tk_output_adapter_renders_payload(self):
        adapter = TkOutputAdapter()
        gui = FakeGUI()
        df = pd.DataFrame([{"a": 1, "b": 2}, {"a": 3, "b": 4}])
        payload = CalculationOutput(results=df, report="ok")

        adapter.render(gui, payload)
        adapter.set_status(gui, "done")

        self.assertEqual(len(gui.tree.rows), 2)
        self.assertEqual(gui.debug_text.content, "ok")
        self.assertEqual(gui.status_var.value, "done")


if __name__ == "__main__":
    unittest.main()
