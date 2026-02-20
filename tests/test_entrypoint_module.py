import unittest

import minwool_engine
from minwool.engine import MinwoolEngine
from minwool.gui import MinwoolGUI, ToolTip


class TestEntrypointModule(unittest.TestCase):
    def test_compat_exports(self):
        self.assertIs(minwool_engine.MinwoolEngine, MinwoolEngine)
        self.assertIs(minwool_engine.MinwoolGUI, MinwoolGUI)
        self.assertIs(minwool_engine.ToolTip, ToolTip)
        self.assertTrue(callable(minwool_engine.launch_app))


if __name__ == "__main__":
    unittest.main()
