"""Совместимый entrypoint v2 для калькулятора минваты."""

from minwool.engine import MinwoolEngine
from minwool.gui import MinwoolGUI, ToolTip, launch_app

__all__ = ["MinwoolEngine", "MinwoolGUI", "ToolTip", "launch_app"]


if __name__ == "__main__":
    launch_app()
