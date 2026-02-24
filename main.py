"""Project root launcher for USB-util.

This launcher resolves the installed console script from internal PyPI.
"""

from importlib.metadata import entry_points


SCRIPT_NAME = "usb-util-gui"


def _iter_console_scripts():
    eps = entry_points()
    if hasattr(eps, "select"):
        return eps.select(group="console_scripts")
    return eps.get("console_scripts", [])


def _resolve_script(name: str):
    for ep in _iter_console_scripts():
        if ep.name == name:
            return ep.load()
    raise SystemExit(
        f"Console script '{name}' が見つかりません。"
        " internal PyPI パッケージをインストールしてください。"
    )


def main() -> None:
    _resolve_script(SCRIPT_NAME)()


if __name__ == "__main__":
    main()
