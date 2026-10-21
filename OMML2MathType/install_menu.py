# -*- coding: utf-8 -*-
"""在 .docx 右键菜单里注册 / 移除「Word 公式转 MathType」

用 winreg 直接写 HKCU，避免批处理里引号转义导致路径出错。

    python install_menu.py install
    python install_menu.py uninstall
    python install_menu.py check
"""
import os
import sys
import winreg

HERE = os.path.dirname(os.path.abspath(__file__))
BAT = os.path.join(HERE, "转换公式.bat")
KEY = r"Software\Classes\SystemFileAssociations\.docx\shell\OMML2MathType"
CMD = KEY + r"\command"
LABEL = "Word 公式转 MathType"


def install() -> bool:
    if not os.path.isfile(BAT):
        print("[x] 找不到 %s" % BAT)
        return False
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, KEY) as k:
        winreg.SetValue(k, "", winreg.REG_SZ, LABEL)
        icon = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"),
                            "System32", "shell32.dll") + ",70"
        winreg.SetValueEx(k, "Icon", 0, winreg.REG_SZ, icon)
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, CMD) as k:
        winreg.SetValue(k, "", winreg.REG_SZ, '"%s" "%%1"' % BAT)
    print("[v] 已注册右键菜单：%s" % LABEL)
    print("    命令：\"%s\" \"%%1\"" % BAT)
    return True


def uninstall() -> bool:
    ok = True
    for sub in (CMD, KEY):
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, sub)
        except FileNotFoundError:
            ok = False
    print("[v] 已移除右键菜单" if ok else "[-] 本来就没有安装")
    return True


def check() -> bool:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, CMD) as k:
            cmd, _ = winreg.QueryValueEx(k, "")
        print("已安装：%s" % cmd)
        return True
    except FileNotFoundError:
        print("未安装")
        return False


if __name__ == "__main__":
    action = sys.argv[1] if len(sys.argv) > 1 else "install"
    fn = {"install": install, "uninstall": uninstall, "check": check}.get(action)
    if not fn:
        sys.exit("用法: python install_menu.py [install|uninstall|check]")
    sys.exit(0 if fn() else 1)
