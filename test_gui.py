# -*- coding: utf-8 -*-
"""GUI 冒烟测试：验证界面能正常构造。"""
import sys
import tkinter as tk
from pathlib import Path

sys.path.insert(0, r"D:\21821\products\wordtool")

pass_n = 0
fail_n = 0


def check(name, cond, detail=""):
    global pass_n, fail_n
    if cond:
        pass_n += 1
        print(f"  PASS  {name}")
    else:
        fail_n += 1
        print(f"  FAIL  {name}  {detail}")


def main():
    root = tk.Tk()
    root.withdraw()  # 不显示窗口

    try:
        from wordtool import App
        app = App()
        check("App 构造成功", True)

        # 验证页签数量
        tabs = app.nb.tabs()
        check("页签数量=6", len(tabs) == 6, f"实际 {len(tabs)}")

        # 验证页签标题
        titles = [app.nb.tab(t, "text") for t in tabs]
        expected = ["批量替换", "提取图片", "文档统计", "文档合并", "水印", "导出文本"]
        check("页签标题正确", titles == expected, f"实际 {titles}")

        # 验证变量存在
        check("find_text 变量存在", hasattr(app, 'find_text'))
        check("replace_text 变量存在", hasattr(app, 'replace_text'))
        check("case_sensitive 变量存在", hasattr(app, 'case_sensitive'))
        check("wm_mode 变量存在", hasattr(app, 'wm_mode'))
        check("wm_text 变量存在", hasattr(app, 'wm_text'))

        # 验证输出目录默认值
        out = app.out_dir.get()
        check("输出目录非空", len(out) > 0)

        app.destroy()

    except Exception as e:
        check("App 构造成功", False, str(e))
        import traceback
        traceback.print_exc()

    root.destroy()
    print(f"\n=== GUI 冒烟测试：{pass_n} PASS / {fail_n} FAIL ===")
    return 0 if fail_n == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
