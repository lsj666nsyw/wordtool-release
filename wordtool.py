# -*- coding: utf-8 -*-
"""
Word 批量处理助手 —— GUI 入口

引擎层在 wordtool_core.py，本文件只负责界面与调度。
"""
import os
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from wordtool_core import (
    APP_NAME, APP_VERSION, is_doc, human_size, unique_path,
    batch_replace, extract_images, doc_stats, merge_docs,
    add_text_watermark, add_picture_watermark, export_text,
)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.geometry("960x760")
        self.minsize(880, 680)
        self.configure(bg="#f5f6f8")

        self.files = []
        self.out_dir = tk.StringVar(value=str(Path.home() / "Desktop" / "处理结果"))
        self.running = False

        self._build()

    def _build(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TLabel", background="#f5f6f8", font=("Microsoft YaHei UI", 9))
        style.configure("TButton", font=("Microsoft YaHei UI", 9))
        style.configure("TCheckbutton", background="#f5f6f8", font=("Microsoft YaHei UI", 9))
        style.configure("TRadiobutton", background="#f5f6f8", font=("Microsoft YaHei UI", 9))
        style.configure("TNotebook.Tab", font=("Microsoft YaHei UI", 9), padding=(14, 6))
        style.configure("Accent.TButton", font=("Microsoft YaHei UI", 10, "bold"))

        # 顶部文件区
        top = ttk.Frame(self, padding=(12, 10, 12, 6))
        top.pack(fill="x")

        r1 = ttk.Frame(top)
        r1.pack(fill="x")
        ttk.Button(r1, text="添加文档…", command=self.add_files).pack(side="left")
        ttk.Button(r1, text="添加文件夹…", command=self.add_folder).pack(side="left", padx=(6, 0))
        ttk.Button(r1, text="清空", command=self.clear_files).pack(side="left", padx=(6, 0))
        self.count_label = ttk.Label(r1, text="尚未选择文档")
        self.count_label.pack(side="left", padx=(12, 0))

        r2 = ttk.Frame(top)
        r2.pack(fill="x", pady=(8, 0))
        ttk.Label(r2, text="输出到：").pack(side="left")
        ttk.Entry(r2, textvariable=self.out_dir).pack(side="left", fill="x", expand=True, padx=(4, 6))
        ttk.Button(r2, text="选择…", command=self.choose_out).pack(side="left")

        # 功能页
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=(6, 6))
        self.nb = nb

        self._tab_replace(nb)
        self._tab_images(nb)
        self._tab_stats(nb)
        self._tab_merge(nb)
        self._tab_watermark(nb)
        self._tab_text(nb)

        # 底部
        bottom = ttk.Frame(self, padding=(12, 6, 12, 10))
        bottom.pack(fill="x")
        br = ttk.Frame(bottom)
        br.pack(fill="x")
        self.run_btn = ttk.Button(br, text="开始处理", command=self.start, style="Accent.TButton")
        self.run_btn.pack(side="left", ipadx=18, ipady=4)
        ttk.Button(br, text="打开输出文件夹", command=self.open_out).pack(side="right")

        self.progress = ttk.Progressbar(bottom, mode="indeterminate")
        self.progress.pack(fill="x", pady=(8, 4))
        self.status = ttk.Label(bottom, text="就绪")
        self.status.pack(anchor="w")

        self.log = tk.Text(bottom, height=6, font=("Consolas", 9), bg="#fff",
                           relief="flat", borderwidth=1, wrap="none")
        self.log.pack(fill="x", pady=(6, 0))
        self.log.configure(state="disabled")

    # ---------- 页签 ----------
    def _tab_replace(self, nb):
        f = ttk.Frame(nb, padding=12)
        nb.add(f, text="批量替换")

        ttk.Label(f, text="在所有段落、表格、页眉页脚中查找并替换文字。",
                  foreground="#666").pack(anchor="w", pady=(0, 10))

        r1 = ttk.Frame(f)
        r1.pack(anchor="w", pady=3)
        ttk.Label(r1, text="查找：").pack(side="left")
        self.find_text = tk.StringVar(value="")
        ttk.Entry(r1, textvariable=self.find_text, width=30).pack(side="left", padx=(2, 8))

        r2 = ttk.Frame(f)
        r2.pack(anchor="w", pady=3)
        ttk.Label(r2, text="替换为：").pack(side="left")
        self.replace_text = tk.StringVar(value="")
        ttk.Entry(r2, textvariable=self.replace_text, width=30).pack(side="left", padx=(2, 0))

        r3 = ttk.Frame(f)
        r3.pack(anchor="w", pady=3)
        self.case_sensitive = tk.BooleanVar(value=True)
        ttk.Checkbutton(r3, text="区分大小写", variable=self.case_sensitive).pack(side="left")

        ttk.Label(f, text="提示：留空「查找」则不替换。等长替换（如 张三→王五）计数正常。",
                  foreground="#888").pack(anchor="w", pady=(10, 0))

    def _tab_images(self, nb):
        f = ttk.Frame(nb, padding=12)
        nb.add(f, text="提取图片")

        ttk.Label(f, text="从 DOCX 文档中提取所有内嵌图片，保存为独立图片文件。",
                  foreground="#666").pack(anchor="w", pady=(0, 10))

        ttk.Label(f, text="输出格式：保持原格式（png / jpg / gif / bmp 等）",
                  foreground="#888").pack(anchor="w", pady=(10, 0))

    def _tab_stats(self, nb):
        f = ttk.Frame(nb, padding=12)
        nb.add(f, text="文档统计")

        ttk.Label(f, text="统计每个文档的段落数、表格数、图片数、字符数等。",
                  foreground="#666").pack(anchor="w", pady=(0, 10))

        ttk.Label(f, text="统计结果写入输出文件夹的「统计结果.txt」",
                  foreground="#888").pack(anchor="w", pady=(10, 0))

    def _tab_merge(self, nb):
        f = ttk.Frame(nb, padding=12)
        nb.add(f, text="文档合并")

        ttk.Label(f, text="把多个 DOCX 合并成一个文件。文件列表顺序即合并顺序。",
                  foreground="#666").pack(anchor="w", pady=(0, 10))

        self.merge_page_break = tk.BooleanVar(value=True)
        ttk.Checkbutton(f, text="在每个文档之间插入分页符",
                        variable=self.merge_page_break).pack(anchor="w", pady=3)

        ttk.Label(f, text="提示：表格以纯文本方式合并（保留内容，不保留样式）。",
                  foreground="#888").pack(anchor="w", pady=(10, 0))

    def _tab_watermark(self, nb):
        f = ttk.Frame(nb, padding=12)
        nb.add(f, text="水印")

        ttk.Label(f, text="给文档每页加水印（通过页眉图片实现）。",
                  foreground="#666").pack(anchor="w", pady=(0, 10))

        self.wm_mode = tk.StringVar(value="text")
        ttk.Radiobutton(f, text="文字水印", variable=self.wm_mode,
                        value="text").pack(anchor="w", pady=2)
        ttk.Radiobutton(f, text="图片水印", variable=self.wm_mode,
                        value="picture").pack(anchor="w", pady=2)

        r0 = ttk.Frame(f)
        r0.pack(anchor="w", pady=3)
        ttk.Label(r0, text="水印文字：").pack(side="left")
        self.wm_text = tk.StringVar(value="")
        ttk.Entry(r0, textvariable=self.wm_text, width=30).pack(side="left", padx=(2, 0))

        r1 = ttk.Frame(f)
        r1.pack(anchor="w", pady=3)
        ttk.Label(r1, text="图片路径：").pack(side="left")
        self.wm_img = tk.StringVar(value="")
        ttk.Entry(r1, textvariable=self.wm_img, width=30).pack(side="left", padx=(2, 6))
        ttk.Button(r1, text="选择图片…", command=self.choose_wm_img).pack(side="left")

        r2 = ttk.Frame(f)
        r2.pack(anchor="w", pady=3)
        ttk.Label(r2, text="透明度：").pack(side="left")
        self.wm_opacity = tk.StringVar(value="15")
        ttk.Entry(r2, textvariable=self.wm_opacity, width=6).pack(side="left", padx=(2, 8))
        ttk.Label(r2, text="字号：").pack(side="left")
        self.wm_fontsize = tk.StringVar(value="40")
        ttk.Entry(r2, textvariable=self.wm_fontsize, width=6).pack(side="left", padx=(2, 8))
        ttk.Label(r2, text="旋转角度：").pack(side="left")
        self.wm_rotation = tk.StringVar(value="45")
        ttk.Entry(r2, textvariable=self.wm_rotation, width=6).pack(side="left", padx=(2, 0))

        ttk.Label(f, text="提示：文字水印 15% 透明度 + 45° 旋转是常见样式。",
                  foreground="#888").pack(anchor="w", pady=(10, 0))

    def _tab_text(self, nb):
        f = ttk.Frame(nb, padding=12)
        nb.add(f, text="导出文本")

        ttk.Label(f, text="把 DOCX 的正文与表格内容导出为纯文本 .txt。",
                  foreground="#666").pack(anchor="w", pady=(0, 10))

        ttk.Label(f, text="每个文档导出一个同名 .txt 文件。",
                  foreground="#888").pack(anchor="w", pady=(10, 0))

    # ---------- 文件 ----------
    def add_files(self):
        fs = filedialog.askopenfilenames(
            title="选择文档",
            filetypes=[("Word 文档", "*.docx *.docm"), ("所有文件", "*.*")])
        for x in fs or ():
            p = Path(x)
            if p not in self.files:
                self.files.append(p)
        self._refresh()

    def add_folder(self):
        d = filedialog.askdirectory(title="选择文件夹")
        if not d:
            return
        n = 0
        for p in Path(d).rglob("*"):
            if p.is_file() and is_doc(p):
                if p not in self.files:
                    self.files.append(p)
                    n += 1
        if n == 0:
            messagebox.showinfo(APP_NAME, "这个文件夹里没有找到 DOCX 文档。")
        self._refresh()

    def clear_files(self):
        self.files = []
        self._refresh()

    def _refresh(self):
        n = len(self.files)
        if n == 0:
            self.count_label.configure(text="尚未选择文档")
        else:
            total = sum(p.stat().st_size for p in self.files if p.exists())
            self.count_label.configure(
                text=f"已选择 {n} 个文档，共 {human_size(total)}")

    def choose_out(self):
        d = filedialog.askdirectory(title="选择输出文件夹")
        if d:
            self.out_dir.set(d)

    def choose_wm_img(self):
        p = filedialog.askopenfilename(
            title="选择水印图片",
            filetypes=[("图片", "*.png *.jpg *.jpeg *.bmp *.gif *.webp")])
        if p:
            self.wm_img.set(p)

    def open_out(self):
        d = self.out_dir.get()
        if os.path.isdir(d):
            os.startfile(d)
        else:
            messagebox.showinfo(APP_NAME, "输出文件夹还不存在，先处理一次吧。")

    # ---------- 执行 ----------
    def logline(self, s):
        self.log.configure(state="normal")
        self.log.insert("end", s + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _ui(self, fn):
        try:
            self.after(0, fn)
        except Exception:
            pass

    def start(self):
        if self.running:
            return
        if not self.files:
            messagebox.showwarning(APP_NAME, "请先添加文档。")
            return
        out_dir = self.out_dir.get().strip()
        if not out_dir:
            messagebox.showwarning(APP_NAME, "请选择输出文件夹。")
            return
        try:
            Path(out_dir).mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror(APP_NAME, f"无法创建输出文件夹：\n{e}")
            return

        tab = self.nb.index(self.nb.select())
        try:
            ctx = self._make_task(tab, out_dir)
            if ctx is None:
                return
        except ValueError as e:
            messagebox.showwarning(APP_NAME, str(e))
            return

        self.running = True
        self.run_btn.configure(state="disabled")
        self.progress.start(12)
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.status.configure(text="处理中…")

        threading.Thread(target=self._run, args=(ctx,), daemon=True).start()

    def _make_task(self, tab, out_dir):
        if tab == 0:  # 批量替换
            find = self.find_text.get()
            if not find:
                raise ValueError("请填写「查找」内容。")
            return ("replace", dict(files=self.files, out_dir=out_dir,
                                    find_text=find,
                                    replace_text=self.replace_text.get(),
                                    case_sensitive=self.case_sensitive.get()))

        if tab == 1:  # 提取图片
            return ("images", dict(files=self.files, out_dir=out_dir))

        if tab == 2:  # 文档统计
            return ("stats", dict(files=self.files, out_dir=out_dir))

        if tab == 3:  # 文档合并
            if len(self.files) < 2:
                raise ValueError("合并至少需要 2 个文档。")
            return ("merge", dict(files=self.files,
                                  out=unique_path(Path(out_dir) / "合并结果.docx"),
                                  add_page_break=self.merge_page_break.get()))

        if tab == 4:  # 水印
            mode = self.wm_mode.get()
            if mode == "text":
                text = self.wm_text.get().strip()
                if not text:
                    raise ValueError("请填写水印文字，或选择图片水印。")
                try:
                    fs = int(self.wm_fontsize.get() or 40)
                    rot = int(self.wm_rotation.get() or 45)
                except Exception:
                    raise ValueError("字号和旋转角度必须是数字")
                return ("watermark_text", dict(
                    files=self.files, out_dir=out_dir, text=text,
                    opacity=int(self.wm_opacity.get()), font_size=fs, rotation=rot))
            else:
                img = self.wm_img.get().strip()
                if not img or not os.path.isfile(img):
                    raise ValueError("请选择水印图片。")
                return ("watermark_pic", dict(
                    files=self.files, out_dir=out_dir, img_path=img,
                    opacity=int(self.wm_opacity.get())))

        # tab 5: 导出文本
        return ("text", dict(files=self.files, out_dir=out_dir))

    def _run(self, ctx):
        kind, a = ctx
        try:
            if kind == "replace":
                for f, n, out_p, err in batch_replace(
                        a["files"], a["out_dir"], a["find_text"],
                        a["replace_text"], case_sensitive=a["case_sensitive"]):
                    if err:
                        self._ui(lambda: self.logline(f"  ✗ {Path(f).name}：{err}"))
                    else:
                        self._ui(lambda n=n, f=f: self.logline(
                            f"  {Path(f).name}  替换 {n} 处 → {out_p}"))
                self._ui(lambda: self.logline(""))
                self._ui(lambda: self.logline("批量替换完成"))

            elif kind == "images":
                for f in a["files"]:
                    try:
                        imgs = extract_images(f, a["out_dir"])
                        self._ui(lambda imgs=imgs, f=f: self.logline(
                            f"  {Path(f).name}  提取 {len(imgs)} 张图片"))
                    except Exception as e:
                        self._ui(lambda f=f, e=e: self.logline(f"  ✗ {Path(f).name}：{e}"))
                self._ui(lambda: self.logline(""))
                self._ui(lambda: self.logline("图片提取完成"))

            elif kind == "stats":
                lines = []
                for f in a["files"]:
                    try:
                        s = doc_stats(f)
                        lines.append(f"=== {Path(f).name} ===")
                        for k, v in s.items():
                            lines.append(f"  {k}: {v}")
                        self._ui(lambda f=f, s=s: self.logline(
                            f"  {Path(f).name}  段落{s['段落数']} 表格{s['表格数']} "
                            f"图片{s['图片数']} 字符{s['字符数']}"))
                    except Exception as e:
                        self._ui(lambda f=f, e=e: self.logline(f"  ✗ {Path(f).name}：{e}"))
                out_txt = Path(a["out_dir"]) / "统计结果.txt"
                out_txt.write_text("\n".join(lines), encoding="utf-8")
                self._ui(lambda: self.logline(""))
                self._ui(lambda: self.logline(f"统计结果 → {out_txt}"))

            elif kind == "merge":
                total, oks = merge_docs(a["files"], a["out"], add_page_break=a["add_page_break"])
                for f in oks:
                    self._ui(lambda f=f: self.logline(f"  {Path(f).name}"))
                self._ui(lambda: self.logline(""))
                self._ui(lambda t=total, o=a["out"]: self.logline(
                    f"合并完成：共 {t} 段落 → {o}"))

            elif kind == "watermark_text":
                for f in a["files"]:
                    try:
                        out_p = unique_path(Path(a["out_dir"]) / f"{Path(f).stem}_水印.docx")
                        add_text_watermark(f, str(out_p), a["text"],
                                         opacity=a["opacity"], font_size=a["font_size"],
                                         rotation=a["rotation"])
                        self._ui(lambda out_p=out_p: self.logline(f"  → {out_p}"))
                    except Exception as e:
                        self._ui(lambda f=f, e=e: self.logline(f"  ✗ {Path(f).name}：{e}"))
                self._ui(lambda: self.logline(""))
                self._ui(lambda: self.logline("文字水印完成"))

            elif kind == "watermark_pic":
                for f in a["files"]:
                    try:
                        out_p = unique_path(Path(a["out_dir"]) / f"{Path(f).stem}_水印.docx")
                        add_picture_watermark(f, str(out_p), a["img_path"],
                                             opacity=a["opacity"])
                        self._ui(lambda out_p=out_p: self.logline(f"  → {out_p}"))
                    except Exception as e:
                        self._ui(lambda f=f, e=e: self.logline(f"  ✗ {Path(f).name}：{e}"))
                self._ui(lambda: self.logline(""))
                self._ui(lambda: self.logline("图片水印完成"))

            elif kind == "text":
                for f in a["files"]:
                    try:
                        out_p = unique_path(Path(a["out_dir"]) / f"{Path(f).stem}.txt")
                        n, chars = export_text(f, str(out_p))
                        self._ui(lambda n=n, chars=chars, f=f: self.logline(
                            f"  {Path(f).name}  {n} 行 {chars} 字符 → {out_p}"))
                    except Exception as e:
                        self._ui(lambda f=f, e=e: self.logline(f"  ✗ {Path(f).name}：{e}"))
                self._ui(lambda: self.logline(""))
                self._ui(lambda: self.logline("文本导出完成"))

        except Exception as e:
            self._ui(lambda: self.logline(f"✗ 错误：{e}"))
            traceback.print_exc()
        finally:
            self._ui(lambda: self._done())

    def _done(self):
        self.running = False
        self.run_btn.configure(state="normal")
        self.progress.stop()
        self.status.configure(text="完成")


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
