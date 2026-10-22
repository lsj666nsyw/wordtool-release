# -*- coding: utf-8 -*-
"""Word 批量处理助手 发布打包"""
import shutil
import zipfile
from pathlib import Path

BASE = Path(r"D:\21821\products")


def md_to_txt(md_path: Path, txt_path: Path):
    text = md_path.read_text(encoding="utf-8")
    out = []
    for line in text.splitlines():
        s = line
        if s.startswith("#"):
            s = s.lstrip("#").strip()
            s = s + "\n" + "=" * min(40, len(s) * 2)
        s = s.replace("**", "").replace("`", "")
        out.append(s)
    txt_path.write_text("\n".join(out), encoding="utf-8-sig")


NOTE = (
    "【重要提醒】\r\n\r\n"
    "1. 解压：请右键 zip → 「全部解压缩」到一个文件夹，然后再双击里面的 exe。\r\n"
    "   不要直接在压缩软件里双击 exe（会缺少同目录的运行库，报各种错）。\r\n\r\n"
    "2. 文件夹不要改名、不要拆开：exe 旁边那个 _internal 文件夹是程序本体，必须放在一起。\r\n\r\n"
    "3. 如果 Windows 弹出「已保护你的电脑」：点「更多信息」→「仍要运行」。\r\n"
    "   如果腾讯电脑管家/火绒拦截，请选择「允许运行」或加入信任区。\r\n"
    "   本软件完全本地运行、不联网、无任何后台行为，可放心。\r\n\r\n"
    "4. 可以整个文件夹拷到 U 盘、换电脑用，不限设备。\r\n"
)


def pack_onedir(name, app_dir, readme_md, out_zip, note_txt=None):
    app_dir = Path(app_dir)
    assert app_dir.is_dir(), f"目录不存在: {app_dir}"
    out_zip = Path(out_zip)
    out_zip.parent.mkdir(parents=True, exist_ok=True)
    if out_zip.exists():
        out_zip.unlink()

    stage = BASE / "_stage_wordtool" / name
    if stage.exists():
        shutil.rmtree(stage)
    shutil.copytree(app_dir, stage / app_dir.name)
    inner = stage / app_dir.name
    md_to_txt(readme_md, inner / "使用说明.txt")
    if note_txt:
        (inner / "重要提醒.txt").write_text(note_txt, encoding="utf-8-sig")

    with zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED, compresslevel=5) as z:
        for f in stage.rglob("*"):
            if f.is_file():
                z.write(f, f.relative_to(stage))
    shutil.rmtree(BASE / "_stage_wordtool", ignore_errors=True)
    size_mb = out_zip.stat().st_size / 1048576
    print(f"  {out_zip.name}  {size_mb:.1f} MB")


def main():
    print("打包 Word 批量处理助手：")
    pack_onedir(
        "Word批量处理助手",
        BASE / "wordtool" / "dist" / "Word批量处理助手",
        BASE / "wordtool" / "使用说明.md",
        BASE / "发布" / "Word批量处理助手_v1.0.zip", NOTE)
    print("\n完成。")


if __name__ == "__main__":
    main()
