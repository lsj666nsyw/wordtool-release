# -*- coding: utf-8 -*-
"""
Word 批量处理助手 —— 引擎层

依赖：
  · python-docx  MIT   —— DOCX 结构读写（段落/表格/页眉页脚/图片/属性）
  · Pillow      MIT   —— 图片水印 / 图片尺寸
  · lxml（python-docx 自带）

功能：
  1. 批量查找替换（段落 + 表格 + 页眉页脚）
  2. 提取文档内嵌图片
  3. 文档统计（段落/表格/图片/字数）
  4. 文档合并
  5. 图片 / 文字水印
  6. 导出纯文本
"""
import os
import re
import shutil
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Inches, Emu
from PIL import Image

APP_NAME = "Word 批量处理助手"
APP_VERSION = "1.0.0"

DOC_EXT = (".docx", ".docm")
IMG_EXT = (".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff", ".gif", ".emf", ".wmf")


# ============================================================================
# 小工具
# ============================================================================
def is_doc(p):
    return Path(p).suffix.lower() in DOC_EXT


def human_size(n):
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    for i in range(1, 10000):
        cand = path.with_name(f"{stem}_{i}{suffix}")
        if not cand.exists():
            return cand
    return path


def safe_name(s, maxlen=60):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", str(s)).strip()
    s = re.sub(r"_+", "_", s).strip("_ .")
    return (s or "未命名")[:maxlen]


# ============================================================================
# 段落遍历（含表格 / 页眉页脚）
# ============================================================================
def iter_paragraphs(doc):
    """遍历正文段落 + 表格单元格段落 + 页眉页脚段落。"""
    def walk_blocks(parent_elm):
        from docx.table import Table
        from docx.text.paragraph import Paragraph
        for child in parent_elm:
            if child.tag == qn('w:p'):
                yield Paragraph(child, None)
            elif child.tag == qn('w:tbl'):
                tbl = Table(child, None)
                for row in tbl.rows:
                    for cell in row.cells:
                        yield from walk_blocks(cell._tc)
    yield from walk_blocks(doc.element.body)
    for section in doc.sections:
        for hf in (section.header, section.footer,
                   section.first_page_header, section.first_page_footer,
                   section.even_page_header, section.even_page_footer):
            try:
                yield from walk_blocks(hf._element)
            except Exception:
                pass


def _split_runs(text):
    return [r for r in text.split("\n") if r != ""]


# ============================================================================
# 1. 批量查找替换
# ============================================================================
def _replace_in_paragraph(paragraph, find_text, replace_text, case_sensitive, count_ref):
    """
    在单个段落内替换。返回是否发生变化。

    实现思路：把段落所有 run 的文字拼成一个字符串，在整体串上做替换，
    再按原 run 的切分位置把新串逐段写回（保留原 run 的字体/颜色/格式）。
    """
    runs = paragraph.runs
    if not runs:
        return False
    texts = [r.text for r in runs]
    full = "".join(texts)
    if find_text == "":
        return False
    if case_sensitive:
        hits = full.count(find_text)
        if hits == 0:
            return False
        new_full = full.replace(find_text, replace_text)
    else:
        low_find = find_text.lower()
        if low_find == "":
            return False
        new_full, hits = _casefold_replace(full, find_text, replace_text)
    if new_full == full:
        return False
    # 把新串写回：按原 run 的字数切分（text 里可能有 \n 但 run 拆分按 run 边界）
    new_texts = _redistribute(texts, new_full)
    changed = False
    for r, t in zip(runs, new_texts):
        if r.text != t:
            r.text = t
            changed = True
    if changed:
        count_ref[0] += hits
    return changed


def _casefold_replace(full, find_text, replace_text):
    """大小写不敏感的 replace（保持原文其余部分不变）。"""
    n = 0
    out = []
    i = 0
    low_full = full.lower()
    low_find = find_text.lower()
    while i < len(full):
        j = low_full.find(low_find, i)
        if j < 0:
            out.append(full[i:])
            break
        out.append(full[i:j])
        out.append(replace_text)
        i = j + len(find_text)
        n += 1
    return "".join(out), n


def _redistribute(old_texts, new_full):
    """
    把 new_full 按 old_texts 的长度比例切回各 run。
    保证各 run 的字符数与原 run 完全一致（避免 run 数量变化导致格式错位）。
    """
    old_lens = [len(t) for t in old_texts]
    total_old = sum(old_lens)
    if total_old == 0:
        return old_texts
    new_lens = [len(new_full) * ol // total_old for ol in old_lens]
    # 修正舍入误差
    diff = len(new_full) - sum(new_lens)
    i = 0
    while diff != 0:
        if i >= len(new_lens):
            i = 0
        if diff > 0:
            new_lens[i] += 1
            diff -= 1
        elif new_lens[i] > 0:
            new_lens[i] -= 1
            diff += 1
        i += 1
    out = []
    pos = 0
    for nl in new_lens:
        out.append(new_full[pos:pos + nl])
        pos += nl
    return out


def batch_replace(files, out_dir, find_text, replace_text, case_sensitive=False,
                  all_sheets=True):
    """
    批量查找替换。在所有段落、表格、页眉页脚中替换文本。
    返回 (files, count, out_path, error)
    """
    out_root = Path(out_dir)
    out_root.mkdir(parents=True, exist_ok=True)
    results = []
    for f in files:
        p = Path(f)
        if not is_doc(p):
            results.append((f, 0, None, "只支持 .docx / .docm"))
            continue
        try:
            doc = Document(str(p))
            count_ref = [0]
            for para in iter_paragraphs(doc):
                _replace_in_paragraph(para, find_text, replace_text,
                                     case_sensitive, count_ref)
            out_p = unique_path(out_root / p.name)
            doc.save(out_p)
            results.append((f, count_ref[0], out_p, None))
        except Exception as e:
            results.append((f, 0, None, str(e)))
    return results


# ============================================================================
# 2. 提取文档内嵌图片
# ============================================================================
def extract_images(doc_path, out_dir):
    """从 DOCX 中提取所有内嵌图片，保存到 out_dir。返回 (文件列表, 错误)"""
    out_root = Path(out_dir)
    out_root.mkdir(parents=True, exist_ok=True)
    doc = Document(str(doc_path))
    rels = doc.part.related_parts
    results = []
    seen = set()
    for rid, part in list(rels.items()):
        try:
            if "image" not in part.content_type:
                continue
            data = part.blob
            stem = safe_name(part.partname.rsplit("/", 1)[-1].split(".")[0] if part.partname else "image")
            ext = part.partname.rsplit(".", 1)[-1].lower() if part.partname else "png"
            if ext not in ("jpg", "jpeg", "png", "gif", "bmp", "tif", "tiff", "webp"):
                ext = "png"
            fn = f"{stem}.{ext}"
            fp = unique_path(out_root / fn)
            fp.write_bytes(data)
            results.append(fp)
        except Exception:
            continue
    return results


# ============================================================================
# 3. 文档统计
# ============================================================================
def doc_stats(doc_path):
    """
    返回字典：段落数、表格数、图片数、字符数、行数、
    以及文档属性（标题、作者、创建时间等）。
    """
    doc = Document(str(doc_path))
    para_count = len(doc.paragraphs)
    tbl_count = len(doc.tables)
    img_count = 0
    for rel in doc.part.related_parts.values():
        if "image" in rel.content_type:
            img_count += 1
    text = "\n".join(p.text for p in doc.paragraphs)
    char_count = len(text)
    line_count = sum(1 for line in text.split("\n") if line.strip())
    props = doc.core_properties
    stats = {
        "段落数": para_count,
        "表格数": tbl_count,
        "图片数": img_count,
        "字符数": char_count,
        "非空行数": line_count,
        "标题": props.title or "",
        "作者": props.author or "",
        "创建时间": str(props.created or ""),
        "修改时间": str(props.modified or ""),
    }
    return stats


# ============================================================================
# 4. 文档合并
# ============================================================================
def merge_docs(files, out_path, add_page_break=True):
    """
    合并多个 DOCX。第一个文件作为基础，其余追加到末尾。
    返回 (总段落数, 成功文件列表)
    """
    if not files:
        raise ValueError("没有要合并的文件")
    base = Document(str(files[0]))
    total_paras = len(base.paragraphs)
    ok_files = [files[0]]
    for f in files[1:]:
        try:
            src = Document(str(f))
            # 插入分页符
            if add_page_break:
                base.add_page_break()
            for para in src.paragraphs:
                base.add_paragraph(para.text)
            # 追加表格
            for tbl in src.tables:
                new_tbl = base.add_table(rows=len(tbl.rows), cols=len(tbl.columns))
                for r, row in enumerate(tbl.rows):
                    for c, cell in enumerate(row.cells):
                        new_tbl.cell(r, c).text = cell.text
            total_paras += len(src.paragraphs)
            ok_files.append(f)
        except Exception:
            continue
    base.save(out_path)
    return total_paras, ok_files


# ============================================================================
# 5. 水印
# ============================================================================
def add_text_watermark(doc_path, out_path, text, opacity=15, font_size=40, rotation=45):
    """
    给 DOCX 每页加文字水印（通过页眉图片实现，兼容性好）。
    """
    doc = Document(str(doc_path))
    # 用 PIL 生成半透明文字图
    from PIL import Image, ImageDraw, ImageFont
    import io

    W, H = 1200, 1600
    img = Image.new("RGBA", (W, H), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("msyh.ttc", font_size)
    except Exception:
        font = ImageFont.load_default()
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    fill = (128, 128, 128, int(255 * opacity / 100))
    x, y = (W - tw) // 2, (H - th) // 2
    draw.text((x, y), text, font=font, fill=fill)
    # 旋转
    img = img.rotate(rotation, expand=True, resample=Image.BICUBIC)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)

    for section in doc.sections:
        header = section.header
        header.is_linked_to_previous = False
        p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        run = p.add_run()
        run.add_picture(buf, width=Inches(5))
    doc.save(out_path)
    return len(doc.sections)


def add_picture_watermark(doc_path, out_path, img_path, opacity=30):
    """给 DOCX 每页加图片水印。"""
    doc = Document(str(doc_path))
    img = Image.open(img_path).convert("RGBA")
    # 简单透明化
    alpha = img.split()[3]
    alpha = alpha.point(lambda a: int(a * opacity / 100))
    img.putalpha(alpha)
    tmp = Path(out_path).with_suffix(".png")
    img.save(tmp, format="PNG")
    for section in doc.sections:
        header = section.header
        header.is_linked_to_previous = False
        p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        p.add_run().add_picture(str(tmp), width=Inches(3))
    doc.save(out_path)
    try:
        tmp.unlink()
    except Exception:
        pass
    return len(doc.sections)


# ============================================================================
# 6. 导出纯文本
# ============================================================================
def export_text(doc_path, out_path):
    """把 DOCX 导出为纯文本 txt。"""
    doc = Document(str(doc_path))
    lines = []
    for p in doc.paragraphs:
        lines.append(p.text)
    for tbl in doc.tables:
        for row in tbl.rows:
            lines.append("\t".join(cell.text for cell in row.cells))
    content = "\n".join(lines)
    Path(out_path).write_text(content, encoding="utf-8")
    return len(lines), len(content)


if __name__ == "__main__":
    print(f"{APP_NAME} v{APP_VERSION} 引擎层加载成功")
    print("可用功能：")
    for fn in ("batch_replace", "extract_images", "doc_stats",
               "merge_docs", "add_text_watermark", "add_picture_watermark",
               "export_text"):
        print(f"  - {fn}")
