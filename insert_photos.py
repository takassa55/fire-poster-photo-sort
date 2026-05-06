"""
template_ホームページ用.docx をもとにホームページ用_完成.docx を作成するスクリプト

【処理の流れ】
  template_ホームページ用.docx（テンプレート）
      ↓ コピーして土台にする
  ★ フェーズ1（連番方式）
      各表の写真セルに「写真」フォルダの連番フォルダ（00_xxx, 01_xxx ...）の画像を挿入
      → 金賞テーブルを対象とする
  ★ フェーズ2（賞カテゴリ方式）
      【銀賞】【銅賞】【入選】の各セクションに対して、
      対応する賞名を含むフォルダ（例: 銀賞_xxx）の画像を順番に挿入
      ↓
  ホームページ用_完成.docx として出力

【前提フォルダ構成】
作業フォルダ/
├── template_ホームページ用.docx  ← テンプレート（変更しない）
├── insert_photos.py               ← このスクリプト
└── 写真/
    ├── 00_大垣市長賞/             ← 金賞（連番）
    │   └── （写真ファイル）
    ├── 01_神戸町長賞/
    │   └── （写真ファイル）
    ├── ...（以降 02, 03, ... と続く）
    ├── 銀賞_xxx/                  ← 銀賞（「銀賞」を含む名前）
    │   └── （写真ファイル複数可）
    ├── 銅賞_xxx/                  ← 銅賞（「銅賞」を含む名前）
    │   └── （写真ファイル）
    └── 入選_xxx/                  ← 入選（「入選」を含む名前）
        └── （写真ファイル）

【HEIC/HEIF 対応】
  pillow-heif をインストールすると HEIC ファイルを直接処理できます。
    pip install pillow-heif
  未インストールの場合は ffmpeg または ImageMagick で自動変換を試みます。
"""

import sys
import shutil
import zipfile
import re
import subprocess
from pathlib import Path

# ── Pillow と pillow-heif の読み込み ─────────────────────────────────
from PIL import Image

HEIF_AVAILABLE = False
try:
    import pillow_heif
    pillow_heif.register_heif_opener()
    HEIF_AVAILABLE = True
    print("[情報] pillow-heif 有効 → HEIC/HEIF ファイルを直接処理します")
except ImportError:
    print("[情報] pillow-heif 未インストール → HEIC は ffmpeg/ImageMagick で変換を試みます")
    print("       （確実に動かすには: pip install pillow-heif）")


# ============================================================
# 設定
# ============================================================
TEMPLATE_DOCX    = "template_ホームページ用.docx"
OUTPUT_DOCX      = "ホームページ用_完成.docx"
PHOTO_DIR        = "写真"
TARGET_HEIGHT_MM = 110

EMU_PER_MM        = 914400 / 25.4
TARGET_HEIGHT_EMU = int(TARGET_HEIGHT_MM * EMU_PER_MM)

# フェーズ1（連番方式）で処理する賞の見出し
SEQUENTIAL_AWARD = "金賞"

# フェーズ2（カテゴリ名マッチ）で処理する賞の見出しとフォルダキーワード
CATEGORY_AWARDS = ["銀賞", "銅賞", "入選"]

IMAGE_EXTENSIONS = [
    ".heic", ".heif",
    ".jpg", ".jpeg", ".png",
    ".gif", ".bmp", ".tiff", ".tif",
]


# ============================================================
# ① ファイル検索
# ============================================================

def find_photo_in_folder(folder_path: Path):
    """フォルダ内の最初の画像ファイルを返す（HEIC/HEIF 優先）"""
    ext_set = {e.lower() for e in IMAGE_EXTENSIONS}
    all_files = sorted([f for f in folder_path.iterdir()
                        if f.is_file() and f.suffix.lower() in ext_set])
    heic  = [f for f in all_files if f.suffix.lower() in (".heic", ".heif")]
    other = [f for f in all_files if f.suffix.lower() not in (".heic", ".heif")]
    ordered = heic + other
    if ordered:
        print(f"    [検索] {folder_path.name}: {ordered[0].name} を選択")
    return ordered[0] if ordered else None


def find_all_photos_in_folder(folder_path: Path):
    """フォルダ内の全画像ファイルをソート順（HEIC/HEIF 優先）で返す"""
    ext_set = {e.lower() for e in IMAGE_EXTENSIONS}
    all_files = sorted([f for f in folder_path.iterdir()
                        if f.is_file() and f.suffix.lower() in ext_set])
    heic  = [f for f in all_files if f.suffix.lower() in (".heic", ".heif")]
    other = [f for f in all_files if f.suffix.lower() not in (".heic", ".heif")]
    return heic + other


def get_folder_by_index(photo_base: Path, index: int):
    """'02' や '02_xxx' のようにフォルダ名が2桁数字で始まるものを返す"""
    prefix = f"{index:02d}"
    candidates = [p for p in photo_base.iterdir()
                  if p.is_dir() and p.name.startswith(prefix)]
    return candidates[0] if candidates else None


def collect_photos_for_award(photo_base: Path, award_name: str):
    """
    award_name を含む全フォルダから画像ファイルを順番に収集して返す。
    戻り値: [(folder_path, photo_path), ...]
    """
    folders = sorted([p for p in photo_base.iterdir()
                      if p.is_dir() and award_name in p.name])
    if not folders:
        print(f"  [警告] 「{award_name}」を含むフォルダが見つかりません")
        return []
    result = []
    for folder in folders:
        for photo in find_all_photos_in_folder(folder):
            result.append((folder, photo))
    print(f"  [{award_name}] {len(folders)}フォルダ / {len(result)}枚の画像を収集")
    return result


# ============================================================
# ② HEIC 判定・変換
# ============================================================

def is_heic_by_magic(path: Path) -> bool:
    try:
        header = path.read_bytes()[:12]
        if header[4:8] != b"ftyp":
            return False
        brand = header[8:12]
        HEIC_BRANDS = {b"heic", b"heix", b"hevc", b"hevx",
                       b"heim", b"heis", b"hevm", b"hevs",
                       b"mif1", b"msf1", b"avif", b"avis"}
        return brand in HEIC_BRANDS
    except Exception:
        return False


def convert_heic_with_pillow_heif(src: Path, dst: Path) -> bool:
    if not HEIF_AVAILABLE:
        return False
    try:
        heif_file = pillow_heif.open_heif(str(src), convert_hdr_to_8bit=True)
        img = Image.frombytes(heif_file.mode, heif_file.size, heif_file.data, "raw")
        img.save(str(dst), format="JPEG", quality=95)
        return dst.exists() and dst.stat().st_size > 0
    except Exception as e:
        print(f"    [pillow-heif 変換エラー] {e}")
        return False


def convert_heic_with_image_open(src: Path, dst: Path) -> bool:
    if not HEIF_AVAILABLE:
        return False
    try:
        with Image.open(src) as img:
            img.convert("RGB").save(str(dst), format="JPEG", quality=95)
        return dst.exists() and dst.stat().st_size > 0
    except Exception as e:
        print(f"    [Image.open 変換エラー] {e}")
        return False


def convert_heic_with_external(src: Path, dst: Path) -> bool:
    cmds = [
        ["ffmpeg", "-y", "-i", str(src), str(dst)],
        ["magick", "convert", str(src), str(dst)],
        ["convert", str(src), str(dst)],
    ]
    for cmd in cmds:
        try:
            result = subprocess.run(cmd, capture_output=True, timeout=60)
            if result.returncode == 0 and dst.exists() and dst.stat().st_size > 0:
                return True
        except FileNotFoundError:
            continue
        except Exception as e:
            print(f"    [{cmd[0]} エラー] {e}")
    return False


def prepare_image(photo_path: Path, tmp_jpeg_path: Path):
    """画像を Pillow が確実に読めるファイルとして返す。変換失敗時は None。"""
    ext = photo_path.suffix.lower()
    need_convert = ext in (".heic", ".heif") or is_heic_by_magic(photo_path)

    if not need_convert:
        try:
            with Image.open(photo_path) as img:
                img.load()
            return photo_path
        except Exception as e:
            print(f"    [Pillow 読み込みエラー: {e}] HEIC として変換を試みます")
            need_convert = True

    print(f"    [変換中] {photo_path.name} → JPEG ...")
    if convert_heic_with_pillow_heif(photo_path, tmp_jpeg_path):
        print(f"    [変換OK] pillow-heif (open_heif) 使用")
        return tmp_jpeg_path
    if convert_heic_with_image_open(photo_path, tmp_jpeg_path):
        print(f"    [変換OK] Image.open (pillow-heif) 使用")
        return tmp_jpeg_path
    if convert_heic_with_external(photo_path, tmp_jpeg_path):
        print(f"    [変換OK] 外部ツール使用")
        return tmp_jpeg_path

    print(f"    [変換失敗] pip install pillow-heif または ffmpeg/ImageMagick をインストールしてください")
    return None


# ============================================================
# ③ 画像サイズ取得・EMU 計算
# ============================================================

def get_image_size(image_path: Path):
    try:
        with Image.open(image_path) as img:
            return img.size
    except Exception:
        pass
    print(f"    [警告] 画像サイズ取得不可 → 縦横比 1:1 で処理: {image_path.name}")
    return 1, 1


def calc_emu_width(image_path: Path, height_emu: int) -> int:
    w_px, h_px = get_image_size(image_path)
    return int(height_emu * w_px / h_px) if h_px else height_emu


# ============================================================
# ④ Word XML 生成
# ============================================================

def make_drawing_xml(r_id: str, img_name: str, cx: int, cy: int) -> str:
    return (
        f'<w:drawing'
        f' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        f' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        f' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        f' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
        f' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="1" name="{img_name}"/>'
        f'<wp:cNvGraphicFramePr/>'
        f'<a:graphic>'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic>'
        f'<pic:nvPicPr>'
        f'<pic:cNvPr id="0" name="{img_name}"/>'
        f'<pic:cNvPicPr/>'
        f'</pic:nvPicPr>'
        f'<pic:blipFill>'
        f'<a:blip r:embed="{r_id}"/>'
        f'<a:stretch><a:fillRect/></a:stretch>'
        f'</pic:blipFill>'
        f'<pic:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'</pic:spPr>'
        f'</pic:pic>'
        f'</a:graphicData>'
        f'</a:graphic>'
        f'</wp:inline>'
        f'</w:drawing>'
    )


def build_image_paragraph_xml(r_id: str, img_name: str, cx: int, cy: int) -> str:
    drawing = make_drawing_xml(r_id, img_name, cx, cy)
    return (
        f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:pPr><w:jc w:val="center"/></w:pPr>'
        f'<w:r>{drawing}</w:r>'
        f'</w:p>'
    )


def build_filename_paragraph_xml(filename: str) -> str:
    return (
        f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:pPr><w:jc w:val="center"/></w:pPr>'
        f'<w:r><w:t xml:space="preserve">{filename}</w:t></w:r>'
        f'</w:p>'
    )


# ============================================================
# ⑤ .docx（ZIP）操作
# ============================================================

def add_image_relationship(rels_xml: str, r_id: str, img_target: str) -> str:
    tag = (
        f'<Relationship Id="{r_id}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        f'Target="media/{img_target}"/>'
    )
    return rels_xml.replace("</Relationships>", f"  {tag}\n</Relationships>")


def ensure_content_type(ct_xml: str, ext: str) -> str:
    ext_lower = ext.lower().lstrip(".")
    mime_map = {
        "jpg": "image/jpeg", "jpeg": "image/jpeg",
        "png": "image/png",  "gif":  "image/gif",
        "bmp": "image/bmp",  "tiff": "image/tiff", "tif": "image/tiff",
    }
    mime = mime_map.get(ext_lower, "image/jpeg")
    tag  = f'<Default Extension="{ext_lower}" ContentType="{mime}"/>'
    if f'Extension="{ext_lower}"' not in ct_xml:
        ct_xml = ct_xml.replace("</Types>", f"  {tag}\n</Types>")
    return ct_xml


# ============================================================
# ⑥ スロット挿入の共通処理
# ============================================================

def insert_photo_to_slot(
    tbl_elem, photo_row: int, fname_row: int, col_idx: int,
    photo_path: Path, media_dir: Path, img_idx: int,
    rels_xml: str, ct_xml: str, next_rid: int, W: str
):
    """
    1スロット分の画像挿入処理。
    成功時: (updated_rels_xml, updated_ct_xml, next_rid + 1, stem) を返す
    失敗時: (rels_xml, ct_xml, next_rid, None) を返す
    """
    from lxml import etree

    stem     = photo_path.stem
    tmp_jpeg = media_dir / f"img_{img_idx:04d}_tmp.jpg"

    ready_path = prepare_image(photo_path, tmp_jpeg)
    if ready_path is None:
        print(f"    → 変換失敗 → スキップ")
        return rels_xml, ct_xml, next_rid, None

    cy = TARGET_HEIGHT_EMU
    cx = calc_emu_width(ready_path, cy)

    final_ext        = ready_path.suffix.lower()
    media_name       = f"img_{img_idx:04d}{final_ext}"
    final_media_path = media_dir / media_name
    if ready_path != final_media_path:
        shutil.copy2(ready_path, final_media_path)
    if tmp_jpeg.exists() and tmp_jpeg != final_media_path:
        tmp_jpeg.unlink()

    r_id     = f"rId{next_rid}"
    next_rid += 1
    rels_xml = add_image_relationship(rels_xml, r_id, media_name)
    ct_xml   = ensure_content_type(ct_xml, final_ext)

    rows = tbl_elem.findall(f"{{{W}}}tr")

    photo_cell = rows[photo_row].findall(f"{{{W}}}tc")[col_idx]
    for p in photo_cell.findall(f"{{{W}}}p"):
        photo_cell.remove(p)
    photo_cell.append(
        etree.fromstring(build_image_paragraph_xml(r_id, media_name, cx, cy))
    )

    fname_cell = rows[fname_row].findall(f"{{{W}}}tc")[col_idx]
    for p in fname_cell.findall(f"{{{W}}}p"):
        fname_cell.remove(p)
    fname_cell.append(
        etree.fromstring(build_filename_paragraph_xml(stem))
    )

    return rels_xml, ct_xml, next_rid, stem


# ============================================================
# ⑦ ドキュメント走査：賞カテゴリ別スロット収集
# ============================================================

def collect_slots_by_award(root, W: str) -> dict:
    """
    ドキュメント本文を走査し、賞カテゴリごとのスロットリストを返す。
    戻り値: { "金賞": [(tbl_elem, photo_row, fname_row, col), ...], "銀賞": [...], ... }
    """
    body  = root.find(f"{{{W}}}body")
    items = list(body)

    all_awards     = [SEQUENTIAL_AWARD] + CATEGORY_AWARDS
    slots_by_award = {award: [] for award in all_awards}
    current_award  = None

    for elem in items:
        tag = elem.tag.split("}")[1]
        if tag == "p":
            text = "".join(t.text or "" for t in elem.findall(f".//{{{W}}}t")).strip()
            for award in all_awards:
                if text == f"【{award}】":
                    current_award = award
                    break
        elif tag == "tbl" and current_award:
            rows   = elem.findall(f"{{{W}}}tr")
            n_rows = len(rows)
            if n_rows == 3:
                n_cols = len(rows[1].findall(f"{{{W}}}tc"))
                for c in range(n_cols):
                    slots_by_award[current_award].append((elem, 1, 2, c))
            elif n_rows == 6:
                n_cols = len(rows[1].findall(f"{{{W}}}tc"))
                for c in range(n_cols):
                    slots_by_award[current_award].append((elem, 1, 2, c))
                for c in range(n_cols):
                    slots_by_award[current_award].append((elem, 4, 5, c))

    for award in all_awards:
        print(f"  [{award}] スロット数: {len(slots_by_award[award])}")

    return slots_by_award


# ============================================================
# ⑧ メイン処理
# ============================================================

def main():
    from lxml import etree

    base_dir      = Path(__file__).parent
    template_path = base_dir / TEMPLATE_DOCX
    output_path   = base_dir / OUTPUT_DOCX
    photo_base    = base_dir / PHOTO_DIR

    for p, label in [
        (template_path, "テンプレートファイル"),
        (photo_base,    "写真フォルダ"),
    ]:
        if not p.exists():
            print(f"[エラー] {label}が見つかりません: {p}")
            sys.exit(1)

    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    # ── テンプレートを ZIP 展開 ──
    print(f"[テンプレート読み込み] {template_path.name}")
    tmp_dir = base_dir / "_docx_tmp"
    if tmp_dir.exists():
        shutil.rmtree(tmp_dir)
    with zipfile.ZipFile(template_path, "r") as z:
        z.extractall(tmp_dir)

    doc_xml_path  = tmp_dir / "word" / "document.xml"
    rels_xml_path = tmp_dir / "word" / "_rels" / "document.xml.rels"
    ct_xml_path   = tmp_dir / "[Content_Types].xml"
    media_dir     = tmp_dir / "word" / "media"
    media_dir.mkdir(exist_ok=True)

    rels_xml = rels_xml_path.read_text(encoding="utf-8")
    ct_xml   = ct_xml_path.read_text(encoding="utf-8")

    existing_ids = re.findall(r'Id="rId(\d+)"', rels_xml)
    next_rid = max((int(x) for x in existing_ids), default=0) + 1

    tree = etree.parse(str(doc_xml_path))
    root = tree.getroot()

    # ── スロット収集 ──
    print("\n[スロット収集]")
    slots_by_award = collect_slots_by_award(root, W)

    inserted = skipped = 0
    img_idx  = 0  # メディアファイル名用グローバル連番

    # ══════════════════════════════════════════════════════════
    # フェーズ1：連番方式で金賞スロットへ挿入
    # ══════════════════════════════════════════════════════════
    print(f"\n[フェーズ1] {SEQUENTIAL_AWARD}（連番フォルダ方式）")
    seq_slots = slots_by_award[SEQUENTIAL_AWARD]
    print(f"  スロット総数: {len(seq_slots)}")

    for folder_idx, (tbl_elem, photo_row, fname_row, col_idx) in enumerate(seq_slots):
        folder_path = get_folder_by_index(photo_base, folder_idx)
        if folder_path is None:
            print(f"  [{folder_idx:02d}] フォルダなし → スキップ")
            skipped += 1
            img_idx  += 1
            continue

        photo_path = find_photo_in_folder(folder_path)
        if photo_path is None:
            print(f"  [{folder_idx:02d}] {folder_path.name} → 画像なし → スキップ")
            skipped += 1
            img_idx  += 1
            continue

        print(f"  [{folder_idx:02d}] {folder_path.name} / {photo_path.name}")
        rels_xml, ct_xml, next_rid, stem = insert_photo_to_slot(
            tbl_elem, photo_row, fname_row, col_idx,
            photo_path, media_dir, img_idx,
            rels_xml, ct_xml, next_rid, W
        )
        if stem is not None:
            print(f"    → 挿入OK")
            inserted += 1
        else:
            skipped += 1
        img_idx += 1

    # ══════════════════════════════════════════════════════════
    # フェーズ2：賞カテゴリ方式で銀賞・銅賞・入選スロットへ挿入
    # ══════════════════════════════════════════════════════════
    print(f"\n[フェーズ2] {', '.join(CATEGORY_AWARDS)}（賞カテゴリフォルダ方式）")

    for award in CATEGORY_AWARDS:
        print(f"\n  [{award}] 処理開始")
        photos = collect_photos_for_award(photo_base, award)
        slots  = slots_by_award[award]

        if not photos:
            print(f"  → 画像なし、スキップ")
            skipped += len(slots)
            continue

        for slot_idx, (tbl_elem, photo_row, fname_row, col_idx) in enumerate(slots):
            if slot_idx >= len(photos):
                print(f"  [スロット{slot_idx}] 画像不足 → スキップ")
                skipped += 1
                img_idx  += 1
                continue

            folder_path, photo_path = photos[slot_idx]
            print(f"  [スロット{slot_idx}] {folder_path.name} / {photo_path.name}")
            rels_xml, ct_xml, next_rid, stem = insert_photo_to_slot(
                tbl_elem, photo_row, fname_row, col_idx,
                photo_path, media_dir, img_idx,
                rels_xml, ct_xml, next_rid, W
            )
            if stem is not None:
                print(f"    → 挿入OK")
                inserted += 1
            else:
                skipped += 1
            img_idx += 1

    # ── 書き戻し ──
    doc_xml_path.write_bytes(
        etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    )
    rels_xml_path.write_text(rels_xml, encoding="utf-8")
    ct_xml_path.write_text(ct_xml,   encoding="utf-8")

    # ── ZIP 再パック ──
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for fpath in sorted(tmp_dir.rglob("*")):
            if fpath.is_file():
                zout.write(fpath, fpath.relative_to(tmp_dir))

    shutil.rmtree(tmp_dir)

    print()
    print("=" * 50)
    print(f"完了: {inserted} 件挿入 / {skipped} 件スキップ")
    print(f"出力: {output_path}")


if __name__ == "__main__":
    main()