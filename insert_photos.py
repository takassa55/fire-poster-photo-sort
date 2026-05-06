"""
template_ホームページ用.docx をもとにホームページ用_完成.docx を作成するスクリプト

【処理の流れ】
  template_ホームページ用.docx（テンプレート）
      ↓ コピーして土台にする
  各表の写真セルに「写真」フォルダの対応画像を挿入（縦110mm）
  各表のファイル名セルに画像ファイル名（拡張子なし）を記載
      ↓
  ホームページ用_完成.docx として出力

【前提フォルダ構成】
作業フォルダ/
├── template_ホームページ用.docx  ← テンプレート（変更しない）
├── insert_photos.py               ← このスクリプト
└── 写真/
    ├── 00_大垣市長賞/
    │   └── （写真ファイル）
    ├── 01_神戸町長賞/
    │   └── （写真ファイル）
    └── ...（以降 02, 03, ... と続く）

【HEIC/HEIF 対応】
  pillow-heif をインストールすると HEIC ファイルを直接処理できます。
    pip install pillow-heif
  未インストールの場合は ffmpeg または ImageMagick で自動変換を試みます。
"""

import os
import sys
import shutil
import zipfile
import re
import subprocess
import io
from pathlib import Path

# ── Pillow と pillow-heif の読み込み ─────────────────────────────────
from PIL import Image

HEIF_AVAILABLE = False
try:
    import pillow_heif
    pillow_heif.register_heif_opener()   # Image.open() で HEIC を透過的に開けるようにする
    HEIF_AVAILABLE = True
    print("[情報] pillow-heif 有効 → HEIC/HEIF ファイルを直接処理します")
except ImportError:
    print("[情報] pillow-heif 未インストール → HEIC は ffmpeg/ImageMagick で変換を試みます")
    print("       （確実に動かすには: pip install pillow-heif）")


# ============================================================
# 設定
# ============================================================
TEMPLATE_DOCX    = "template_ホームページ用.docx"  # 入力テンプレート（コピーして使用）
OUTPUT_DOCX      = "ホームページ用_完成.docx"       # 出力ファイル名
PHOTO_DIR        = "写真"                            # 写真フォルダ名
TARGET_HEIGHT_MM = 110                               # 写真の縦サイズ（mm）

EMU_PER_MM        = 914400 / 25.4              # 1 mm → EMU
TARGET_HEIGHT_EMU = int(TARGET_HEIGHT_MM * EMU_PER_MM)

# 検索対象の画像拡張子（HEIC/HEIF を先に含める）
IMAGE_EXTENSIONS = [
    ".heic", ".heif",                          # iPhone 形式
    ".jpg", ".jpeg", ".png",
    ".gif", ".bmp", ".tiff", ".tif",
]


# ============================================================
# ① ファイル検索
# ============================================================

def find_photo_in_folder(folder_path: Path):
    """
    フォルダ内の最初の画像ファイルを返す。
    glob の大文字小文字問題を避けるため、全ファイルを列挙して拡張子で比較する。
    HEIC/HEIF を優先して返す（pillow-heif で確実に読めるため）。
    """
    ext_set = {e.lower() for e in IMAGE_EXTENSIONS}
    all_files = sorted([f for f in folder_path.iterdir()
                        if f.is_file() and f.suffix.lower() in ext_set])
    heic = [f for f in all_files if f.suffix.lower() in (".heic", ".heif")]
    other = [f for f in all_files if f.suffix.lower() not in (".heic", ".heif")]
    ordered = heic + other
    if ordered:
        print(f"    [検索] {folder_path.name}: {ordered[0].name} を選択")
    return ordered[0] if ordered else None


def get_folder_by_index(photo_base: Path, index: int):
    """'02' や '02_xxx' のようにフォルダ名が2桁数字で始まるものを返す"""
    prefix = f"{index:02d}"
    candidates = [p for p in photo_base.iterdir()
                  if p.is_dir() and p.name.startswith(prefix)]
    return candidates[0] if candidates else None


# ============================================================
# ② HEIC 判定・変換
# ============================================================

def is_heic_by_magic(path: Path) -> bool:
    """
    先頭 12 バイトで HEIC/HEIF を判定する。
    ISOBMFF 系は offset=4 から 'ftyp' が来て、
    その後に 'heic'/'heix'/'hevc'/'mif1'/'msf1'/'avif' などが続く。
    """
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
    """pillow-heif 経由で HEIC → JPEG 変換"""
    if not HEIF_AVAILABLE:
        return False
    try:
        heif_file = pillow_heif.open_heif(str(src), convert_hdr_to_8bit=True)
        img = Image.frombytes(
            heif_file.mode,
            heif_file.size,
            heif_file.data,
            "raw",
        )
        img.save(str(dst), format="JPEG", quality=95)
        return dst.exists() and dst.stat().st_size > 0
    except Exception as e:
        print(f"    [pillow-heif 変換エラー] {e}")
        return False


def convert_heic_with_image_open(src: Path, dst: Path) -> bool:
    """pillow-heif が register 済みのとき Image.open() で変換"""
    if not HEIF_AVAILABLE:
        return False
    try:
        with Image.open(src) as img:
            img_rgb = img.convert("RGB")
            img_rgb.save(str(dst), format="JPEG", quality=95)
        return dst.exists() and dst.stat().st_size > 0
    except Exception as e:
        print(f"    [Image.open 変換エラー] {e}")
        return False


def convert_heic_with_external(src: Path, dst: Path) -> bool:
    """ffmpeg または ImageMagick で HEIC → JPEG 変換"""
    cmds = [
        ["ffmpeg", "-y", "-i", str(src), str(dst)],
        ["magick", "convert", str(src), str(dst)],
        ["convert", str(src), str(dst)],   # ImageMagick 6 系
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
    """
    画像を Pillow が確実に読めるファイルとして返す。
    - 通常の JPEG/PNG 等 → そのまま返す
    - HEIC → JPEG に変換して tmp_jpeg_path を返す
    - 変換失敗 → None を返す
    """
    ext = photo_path.suffix.lower()
    # 拡張子が heic/heif なら必ず変換（マジックバイト判定に頼らない）
    need_convert = ext in (".heic", ".heif") or is_heic_by_magic(photo_path)

    if not need_convert:
        try:
            with Image.open(photo_path) as img:
                img.load()   # verify() だと PNG 等で誤判定があるので load() を使う
            return photo_path
        except Exception as e:
            print(f"    [Pillow 読み込みエラー: {e}] HEIC として変換を試みます")
            need_convert = True

    print(f"    [変換中] {photo_path.name} → JPEG ...")

    # 方法1: pillow-heif 直接変換
    if convert_heic_with_pillow_heif(photo_path, tmp_jpeg_path):
        print(f"    [変換OK] pillow-heif (open_heif) 使用")
        return tmp_jpeg_path

    # 方法2: Image.open（register_heif_opener 経由）
    if convert_heic_with_image_open(photo_path, tmp_jpeg_path):
        print(f"    [変換OK] Image.open (pillow-heif) 使用")
        return tmp_jpeg_path

    # 方法3: 外部ツール（ffmpeg / ImageMagick）
    if convert_heic_with_external(photo_path, tmp_jpeg_path):
        print(f"    [変換OK] 外部ツール使用")
        return tmp_jpeg_path

    print(f"    [変換失敗] 以下のいずれかをインストールしてください:")
    print(f"      pip install pillow-heif")
    print(f"      または ffmpeg / ImageMagick をインストール")
    return None


# ============================================================
# ③ 画像サイズ取得・EMU 計算
# ============================================================

def get_image_size(image_path: Path):
    """(幅px, 高さpx) を返す。取得できなければ (1, 1)"""
    try:
        with Image.open(image_path) as img:
            return img.size
    except Exception:
        pass
    print(f"    [警告] 画像サイズ取得不可 → 縦横比 1:1 で処理: {image_path.name}")
    return 1, 1


def calc_emu_width(image_path: Path, height_emu: int) -> int:
    """アスペクト比を保ちながら幅(EMU)を計算する"""
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
# ⑥ メイン処理
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

    # ── テンプレートを ZIP 展開（これが処理の土台）──
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

    # ── DOM 解析（テンプレートをそのまま使う）──
    tree = etree.parse(str(doc_xml_path))
    root = tree.getroot()
    tables = root.findall(f".//{{{W}}}tbl")

    # スロット列挙（テンプレートの表構造から生成）
    slots = []
    for t_idx, tbl in enumerate(tables):
        rows = tbl.findall(f"{{{W}}}tr")
        if len(rows) == 3:
            n = len(rows[1].findall(f"{{{W}}}tc"))
            for c in range(n):
                slots.append((t_idx, 1, 2, c))
        elif len(rows) == 6:
            for photo_row, fname_row in [(1, 2), (4, 5)]:
                n = len(rows[photo_row].findall(f"{{{W}}}tc"))
                for c in range(n):
                    slots.append((t_idx, photo_row, fname_row, c))

    print(f"スロット総数: {len(slots)}\n")

    inserted = skipped = 0

    for folder_idx, (t_idx, photo_row, fname_row, col_idx) in enumerate(slots):

        folder_path = get_folder_by_index(photo_base, folder_idx)
        if folder_path is None:
            print(f"  [{folder_idx:02d}] フォルダなし → スキップ")
            skipped += 1
            continue

        photo_path = find_photo_in_folder(folder_path)
        if photo_path is None:
            print(f"  [{folder_idx:02d}] {folder_path.name} → 画像なし → スキップ")
            skipped += 1
            continue

        stem     = photo_path.stem
        tmp_jpeg = media_dir / f"img_{folder_idx:03d}_tmp.jpg"

        # 画像の準備（HEIC 変換含む）
        ready_path = prepare_image(photo_path, tmp_jpeg)
        if ready_path is None:
            print(f"  [{folder_idx:02d}] {folder_path.name} / {photo_path.name} → 変換失敗 → スキップ")
            skipped += 1
            continue

        # EMU サイズ計算（変換後ファイルで行う）
        cy = TARGET_HEIGHT_EMU
        cx = calc_emu_width(ready_path, cy)

        # メディアフォルダへコピー
        final_ext        = ready_path.suffix.lower()
        media_name       = f"img_{folder_idx:03d}{final_ext}"
        final_media_path = media_dir / media_name
        if ready_path != final_media_path:
            shutil.copy2(ready_path, final_media_path)
        if tmp_jpeg.exists() and tmp_jpeg != final_media_path:
            tmp_jpeg.unlink()

        # リレーションシップ追加
        r_id     = f"rId{next_rid}"
        next_rid += 1
        rels_xml = add_image_relationship(rels_xml, r_id, media_name)
        ct_xml   = ensure_content_type(ct_xml, final_ext)

        # DOM 更新
        tbl  = tables[t_idx]
        rows = tbl.findall(f"{{{W}}}tr")

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

        print(f"  [{folder_idx:02d}] {folder_path.name} / {photo_path.name} → 挿入OK")
        inserted += 1

    # 書き戻し
    doc_xml_path.write_bytes(
        etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    )
    rels_xml_path.write_text(rels_xml, encoding="utf-8")
    ct_xml_path.write_text(ct_xml, encoding="utf-8")

    # ZIP 再パック
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
