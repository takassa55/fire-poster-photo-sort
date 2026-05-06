"""
HEIC 読み込み診断スクリプト
使い方: python check_heic.py  （HEICファイルが入ったフォルダと同じ場所に置いて実行）
"""
import sys
from pathlib import Path

# ── 1. pillow-heif のインポート確認 ──────────────────────────────────
print("=" * 50)
print("【1】pillow-heif のインポート確認")
try:
    import pillow_heif
    print(f"  OK: pillow_heif バージョン = {pillow_heif.__version__}")
except ImportError as e:
    print(f"  NG: インポート失敗 → {e}")
    sys.exit(1)

# ── 2. register_heif_opener の実行確認 ───────────────────────────────
print()
print("【2】register_heif_opener() の実行確認")
try:
    pillow_heif.register_heif_opener()
    print("  OK: 登録成功")
except Exception as e:
    print(f"  NG: 登録失敗 → {e}")

# ── 3. 写真フォルダから最初の HEIC ファイルを探す ─────────────────────
print()
print("【3】写真フォルダ内の HEIC ファイルを検索")
photo_base = Path(__file__).parent / "写真"
heic_files = list(photo_base.rglob("*.heic")) + list(photo_base.rglob("*.HEIC")) \
           + list(photo_base.rglob("*.heif")) + list(photo_base.rglob("*.HEIF"))

if not heic_files:
    print("  HEIC ファイルが見つかりませんでした")
    print("  ※ 拡張子が .jpg でも実態が HEIC の場合は次の【4】で確認します")
else:
    print(f"  {len(heic_files)} 件見つかりました")
    for f in heic_files[:3]:
        print(f"    {f}")

# ── 4. .jpg ファイルのマジックバイト確認（HEIC偽装チェック）──────────
print()
print("【4】.jpg ファイルのマジックバイト確認（HEIC偽装チェック）")
jpg_files = list(photo_base.rglob("*.jpg")) + list(photo_base.rglob("*.JPG"))
heic_disguised = []
for f in jpg_files:
    try:
        header = f.read_bytes()[:12]
        if header[4:8] == b"ftyp":
            heic_disguised.append((f, header[8:12]))
    except Exception:
        pass

if heic_disguised:
    print(f"  HEIC偽装ファイル {len(heic_disguised)} 件検出:")
    for f, brand in heic_disguised[:5]:
        print(f"    {f.name}  (brand={brand})")
else:
    print("  HEIC偽装なし（全て正規のJPEGです）")

# ── 5. 実際に HEIC ファイルを Image.open() で開いてみる ──────────────
print()
print("【5】HEIC/偽装JPGを Image.open() で開くテスト")
from PIL import Image
targets = heic_files[:1] + [f for f, _ in heic_disguised[:1]]
if not targets:
    print("  テスト対象ファイルなし")
else:
    for target in targets:
        print(f"  テスト: {target.name}")
        try:
            with Image.open(target) as img:
                print(f"    OK: サイズ={img.size}, モード={img.mode}")
        except Exception as e:
            print(f"    NG: {e}")

# ── 6. pillow_heif.open_heif() 直接テスト ────────────────────────────
print()
print("【6】pillow_heif.open_heif() 直接テスト")
if not targets:
    print("  テスト対象ファイルなし")
else:
    for target in targets:
        print(f"  テスト: {target.name}")
        try:
            heif_file = pillow_heif.open_heif(str(target), convert_hdr_to_8bit=True)
            img = Image.frombytes(heif_file.mode, heif_file.size, heif_file.data, "raw")
            print(f"    OK: サイズ={img.size}, モード={img.mode}")
        except Exception as e:
            print(f"    NG: {e}")

print()
print("=" * 50)
print("診断完了。上記の内容を Claude に貼り付けてください。")
