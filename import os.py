import os
from docx import Document
from docx.shared import Mm

# ファイルパスの設定
docx_path = '防火ポスター入賞作品.docx'
photo_folder = '写真'
output_path = '防火ポスター入賞作品_写真あり.docx'

def insert_photos():
    if not os.path.exists(docx_path):
        print("Wordファイルが見つかりません。")
        return

    doc = Document(docx_path)
    
    # 全てのテーブルをループして枠を確認
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                # ここに挿入条件を記述（例：特定の名前があるセルの隣など）
                # 今回は仮に「青山 了平」さんの枠を特定する場合
                if "青山　了平" in cell.text:
                    # 同じテーブル内の画像を入れるべきセルを特定
                    # 例：名前の上のセルに写真を入れるなど
                    photo_path = os.path.join(photo_folder, 'aoyama_poster.jpg')
                    
                    if os.path.exists(photo_path):
                        paragraph = cell.paragraphs[0]
                        run = paragraph.add_run()
                        # 高さを110mmに指定して挿入
                        run.add_picture(photo_path, height=Mm(110))
                        print(f"{photo_path} を挿入しました。")

    doc.save(output_path)
    print(f"保存完了: {output_path}")

if __name__ == "__main__":
    insert_photos()