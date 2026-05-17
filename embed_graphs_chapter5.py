# =============================================================================
# embed_graphs_chapter5.py  —  Вбудовування згенерованих графіків у docx
# =============================================================================
# Запускати після gen_graphs_chapter5.py (всі 12 місяців).
#
# Використання:
#   python embed_graphs_chapter5.py
#
# Читає PNG з IMG_DIR, замінює зображення в DOCX_SRC, зберігає як DOCX_DST.
# Залежності: pip install openpyxl
# =============================================================================

import os, io, zipfile

# ── налаштуйте шляхи ──────────────────────────────────────────────────────────
IMG_DIR  = r"E:\КЛОД\Розіл4\graphs"          # папка з PNG (змініть за потреби)
DOCX_SRC = r"E:\КЛОД\Розіл4\РОЗДІЛ5_ ФІНАЛ_fixed.docx"
DOCX_DST = r"E:\КЛОД\Розіл4\РОЗДІЛ5_ ФІНАЛ_graphs.docx"

# ── відповідність: номер рисунку → ім'я файлу зображення в docx ──────────────
FIG_IMG = {
    1:'image17.png', 2:'image2.png',  3:'image5.png',  4:'image26.png',
    5:'image11.png', 6:'image27.png', 7:'image22.png', 8:'image25.png',
    9:'image24.png',10:'image36.png',11:'image21.png',12:'image31.png',
   13:'image23.png',14:'image34.png',15:'image29.png',16:'image28.png',
   17:'image35.png',18:'image32.png',19:'image30.png',20:'image33.png',
   21:'image20.png',22:'image4.png', 23:'image3.png', 24:'image7.png',
   25:'image9.png', 26:'image16.png',27:'image10.png',28:'image13.png',
   29:'image18.png',30:'image19.png',31:'image8.png', 32:'image6.png',
   33:'image14.png',34:'image1.png', 35:'image12.png',36:'image15.png',
}

# Load generated PNGs
new_images = {}
for fn in os.listdir(IMG_DIR):
    if fn.endswith('.png'):
        with open(os.path.join(IMG_DIR, fn), 'rb') as f:
            new_images[fn] = f.read()

print(f"Завантажено {len(new_images)} PNG з {IMG_DIR}")

buf_docx = io.BytesIO()
replaced = 0

with zipfile.ZipFile(DOCX_SRC, 'r') as zin, \
     zipfile.ZipFile(buf_docx, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        if item.filename.startswith('word/media/'):
            img_name = item.filename.split('/')[-1]
            if img_name in new_images:
                zout.writestr(item, new_images[img_name])
                replaced += 1
                print(f"  замінено {img_name}")
                continue
        zout.writestr(item, zin.read(item.filename))

with open(DOCX_DST, 'wb') as f:
    f.write(buf_docx.getvalue())

print(f"\nЗамінено {replaced}/36 зображень.")
print(f"Збережено: {DOCX_DST}")
