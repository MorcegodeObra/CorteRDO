import fitz
import os
import utils.func_cortar_image as imagens
from tkinter import Tk,filedialog

root = Tk()
root.withdraw()
pdf_path = filedialog.askopenfilename(
    title="Selecione o arquivo PDF",
    filetypes=[("Arquivos PDF","*.pdf")]
)
if not pdf_path:
    raise SystemExit("Nenhum arquivo selecionado.")

pasta_saida = os.path.join(os.path.dirname(pdf_path),"imagens_cortadas")

os.makedirs(pasta_saida, exist_ok=True)

doc = fitz.open(pdf_path)

for i, page in enumerate(doc):
    pix = page.get_pixmap(dpi=300)

    temp_png = os.path.join(
        pasta_saida,
        f"pagina_{i+1:03d}.png"
    )
    pix.save(temp_png)

    # 🔪 reaproveita seu corte
    imagens.cortar_imagem(temp_png)

    print(f"✂️ Página {i+1} processada")

doc.close()
