import os
import shutil
import tempfile
import fitz
import utils.func_extras as extras

def exportar_aba_para_imagem(excel, xlsx_path, output_img):
    temp_dir = tempfile.mkdtemp(prefix="rdo_")
    temp_xlsx = os.path.join(temp_dir, "arquivo.xlsx")
    temp_pdf = os.path.join(temp_dir, "saida.pdf")

    shutil.copy2(xlsx_path, temp_xlsx)

    wb = excel.Workbooks.Open(
        temp_xlsx,
        ReadOnly=True,
        UpdateLinks=0,
        IgnoreReadOnlyRecommended=True,
        AddToMru=False
    )

    if wb is None:
        print("❌ Excel não conseguiu abrir o arquivo")
        shutil.rmtree(temp_dir, ignore_errors=True)
        return False

    # 🔎 ENCONTRAR ABA CORRETA
    ws = extras.encontrar_aba_ocorrencias(wb)
    if ws is None:
        wb.Close(False)
        shutil.rmtree(temp_dir, ignore_errors=True)
        return False

    # 🖨️ CONFIGURAÇÃO DE IMPRESSÃO (NA ABA!)
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False

    ws.Activate()

    # 📄 EXPORTAR A ABA (não o workbook)
    ws.ExportAsFixedFormat(
        Type=0,  # PDF
        Filename=temp_pdf,
        Quality=0,
        IncludeDocProperties=True,
        IgnorePrintAreas=False,
        OpenAfterPublish=False
    )

    wb.Close(False)

    # PDF → PNG
    temp_img_dir = tempfile.mkdtemp(prefix="img_")
    temp_png = os.path.join(temp_img_dir, "saida.png")

    doc = fitz.open(temp_pdf)
    page = doc.load_page(0)
    pix = page.get_pixmap(dpi=300)
    pix.save(temp_png)
    doc.close()

    nome_final = extras.normalizar(os.path.basename(output_img))
    saida_final = os.path.join(os.path.dirname(output_img), nome_final)
    shutil.copy2(temp_png, saida_final)

    shutil.rmtree(temp_img_dir, ignore_errors=True)
    shutil.rmtree(temp_dir, ignore_errors=True)

    return saida_final