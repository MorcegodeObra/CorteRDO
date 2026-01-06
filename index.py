import os
import win32com.client
from tkinter import Tk, filedialog
import utils.func_exportar_aba as export
import utils.func_cortar_image as corta

# ===============================
# SELETOR DE PASTA
# ===============================
root = Tk()
root.withdraw()
PASTA_PLANILHAS = filedialog.askdirectory(title="Selecione a pasta com os arquivos Excel")

if not PASTA_PLANILHAS:
    raise SystemExit("Nenhuma pasta selecionada.")

PASTA_IMAGENS = os.path.join(PASTA_PLANILHAS, "imagens")
os.makedirs(PASTA_IMAGENS, exist_ok=True)

# ===============================
# EXCEL (INVISÍVEL)
# ===============================
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

# ===============================
# PROCESSAMENTO
# ===============================
print(f"Rodando em {os.path.basename(PASTA_PLANILHAS)}...")

for arquivo in os.listdir(PASTA_PLANILHAS):
    if not arquivo.lower().endswith(".xlsx"):
        continue

    caminho = os.path.join(PASTA_PLANILHAS, arquivo)
    nome_img = os.path.splitext(arquivo)[0] + ".png"
    saida = os.path.join(PASTA_IMAGENS, nome_img)
    saida_final = export.exportar_aba_para_imagem(excel, caminho, saida)
    if saida_final:
        corta.cortar_imagem(saida_final)

excel.Quit()
print("✅ Finalizado.")
