import unicodedata
import re
import cv2
import os

def normalizar(texto):
    nome = unicodedata.normalize("NFD", texto)
    nome = "".join(c for c in nome if unicodedata.category(c) != "Mn")
    nome = re.sub(r"[^A-Za-z0-9._-]", "_", nome)
    return nome.upper()


def encontrar_aba_ocorrencias(wb):
    for ws in wb.Worksheets:
        nome_original = ws.Name
        nome = normalizar(nome_original)

        if (
            "OCORR" in nome or
            "OCORREN" in nome or
            "EVENTO" in nome or
            "APONT" in nome
        ):
            return ws

    return None

def salvar_debug(img, top, bottom, left, right, img_path):
    debug = img.copy()

    cv2.rectangle(
        debug,
        (left, top),
        (right, bottom),
        (0, 0, 255),
        5
    )

    base, ext = os.path.splitext(img_path)
    debug_path = base + "_DEBUG" + ext

    # 🔐 Escrita segura Unicode
    success, encoded = cv2.imencode(ext, debug)
    if not success:
        print("❌ Falha ao codificar imagem de debug")
        return False

    with open(debug_path, "wb") as f:
        f.write(encoded)

    print(f"🧪 Debug salvo (unicode-safe) em: {debug_path}")
    return True



