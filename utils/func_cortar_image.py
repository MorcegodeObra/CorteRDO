import os
import cv2
import numpy as np

def cortar_imagem(img_path):
    if not os.path.exists(img_path):
        print(f"❌ Imagem não encontrada: {img_path}")
        return False

    # leitura segura (unicode)
    with open(img_path, "rb") as f:
        data = np.frombuffer(f.read(), np.uint8)
    img = cv2.imdecode(data, cv2.IMREAD_COLOR)

    if img is None:
        print(f"❌ OpenCV não conseguiu abrir a imagem: {img_path}")
        return False

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # detecta linhas escuras
    _, bin_img = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY_INV)

    h, w = bin_img.shape

    proj_h = np.sum(bin_img, axis=1)
    proj_v = np.sum(bin_img, axis=0)

    try:
        top = next(i for i in range(h) if proj_h[i] > 0)
        bottom = next(i for i in range(h - 1, -1, -1) if proj_h[i] > 0)
        left = next(i for i in range(w) if proj_v[i] > 0)
        right = next(i for i in range(w - 1, -1, -1) if proj_v[i] > 0)
    except StopIteration:
        return False

    margem = 3
    top = max(top + margem, 0)
    left = max(left + margem, 0)
    bottom = min(bottom - margem, h)
    right = min(right - margem, w)

    img_cortada = img[top:bottom, left:right]

    if img_cortada.size == 0:
        print("❌ Corte resultou em imagem vazia")
        return False

    # 🔐 SALVAMENTO SEGURO (UNICODE SAFE)
    ext = os.path.splitext(img_path)[1]
    success, encoded = cv2.imencode(ext, img_cortada)

    if not success:
        print("❌ Falha ao codificar imagem cortada")
        return False

    with open(img_path, "wb") as f:
        f.write(encoded)

    return True
