"""Dessine l'icone du programme : un bouclier bleu nuit portant une balise [ ].
Produit icone.png (256 px, pour la fenetre) et icone.ico (multi-tailles, pour
l'executable et l'Explorateur Windows). Dessin original, sans police externe."""

import pathlib
from PIL import Image, ImageDraw

ICI = pathlib.Path(__file__).parent
BLEU = (27, 58, 92, 255)
BLEU_CLAIR = (46, 96, 146, 255)
BLANC = (255, 255, 255, 255)
TEAL = (88, 196, 186, 255)


def bouclier(taille):
    s = 8                                    # sur-echantillonnage pour l'anticrenelage
    T = taille * s
    img = Image.new("RGBA", (T, T), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    m = T * 0.06
    w, h = T - 2 * m, T - 2 * m
    # Contour du bouclier : haut droit, flancs, pointe en bas (polygone lisse)
    pts = []
    import math
    cx = T / 2
    haut = m
    pts.append((m, haut))
    pts.append((T - m, haut))
    pts.append((T - m, haut + h * 0.55))
    for i in range(0, 91, 3):
        a = math.radians(i)
        pts.append((T - m - (w / 2) * (1 - math.cos(a)), haut + h * 0.55 + (h * 0.45) * math.sin(a)))
    for i in range(90, 181, 3):
        a = math.radians(i)
        pts.append((cx + (w / 2) * math.cos(a), haut + h * 0.55 + (h * 0.45) * math.sin(a)))
    pts.append((m, haut + h * 0.55))
    d.polygon(pts, fill=BLEU)
    # Bande claire en haut (relief discret)
    d.rectangle([m, haut, T - m, haut + h * 0.14], fill=BLEU_CLAIR)
    # Les crochets de la balise : [  ]
    ep = T * 0.065
    x1, x2 = T * 0.27, T * 0.73
    y1, y2 = T * 0.30, T * 0.70
    larg = T * 0.11
    for x, sens in ((x1, 1), (x2, -1)):
        d.rectangle([x if sens == 1 else x - ep, y1, x + ep if sens == 1 else x, y2], fill=BLANC)
        d.rectangle([x if sens == 1 else x - larg, y1, x + larg if sens == 1 else x, y1 + ep], fill=BLANC)
        d.rectangle([x if sens == 1 else x - larg, y2 - ep, x + larg if sens == 1 else x, y2], fill=BLANC)
    # Un point couleur au centre : la donnee protegee
    r = T * 0.055
    d.ellipse([cx - r, T * 0.5 - r, cx + r, T * 0.5 + r], fill=TEAL)
    return img.resize((taille, taille), Image.LANCZOS)


def main():
    grande = bouclier(256)
    grande.save(ICI / "icone.png")
    tailles = [16, 24, 32, 48, 64, 128, 256]
    images = [bouclier(t) for t in tailles]
    images[-1].save(ICI / "icone.ico", format="ICO", sizes=[(t, t) for t in tailles],
                    append_images=images[:-1])
    print("icone.png et icone.ico ecrits dans", ICI)


if __name__ == "__main__":
    main()
