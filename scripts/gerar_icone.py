"""Ícone vetorial simples rasterizado para o executável e o instalador."""
from pathlib import Path
from PIL import Image, ImageDraw


def main() -> None:
    destino = Path(__file__).resolve().parents[1] / "assets" / "aplicativo.ico"
    destino.parent.mkdir(exist_ok=True)
    imagem = Image.new("RGBA", (256, 256), (0, 0, 0, 0))
    desenho = ImageDraw.Draw(imagem)
    desenho.rounded_rectangle((8, 8, 248, 248), radius=48, fill="#2563EB")
    desenho.rounded_rectangle((72, 45, 186, 211), radius=10, fill="white")
    for y, largura in ((88, 70), (116, 70), (144, 70), (172, 44)):
        desenho.rounded_rectangle((94, y, 94 + largura, y + 9), radius=4, fill="#2563EB")
    imagem.save(destino, sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])


if __name__ == "__main__":
    main()
