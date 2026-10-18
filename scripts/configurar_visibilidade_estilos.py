from __future__ import annotations

import argparse
import json
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Iterable

from lxml import etree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


# CT_Style has a fixed child order in WordprocessingML. ElementTree preserves
# insertion order, so newly added visibility elements must be placed according
# to this sequence rather than simply appended to the style element.
_STYLE_CHILD_ORDER = {
    name: index
    for index, name in enumerate(
        (
            "name",
            "aliases",
            "basedOn",
            "next",
            "link",
            "autoRedefine",
            "hidden",
            "uiPriority",
            "semiHidden",
            "unhideWhenUsed",
            "qFormat",
            "locked",
            "personal",
            "personalCompose",
            "personalReply",
            "rsid",
            "pPr",
            "rPr",
            "tblPr",
            "trPr",
            "tcPr",
            "tblStylePr",
        )
    )
}


def _qn(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def _child(parent: ET.Element, tag: str) -> ET.Element | None:
    return parent.find(_qn(tag))


def _value(parent: ET.Element, tag: str) -> str:
    element = _child(parent, tag)
    return element.get(_qn("val"), "") if element is not None else ""


def _set_on_off(parent: ET.Element, tag: str, enabled: bool) -> None:
    element = _child(parent, tag)
    if enabled:
        if element is None:
            element = ET.Element(_qn(tag))
            parent.append(element)
        element.set(_qn("val"), "1")
    elif element is not None:
        parent.remove(element)


def _set_value(parent: ET.Element, tag: str, value: str) -> None:
    element = _child(parent, tag)
    if element is None:
        element = ET.Element(_qn(tag))
        parent.append(element)
    element.set(_qn("val"), value)


def _set_attribute(element: ET.Element, name: str, value: str) -> None:
    element.set(_qn(name), value)


def _set_setting(settings: ET.Element, tag: str, attributes: dict[str, str]) -> None:
    element = _child(settings, tag)
    if tag == "stylePaneFormatFilter":
        if element is None:
            element = ET.Element(_qn(tag))
        else:
            settings.remove(element)
        # In CT_Settings, stylePaneFormatFilter comes immediately after
        # proofState and before defaultTabStop. Keep this insertion valid even
        # when the source document did not already contain the filter.
        proof_state = _child(settings, "proofState")
        if proof_state is not None:
            settings.insert(list(settings).index(proof_state) + 1, element)
        else:
            default_tab_stop = _child(settings, "defaultTabStop")
            if default_tab_stop is not None:
                settings.insert(list(settings).index(default_tab_stop), element)
            else:
                settings.insert(0, element)
    elif element is None:
        element = ET.Element(_qn(tag))
        settings.append(element)
    for name, value in attributes.items():
        element.set(_qn(name), value)


def _reorder_style_children(style: ET.Element) -> None:
    children = list(style)
    ordered = sorted(
        enumerate(children),
        key=lambda item: (_STYLE_CHILD_ORDER.get(item[1].tag.split("}")[-1], 10_000), item[0]),
    )
    if [child for _, child in ordered] == children:
        return
    for child in children:
        style.remove(child)
    for _, child in ordered:
        style.append(child)


def _style_name(style: ET.Element) -> str:
    return _value(style, "name")


def _configure_styles(styles_xml: bytes, custom_names: Iterable[str]) -> bytes:
    custom_names = list(custom_names)
    custom_name_set = set(custom_names)
    root = ET.fromstring(styles_xml)
    styles = root.findall(_qn("style"))

    for style in styles:
        name = _style_name(style)
        is_custom = name in custom_name_set

        # Keep the authored FB styles visible and ordered in the style UI.
        # Built-in styles remain available to Word internally, but are hidden
        # from the user-facing style lists instead of being deleted.
        _set_on_off(style, "semiHidden", not is_custom)
        _set_on_off(style, "unhideWhenUsed", False)
        _set_on_off(style, "hidden", False)

        if is_custom:
            priority = custom_names.index(name) + 1
            _set_value(style, "uiPriority", str(priority))
            _set_on_off(style, "qFormat", True)
        else:
            _set_on_off(style, "qFormat", False)

        _reorder_style_children(style)

    latent = _child(root, "latentStyles")
    if latent is not None:
        latent.set(_qn("defSemiHidden"), "1")
        latent.set(_qn("defUnhideWhenUsed"), "0")
        latent.set(_qn("defQFormat"), "0")
        for exception in latent.findall(_qn("lsdException")):
            _set_attribute(exception, "semiHidden", "1")
            _set_attribute(exception, "unhideWhenUsed", "0")
            _set_attribute(exception, "qFormat", "0")

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _configure_settings(settings_xml: bytes) -> bytes:
    root = ET.fromstring(settings_xml)
    # The style pane filter is an OR filter. Selecting customStyles makes the
    # document request only explicitly authored styles, i.e. the FB styles.
    _set_setting(root, "stylePaneFormatFilter", {"customStyles": "1"})
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def ajustar_arquivo(origem: Path, destino: Path, nomes_customizados: Iterable[str]) -> Path:
    """Write a copy of a DOCX with only authored styles exposed in Word."""
    origem = origem.resolve()
    destino = destino.resolve()
    destino.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.NamedTemporaryFile(
        prefix=f"{destino.stem}_", suffix=".docx", dir=destino.parent, delete=False
    ) as temp:
        temporario = Path(temp.name)

    try:
        with zipfile.ZipFile(origem, "r") as source, zipfile.ZipFile(temporario, "w") as output:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename == "word/styles.xml":
                    data = _configure_styles(data, nomes_customizados)
                elif info.filename == "word/settings.xml":
                    data = _configure_settings(data)
                output.writestr(info, data)
        shutil.move(str(temporario), str(destino))
    finally:
        if temporario.exists():
            temporario.unlink()

    return destino


def ajustar_espacamento_estilo(origem: Path, destino: Path, nome_estilo: str, antes_pt: float) -> Path:
    """Patch only a paragraph style's before-spacing while preserving the DOCX package."""
    origem = origem.resolve()
    destino = destino.resolve()
    destino.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        prefix=f"{destino.stem}_", suffix=".docx", dir=destino.parent, delete=False
    ) as temp:
        temporario = Path(temp.name)

    encontrado = False
    try:
        with zipfile.ZipFile(origem, "r") as source, zipfile.ZipFile(temporario, "w") as output:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename == "word/styles.xml":
                    root = ET.fromstring(data)
                    for style in root.findall(_qn("style")):
                        if _style_name(style) != nome_estilo:
                            continue
                        ppr = _child(style, "pPr")
                        if ppr is None:
                            ppr = ET.Element(_qn("pPr"))
                            style.append(ppr)
                        spacing = _child(ppr, "spacing")
                        if spacing is None:
                            spacing = ET.Element(_qn("spacing"))
                            ppr.append(spacing)
                        spacing.set(_qn("before"), str(round(antes_pt * 20)))
                        encontrado = True
                        break
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                output.writestr(info, data)
        if not encontrado:
            raise ValueError(f"Estilo não encontrado no modelo: {nome_estilo}")
        temporario.replace(destino)
    finally:
        if temporario.exists():
            temporario.unlink()

    return destino


def ajustar_negrito_estilo(origem: Path, destino: Path, nome_estilo: str, negrito: bool) -> Path:
    """Patch a paragraph style's bold setting while preserving the DOCX package."""
    origem = origem.resolve()
    destino = destino.resolve()
    destino.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        prefix=f"{destino.stem}_", suffix=".docx", dir=destino.parent, delete=False
    ) as temp:
        temporario = Path(temp.name)

    encontrado = False
    try:
        with zipfile.ZipFile(origem, "r") as source, zipfile.ZipFile(temporario, "w") as output:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename == "word/styles.xml":
                    root = ET.fromstring(data)
                    for style in root.findall(_qn("style")):
                        if _style_name(style) != nome_estilo:
                            continue
                        rpr = _child(style, "rPr")
                        if rpr is None:
                            rpr = ET.Element(_qn("rPr"))
                            style.append(rpr)
                        for tag in ("b", "bCs"):
                            existing = _child(rpr, tag)
                            if existing is not None:
                                rpr.remove(existing)
                        if negrito:
                            bold = ET.Element(_qn("b"))
                            insertion = next(
                                (index for index, child in enumerate(rpr) if child.tag.rsplit("}", 1)[-1] in {"sz", "szCs"}),
                                len(rpr),
                            )
                            rpr.insert(insertion, bold)
                        encontrado = True
                        break
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                output.writestr(info, data)
        if not encontrado:
            raise ValueError(f"Estilo não encontrado no modelo: {nome_estilo}")
        temporario.replace(destino)
    finally:
        if temporario.exists():
            temporario.unlink()

    return destino


def _argumentos() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Exibe apenas os estilos autorais no painel de estilos do Word."
    )
    parser.add_argument("origem", type=Path)
    parser.add_argument("--out", required=True, type=Path, dest="destino")
    parser.add_argument(
        "--estilos-json",
        type=Path,
        default=None,
        help="JSON com a chave estilos_permitidos; por padrão, usa o arquivo ao lado do DOCX.",
    )
    return parser.parse_args()


def main() -> None:
    args = _argumentos()
    estilos_json = args.estilos_json or args.origem.parent / "estilos.json"
    config = json.loads(estilos_json.read_text(encoding="utf-8"))
    ajustar_arquivo(args.origem, args.destino, config["estilos_permitidos"])
    print(args.destino)


if __name__ == "__main__":
    main()
