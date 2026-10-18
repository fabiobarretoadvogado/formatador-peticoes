from __future__ import annotations

import argparse
import copy
import hashlib
import json
import posixpath
import tempfile
import zipfile
from pathlib import Path

from lxml import etree


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}
W = f"{{{NS['w']}}}"
PR = f"{{{NS['pr']}}}"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _xml(data: bytes) -> etree._Element:
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
    return etree.fromstring(data, parser=parser)


def _xml_bytes(root: etree._Element) -> bytes:
    return etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )


def _copy_zip_info(info: zipfile.ZipInfo) -> zipfile.ZipInfo:
    cloned = zipfile.ZipInfo(info.filename, date_time=info.date_time)
    cloned.compress_type = info.compress_type
    cloned.comment = info.comment
    cloned.extra = info.extra
    cloned.create_system = info.create_system
    cloned.create_version = info.create_version
    cloned.extract_version = info.extract_version
    cloned.flag_bits = info.flag_bits
    cloned.volume = info.volume
    cloned.internal_attr = info.internal_attr
    cloned.external_attr = info.external_attr
    return cloned


def _read_package(path: Path) -> tuple[list[zipfile.ZipInfo], dict[str, bytes]]:
    with zipfile.ZipFile(path) as package:
        infos = package.infolist()
        return infos, {info.filename: package.read(info.filename) for info in infos}


def _style_name(style: etree._Element) -> str:
    name = style.find("w:name", namespaces=NS)
    return "" if name is None else name.get(f"{W}val", "")


def _style_id(style: etree._Element) -> str:
    return style.get(f"{W}styleId", "")


def _copy_fb_styles(base_parts: dict[str, bytes], target_parts: dict[str, bytes]) -> list[str]:
    base_styles = _xml(base_parts["word/styles.xml"])
    target_styles = _xml(target_parts["word/styles.xml"])
    selected = [
        style
        for style in base_styles.findall("w:style", namespaces=NS)
        if _style_name(style).startswith("FB - ")
    ]
    if not selected:
        raise RuntimeError("O modelo-base não contém estilos com o prefixo 'FB - '.")

    selected_ids = {_style_id(style) for style in selected}
    for style in list(target_styles.findall("w:style", namespaces=NS)):
        if _style_id(style) in selected_ids or _style_name(style).startswith("FB - "):
            target_styles.remove(style)
    for style in selected:
        target_styles.append(copy.deepcopy(style))

    target_parts["word/styles.xml"] = _xml_bytes(target_styles)
    return sorted(_style_name(style) for style in selected)


def _clear_reference_body(target_parts: dict[str, bytes]) -> None:
    document = _xml(target_parts["word/document.xml"])
    body = document.find("w:body", namespaces=NS)
    if body is None:
        raise RuntimeError("O papel timbrado não contém um corpo Word válido.")
    section = body.find("w:sectPr", namespaces=NS)
    if section is None:
        raise RuntimeError("O papel timbrado não contém propriedades de seção.")

    for child in list(body):
        if child is not section:
            body.remove(child)
    body.insert(0, etree.Element(f"{W}p"))

    page_size = section.find("w:pgSz", namespaces=NS)
    if page_size is None:
        page_size = etree.SubElement(section, f"{W}pgSz")
    page_size.set(f"{W}w", "11906")
    page_size.set(f"{W}h", "16838")
    page_size.attrib.pop(f"{W}orient", None)

    margins = section.find("w:pgMar", namespaces=NS)
    if margins is None:
        margins = etree.SubElement(section, f"{W}pgMar")
    for key, value in {
        "top": "1701",
        "right": "1134",
        "bottom": "1560",
        "left": "1134",
        "header": "709",
        "footer": "856",
        "gutter": "0",
    }.items():
        margins.set(f"{W}{key}", value)

    target_parts["word/document.xml"] = _xml_bytes(document)


def _remove_unused_document_images(target_parts: dict[str, bytes]) -> None:
    rels_name = "word/_rels/document.xml.rels"
    rels = _xml(target_parts[rels_name])
    document = _xml(target_parts["word/document.xml"])
    used_ids = set(document.xpath("//@r:id | //@r:embed | //@r:link", namespaces=NS))
    for relationship in list(rels):
        rel_type = relationship.get("Type", "")
        if rel_type.endswith("/image") and relationship.get("Id") not in used_ids:
            rels.remove(relationship)
    target_parts[rels_name] = _xml_bytes(rels)


def _source_part_for_rels(rels_name: str) -> str:
    if rels_name == "_rels/.rels":
        return ""
    folder, rels_file = posixpath.split(rels_name)
    if not folder.endswith("/_rels") or not rels_file.endswith(".rels"):
        return ""
    source_folder = folder[: -len("/_rels")]
    source_file = rels_file[: -len(".rels")]
    return posixpath.join(source_folder, source_file)


def _remove_unreferenced_media(target_parts: dict[str, bytes]) -> list[str]:
    referenced: set[str] = set()
    for name, data in target_parts.items():
        if not name.endswith(".rels"):
            continue
        source_part = _source_part_for_rels(name)
        source_folder = posixpath.dirname(source_part)
        rels = _xml(data)
        for relationship in rels.findall(f"{PR}Relationship"):
            if relationship.get("TargetMode") == "External":
                continue
            target = relationship.get("Target", "")
            resolved = posixpath.normpath(posixpath.join(source_folder, target))
            if resolved.startswith("word/media/"):
                referenced.add(resolved)

    removed: list[str] = []
    for name in list(target_parts):
        if name.startswith("word/media/") and name not in referenced:
            removed.append(name)
            del target_parts[name]
    return sorted(removed)


def _enable_field_updates(target_parts: dict[str, bytes]) -> None:
    settings = _xml(target_parts["word/settings.xml"])
    update = settings.find("w:updateFields", namespaces=NS)
    if update is None:
        update = etree.SubElement(settings, f"{W}updateFields")
    update.set(f"{W}val", "true")
    target_parts["word/settings.xml"] = _xml_bytes(settings)


def _write_package(
    output: Path,
    reference_infos: list[zipfile.ZipInfo],
    target_parts: dict[str, bytes],
) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        prefix=f"{output.stem}_",
        suffix=".docx",
        dir=output.parent,
        delete=False,
    ) as handle:
        temporary = Path(handle.name)
    try:
        known = set()
        with zipfile.ZipFile(temporary, "w") as package:
            for info in reference_infos:
                if info.filename not in target_parts:
                    continue
                known.add(info.filename)
                package.writestr(_copy_zip_info(info), target_parts[info.filename])
            for name in sorted(set(target_parts) - known):
                package.writestr(name, target_parts[name], compress_type=zipfile.ZIP_DEFLATED)
        temporary.replace(output)
    finally:
        temporary.unlink(missing_ok=True)


def _canonical(element: etree._Element) -> bytes:
    return etree.tostring(element, method="c14n", exclusive=True, with_comments=False)


def _validate(
    output: Path,
    base_parts: dict[str, bytes],
    reference_parts: dict[str, bytes],
    style_names: list[str],
    removed_media: list[str],
) -> dict[str, object]:
    _, parts = _read_package(output)
    document = _xml(parts["word/document.xml"])
    section = document.find("w:body/w:sectPr", namespaces=NS)
    if section is None:
        raise RuntimeError("O modelo final ficou sem propriedades de seção.")

    page_size = section.find("w:pgSz", namespaces=NS)
    margins = section.find("w:pgMar", namespaces=NS)
    if page_size is None or margins is None:
        raise RuntimeError("O modelo final ficou sem tamanho de página ou margens.")
    expected_page = {"w": "11906", "h": "16838"}
    actual_page = {key: page_size.get(f"{W}{key}") for key in expected_page}
    if actual_page != expected_page:
        raise RuntimeError(f"Tamanho de página inesperado: {actual_page}")
    expected_margins = {
        "top": "1701",
        "right": "1134",
        "bottom": "1560",
        "left": "1134",
        "header": "709",
        "footer": "856",
        "gutter": "0",
    }
    actual_margins = {key: margins.get(f"{W}{key}") for key in expected_margins}
    if actual_margins != expected_margins:
        raise RuntimeError(f"Margens inesperadas: {actual_margins}")

    header_refs = section.findall("w:headerReference", namespaces=NS)
    footer_refs = section.findall("w:footerReference", namespaces=NS)
    if not header_refs or not footer_refs:
        raise RuntimeError("O modelo final não referencia cabeçalho e rodapé.")

    header = _xml(parts["word/header1.xml"])
    anchors = header.findall(".//wp:anchor", namespaces=NS)
    large_backgrounds = []
    for anchor in anchors:
        extent = anchor.find("wp:extent", namespaces=NS)
        if extent is None:
            continue
        if (
            anchor.get("behindDoc") == "1"
            and int(extent.get("cx", "0")) >= 7_500_000
            and int(extent.get("cy", "0")) >= 9_500_000
        ):
            large_backgrounds.append(anchor)
    if len(large_backgrounds) != 1:
        raise RuntimeError("Não foi encontrada exatamente uma imagem de fundo A4 no cabeçalho.")

    footer = _xml(parts["word/footer1.xml"])
    instructions = " ".join(
        text.strip()
        for text in footer.xpath(".//w:instrText/text()", namespaces=NS)
        if text.strip()
    )
    footer_text = "".join(footer.xpath(".//w:t/text()", namespaces=NS))
    if "PAGE" not in instructions or "NUMPAGES" not in instructions or "Página" not in footer_text:
        raise RuntimeError("O rodapé não contém a paginação 'Página X de Y'.")

    if document.findall(".//w:drawing", namespaces=NS):
        raise RuntimeError("O corpo vazio do modelo ainda contém imagem ou assinatura.")

    base_styles = _xml(base_parts["word/styles.xml"])
    final_styles = _xml(parts["word/styles.xml"])
    base_by_name = {
        _style_name(style): _canonical(style)
        for style in base_styles.findall("w:style", namespaces=NS)
        if _style_name(style).startswith("FB - ")
    }
    final_by_name = {
        _style_name(style): _canonical(style)
        for style in final_styles.findall("w:style", namespaces=NS)
        if _style_name(style).startswith("FB - ")
    }
    if set(final_by_name) != set(base_by_name):
        raise RuntimeError("A lista de estilos FB do modelo final diverge do modelo-base.")
    if any(final_by_name[name] != base_by_name[name] for name in base_by_name):
        raise RuntimeError("Uma ou mais definições de estilos FB foram alteradas.")

    if parts["word/header1.xml"] != reference_parts["word/header1.xml"]:
        raise RuntimeError("O cabeçalho original do papel timbrado foi modificado.")
    if parts["word/footer1.xml"] != reference_parts["word/footer1.xml"]:
        raise RuntimeError("O rodapé original do papel timbrado foi modificado.")

    media = sorted(name for name in parts if name.startswith("word/media/"))
    if len(media) != 1:
        raise RuntimeError(f"Esperava-se apenas a imagem do timbre; encontradas: {media}")

    settings = _xml(parts["word/settings.xml"])
    update = settings.find("w:updateFields", namespaces=NS)
    if update is None or update.get(f"{W}val") not in {"1", "true", "on"}:
        raise RuntimeError("A atualização automática de campos não foi habilitada.")

    combined_style_hash = hashlib.sha256(
        b"\n".join(base_by_name[name] for name in sorted(base_by_name))
    ).hexdigest()
    return {
        "output": str(output.resolve()),
        "output_sha256": _sha256(output),
        "page": "A4 portrait",
        "margins_cm": {"top": 3.0, "bottom": 2.75, "left": 2.0, "right": 2.0, "gutter": 0.0},
        "header_cm": 1.25,
        "footer_cm": 1.51,
        "background_media": media,
        "pagination": "Página X de Y",
        "styles_fb_count": len(style_names),
        "styles_fb_sha256": combined_style_hash,
        "removed_unreferenced_media": removed_media,
        "package_parts": sorted(parts),
    }


def build(base: Path, reference: Path, output: Path) -> dict[str, object]:
    if base.resolve() == reference.resolve():
        raise ValueError("O modelo-base e o papel timbrado devem ser arquivos diferentes.")
    reference_infos, reference_parts = _read_package(reference)
    _, base_parts = _read_package(base)
    target_parts = dict(reference_parts)

    style_names = _copy_fb_styles(base_parts, target_parts)
    _clear_reference_body(target_parts)
    _remove_unused_document_images(target_parts)
    removed_media = _remove_unreferenced_media(target_parts)
    _enable_field_updates(target_parts)
    _write_package(output, reference_infos, target_parts)
    return _validate(output, base_parts, reference_parts, style_names, removed_media)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Monta o modelo oficial vazio com timbre, paginação e estilos FB preservados."
    )
    parser.add_argument("--modelo-base", required=True, type=Path)
    parser.add_argument("--papel-timbrado", required=True, type=Path)
    parser.add_argument("--saida", required=True, type=Path)
    parser.add_argument("--manifesto", type=Path)
    args = parser.parse_args()

    result = build(args.modelo_base, args.papel_timbrado, args.saida)
    result["base_model"] = str(args.modelo_base.resolve())
    result["base_model_sha256"] = _sha256(args.modelo_base)
    result["paperhead_reference"] = str(args.papel_timbrado.resolve())
    result["paperhead_reference_sha256"] = _sha256(args.papel_timbrado)
    if args.manifesto:
        args.manifesto.parent.mkdir(parents=True, exist_ok=True)
        args.manifesto.write_text(
            json.dumps(result, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
