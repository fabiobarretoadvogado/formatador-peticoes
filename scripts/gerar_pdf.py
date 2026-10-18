from __future__ import annotations

import argparse
import os
import subprocess
from pathlib import Path


def gerar_pdf(docx: Path, pdf: Path) -> Path:
    docx, pdf = docx.resolve(), pdf.resolve()
    pdf.parent.mkdir(parents=True, exist_ok=True)
    if pdf.exists():
        pdf.unlink()
    try:
        import win32com.client  # type: ignore
    except ImportError:
        script = r"""
$Docx = $env:FORMATADOR_DOCX
$Pdf = $env:FORMATADOR_PDF
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Open($Docx)
  $doc.Repaginate()
  $doc.Fields.Update() | Out-Null
  foreach ($section in $doc.Sections) {
    foreach ($header in $section.Headers) { $header.Range.Fields.Update() | Out-Null }
    foreach ($footer in $section.Footers) { $footer.Range.Fields.Update() | Out-Null }
  }
  $doc.ExportAsFixedFormat($Pdf, 17)
  $doc.Close($false)
} finally { if ($null -ne $word) { $word.Quit() } }
"""
        ambiente = os.environ.copy()
        ambiente.update({"FORMATADOR_DOCX": str(docx), "FORMATADOR_PDF": str(pdf)})
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-STA", "-Command", script],
            capture_output=True,
            text=True,
            env=ambiente,
        )
        if proc.returncode != 0:
            raise RuntimeError("Falha ao exportar pelo Microsoft Word: " + (proc.stderr.strip() or proc.stdout.strip()))
        if not pdf.exists() or pdf.stat().st_size == 0: raise RuntimeError("O Word não produziu um PDF válido")
        return pdf
    word = None
    documento = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        documento = word.Documents.Open(str(docx), ReadOnly=True, AddToRecentFiles=False)
        documento.Repaginate()
        for campo in documento.Fields:
            try: campo.Update()
            except Exception: pass
        for secao in documento.Sections:
            for cabecalho in secao.Headers:
                try: cabecalho.Range.Fields.Update()
                except Exception: pass
            for rodape in secao.Footers:
                try: rodape.Range.Fields.Update()
                except Exception: pass
        documento.ExportAsFixedFormat(str(pdf), 17, OpenAfterExport=False, OptimizeFor=0, Range=0, Item=0, IncludeDocProps=True, KeepIRM=True, CreateBookmarks=1, DocStructureTags=True, BitmapMissingFonts=True, UseISO19005_1=False)
    finally:
        if documento is not None: documento.Close(False)
        if word is not None: word.Quit()
    if not pdf.exists() or pdf.stat().st_size == 0:
        raise RuntimeError("O Word não produziu um PDF válido")
    return pdf


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("docx", type=Path); ap.add_argument("pdf", type=Path)
    args = ap.parse_args(); print(gerar_pdf(args.docx, args.pdf))


if __name__ == "__main__": main()
