"""Open the bundled synthetic files with an independent office suite.

Run after npm ci. Requires Python 3.9+, LibreOffice, pdfinfo and pdftoppm.
All documents, profiles and previews created by this check are temporary.
"""
import os
import re
import subprocess
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
SOFFICE = "soffice.com" if os.name == "nt" else "soffice"
EXAMPLES = ROOT / "public/examples"
NS = {prefix: f"urn:oasis:names:tc:opendocument:xmlns:{prefix}:1.0"
      for prefix in ("chart", "table", "text", "office")}
NS["draw"] = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"


def run(*args):
    return subprocess.run(args, cwd=ROOT, check=True, capture_output=True, timeout=120)


def check_presentation(archive, xml):
    slides = xml.findall(".//draw:page", NS)
    assert len(slides) == 2, "Expected two editable slides"
    for slide, expected in zip(slides, ("跨系统，保留可编辑内容", "图表也应该继续可编辑")):
        paragraphs = ["".join(p.itertext()) for p in slide.findall(".//text:p", NS)]
        assert expected in paragraphs, ("Missing editable title", expected, paragraphs)
    obj = slides[1].find(".//draw:object", NS)
    assert obj is not None, "Missing embedded object"
    href = obj.attrib["{http://www.w3.org/1999/xlink}href"].removeprefix("./")
    chart = ET.fromstring(archive.read(href + "/content.xml"))
    assert chart.find(".//chart:chart", NS) is not None, "Chart must be recognized as a chart"


def check_workbook(xml):
    sheets = xml.findall(".//office:spreadsheet/table:table", NS)
    assert [s.get("{%s}name" % NS["table"]) for s in sheets] == ["预算示例", "编号与说明"]
    results = {c.get("{%s}formula" % NS["table"]): c.get("{%s}value" % NS["office"])
               for c in sheets[0].findall(".//table:table-cell", NS)}
    assert results.get("of:=SUM([.C2:.C4])") == "4750", ("Expected recalculated formula", results)
    identifiers = {"".join(c.itertext()) for c in sheets[1].findall(".//table:table-cell", NS)
                   if c.get("{%s}value-type" % NS["office"]) == "string"}
    assert {"00123", "12345678901234567890"} <= identifiers, "Identifiers must retain text type and digits"


def check_printed_pages(pdf, output_dir):
    info = run("pdfinfo", str(pdf)).stdout.decode("utf-8", errors="replace")
    count = int(re.search(r"^Pages:\s+(\d+)", info, re.M)[1])
    assert 1 <= count <= 10, ("Unexpected page count", count)
    prefix = output_dir / pdf.stem
    run("pdftoppm", "-r", "18", "-gray", str(pdf), str(prefix))
    pages = sorted(output_dir.glob(prefix.name + "-*.pgm"))
    assert len(pages) == count
    for page in pages:
        raw = page.read_bytes()
        header = re.match(rb"P5\s+(\d+)\s+(\d+)\s+255\r?\n", raw)
        assert header, "Unsupported grayscale raster"
        pixels = raw[header.end():]
        assert len(pixels) == int(header[1]) * int(header[2])
        assert sum(b < 232 for b in pixels) > 20, ("Blank printed page", page.name)
    return count


def main():
    work = ROOT / "work"
    work.mkdir(exist_ok=True)
    print(run(SOFFICE, "--version").stdout.decode().strip(), flush=True)
    with tempfile.TemporaryDirectory(prefix="office-reader-", dir=work) as folder:
        temp = Path(folder)
        docs = [EXAMPLES / f"officebridge-demo.{ext}" for ext in ("pptx", "xlsx")]
        repaired = temp / "repaired"
        run("node", "scripts/cli.mjs", "repair", *docs, "--accept-font-substitutions", "--out", repaired)
        profile = "-env:UserInstallation=" + (temp / "profile").as_uri()
        for source, odf in zip(docs, ("odp", "ods")):
            copies = (source, repaired / (source.stem + ".windows" + source.suffix))
            page_counts = []
            for doc in copies:
                output = temp / (odf + "-" + doc.stem)
                output.mkdir()
                for fmt in (odf, "pdf"):
                    run(SOFFICE, profile, "--headless", "--convert-to", fmt, "--outdir", output, doc)
                with ZipFile(output / (doc.stem + "." + odf)) as archive:
                    xml = ET.fromstring(archive.read("content.xml"))
                    if odf == "odp":
                        check_presentation(archive, xml)
                    else:
                        check_workbook(xml)
                page_counts.append(check_printed_pages(output / (doc.stem + ".pdf"), output))
            assert page_counts[0] == page_counts[1], ("Page count changed", doc.suffix, page_counts)
            if odf == "odp":
                assert page_counts[0] == 2
            print(f"{doc.suffix}: original and repaired files read, native content checked, "
                  f"{page_counts[0]} nonblank rendered pages each", flush=True)


if __name__ == "__main__":
    main()
