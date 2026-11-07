"""Open test files with an independent office suite (offline by default).

Run after npm ci. Requires Python 3.9+, LibreOffice, pdfinfo and pdftoppm.
All documents, profiles and previews created by this check are temporary.
--with-apple-fixtures also checks two checksum-pinned public exports.
Pass --apple-fixture-dir DIR with that switch to supply them without a download.
"""
import argparse
import hashlib
import os
import re
import subprocess
import tempfile
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.dom import minidom
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
SOFFICE = "soffice.com" if os.name == "nt" else "soffice"
EXAMPLES = ROOT / "public/examples"
NS = {prefix: f"urn:oasis:names:tc:opendocument:xmlns:{prefix}:1.0"
      for prefix in ("chart", "table", "text", "office")}
NS["draw"] = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
SML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
APPLE_FIXTURES = (
    ("numbers_table.xlsx", "scosman/freecell", "c026c7d597903d7e3e3f2a07bfc372ae66d44460",
     "app/crates/freecell-engine/tests/fixtures/numbers_table.xlsx",
     "6fdaf2216a3b715eaf4c58497150721d0f2d725c023eb647151c1f413909b24c"),
    ("slide-deck-keynote.pptx", "jaapstronks/deckyard", "587c0bc312ca53dc0d682c09e033c3cb69164f44",
     "tests/fixtures/slide-deck-keynote.pptx",
     "c0ff4d0f530db56b4e7e5a5f8266ba605102f8a581ed02419cc0752bbe128600"),
)


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        raise ValueError("Unexpected redirect while downloading a pinned test fixture")


def fetch_apple_fixtures(temp, local_dir):
    directory = temp / "apple-inputs"
    directory.mkdir()
    paths = []
    remaining = 100 * 1024  # Aggregate limit for these two public files.
    for name, repository, commit, path, expected in APPLE_FIXTURES:
        url = f"https://raw.githubusercontent.com/{repository}/{commit}/{path}"
        if local_dir is None:
            source = urllib.request.build_opener(NoRedirect).open(url, timeout=30)
        else:
            source = (local_dir / name).open("rb")
        with source:
            data = source.read(remaining + 1)
        if len(data) > remaining:
            raise ValueError("Apple fixtures exceed the combined 100 KiB limit")
        if hashlib.sha256(data).hexdigest() != expected:
            raise ValueError(f"Checksum mismatch: {name}")
        remaining -= len(data)
        dest = directory / name
        dest.write_bytes(data)
        paths.append(dest)
        print(f"Verified {name} ({repository}@{commit[:12]}; {len(data)} bytes)", flush=True)
    return paths


def run(*args):
    return subprocess.run(args, cwd=ROOT, check=True, capture_output=True, timeout=120)


def chart_from_slide(archive, slide):
    obj = slide.find(".//draw:object", NS)
    assert obj is not None, "Missing embedded object"
    href = obj.attrib["{http://www.w3.org/1999/xlink}href"].removeprefix("./")
    chart = ET.fromstring(archive.read(href + "/content.xml")).find(".//chart:chart", NS)
    assert chart is not None, "Chart must be recognized as a chart"
    return chart


def check_presentation(archive, xml):
    slides = xml.findall(".//draw:page", NS)
    assert len(slides) == 2, "Expected two editable slides"
    for slide, expected in zip(slides, ("跨系统，保留可编辑内容", "图表也应该继续可编辑")):
        paragraphs = ["".join(p.itertext()) for p in slide.findall(".//text:p", NS)]
        assert expected in paragraphs, ("Missing editable title", expected, paragraphs)
    chart_from_slide(archive, slides[1])


def check_keynote(archive, xml):
    slides = xml.findall(".//draw:page", NS)
    titles = ["Voorbeeldpresentatie", "Drie punten", "Een kleine tabel",
              "Beeld met tekst", "Groei per kwartaal", "Hoofdstuk twee"]
    assert len(slides) == len(titles), ("Keynote slide count", len(slides))
    for slide, title in zip(slides, titles):
        paragraphs = {"".join(p.itertext()).strip() for p in slide.findall(".//text:p", NS)}
        assert title in paragraphs, ("Missing editable slide text", title)
    table = slides[2].find(".//table:table", NS)
    assert table is not None, "Third slide must contain a native table"
    rows = [["".join(c.itertext()).strip() for c in row.findall("table:table-cell", NS)]
            for row in table.findall(".//table:table-row", NS)]
    assert rows[0][:3] == ["Product", "Aantal", "Prijs"] and rows[1][:2] == ["Appel", "12"], rows
    assert chart_from_slide(archive, slides[4]).find(".//chart:series", NS) is not None, "Missing chart series"


def check_workbook(xml):
    sheets = xml.findall(".//office:spreadsheet/table:table", NS)
    assert [s.get("{%s}name" % NS["table"]) for s in sheets] == ["预算示例", "编号与说明"]
    results = {c.get("{%s}formula" % NS["table"]): c.get("{%s}value" % NS["office"])
               for c in sheets[0].findall(".//table:table-cell", NS)}
    assert results.get("of:=SUM([.C2:.C4])") == "4750", ("Expected recalculated formula", results)
    identifiers = {"".join(c.itertext()) for c in sheets[1].findall(".//table:table-cell", NS)
                   if c.get("{%s}value-type" % NS["office"]) == "string"}
    assert {"00123", "12345678901234567890"} <= identifiers, "Identifiers must retain text type and digits"


def check_numbers(xml):
    sheets = xml.findall(".//office:spreadsheet/table:table", NS)
    assert len(sheets) == 1
    table = "{%s}" % NS["table"]
    found = {}
    row_index = 1
    for row in sheets[0].iter(table + "table-row"):
        col_index = 1
        for cell in row:
            if cell.tag not in (table + "table-cell", table + "covered-table-cell"):
                continue
            if cell.get(table + "formula"):
                value = cell.get("{%s}value" % NS["office"])
                assert value is not None, ("No numeric formula result", row_index, col_index)
                found[row_index, col_index] = float(value)
            col_index += int(cell.get(table + "number-columns-repeated", "1"))
        row_index += int(row.get(table + "number-rows-repeated", "1"))
    expected = {(3, 3): 16384, (18, 2): 32767, (18, 3): 32768}
    for row in range(4, 18):
        expected[row, 2] = 2 ** (row - 3)
        expected[row, 3] = 2 ** max(1, 17 - row)
    assert len(expected) == 31 and found == expected, ("Numbers formula results", found)


def verify_uncached_recalculation(repaired, temp, profile):
    # The real XLSX stores correct cached values. On a separate, disposable
    # copy strip those caches, then require Calc to reproduce all 31 results.
    probe = temp / "numbers-uncached.xlsx"
    removed = 0
    with ZipFile(repaired) as src, ZipFile(probe, "x") as dst:
        for member in src.infolist():
            data = src.read(member)
            if member.filename == "xl/worksheets/sheet1.xml":
                with minidom.parseString(data) as doc:
                    for cell in doc.getElementsByTagNameNS(SML, "c"):
                        if not cell.getElementsByTagNameNS(SML, "f"):
                            continue
                        for value in cell.getElementsByTagNameNS(SML, "v"):
                            cell.removeChild(value)
                            removed += 1
                    data = doc.toxml(encoding="utf-8")
            dst.writestr(member, data)
    assert removed == 31, ("Unexpected number of formula caches", removed)
    run(SOFFICE, profile, "--headless", "--convert-to", "ods", "--outdir", temp, probe)
    with ZipFile(temp / "numbers-uncached.ods") as archive:
        check_numbers(ET.fromstring(archive.read("content.xml")))
    print("Numbers: all 31 formulas recalculated without caches; B18=32767, C18=32768", flush=True)


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
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--with-apple-fixtures", action="store_true", help="also test pinned Numbers/Keynote exports")
    parser.add_argument("--apple-fixture-dir", type=Path, help="use local copies of the same two verified exports")
    args = parser.parse_args()
    if args.apple_fixture_dir is not None and not args.with_apple_fixtures:
        parser.error("--apple-fixture-dir requires --with-apple-fixtures")
    work = ROOT / "work"
    work.mkdir(exist_ok=True)
    print(run(SOFFICE, "--version").stdout.decode().strip(), flush=True)
    with tempfile.TemporaryDirectory(prefix="office-reader-", dir=work) as folder:
        temp = Path(folder)
        docs = [EXAMPLES / f"officebridge-demo.{ext}" for ext in ("pptx", "xlsx")]
        if args.with_apple_fixtures:
            docs.extend(fetch_apple_fixtures(temp, args.apple_fixture_dir))
        repaired = temp / "repaired"
        run("node", "scripts/cli.mjs", "repair", *docs, "--accept-font-substitutions", "--out", repaired)
        profile = "-env:UserInstallation=" + (temp / "profile").as_uri()
        for source in docs:
            odf = "odp" if source.suffix == ".pptx" else "ods"
            copies = (source, repaired / (source.stem + ".windows" + source.suffix))
            page_counts = []
            for doc in copies:
                output = temp / (odf + "-" + doc.stem)
                output.mkdir()
                for fmt in (odf, "pdf"):
                    run(SOFFICE, profile, "--headless", "--convert-to", fmt, "--outdir", output, doc)
                with ZipFile(output / (doc.stem + "." + odf)) as archive:
                    xml = ET.fromstring(archive.read("content.xml"))
                    if source.name == "slide-deck-keynote.pptx":
                        check_keynote(archive, xml)
                    elif source.name == "numbers_table.xlsx":
                        check_numbers(xml)
                    elif odf == "odp":
                        check_presentation(archive, xml)
                    else:
                        check_workbook(xml)
                page_counts.append(check_printed_pages(output / (doc.stem + ".pdf"), output))
            assert page_counts[0] == page_counts[1], ("Page count changed", doc.suffix, page_counts)
            if odf == "odp":
                assert page_counts[0] == (6 if source.name == "slide-deck-keynote.pptx" else 2)
            if source.name == "numbers_table.xlsx":
                verify_uncached_recalculation(copies[1], temp, profile)
            print(f"{source.name}: original and repaired files read, native content checked, "
                  f"{page_counts[0]} nonblank rendered pages each", flush=True)


if __name__ == "__main__":
    main()
