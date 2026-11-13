#!/usr/bin/env python3
"""Ubah PDF menjadi PDF gambar: tiap halaman jadi JPEG, tanpa lapisan teks.

Teks di hasil akhir tidak bisa diblok/dicari/dicopy (hanya piksel).

Butuh: poppler-utils (pdftoppm, pdfinfo) di PATH.

Pemakaian:
    python3 pdf-to-image-pdf.py in.pdf                      # -> in.gambar.pdf
    python3 pdf-to-image-pdf.py in.pdf -o out.pdf -d 200
    python3 pdf-to-image-pdf.py *.pdf -d 150 --suffix .gambar
"""
import argparse
import glob
import os
import shutil
import subprocess
import sys
import tempfile

A4_W_PT = 595.44
A4_H_PT = 842.04


def page_size_pt(pdf):
    """Ukuran halaman pertama (pt) lewat pdfinfo; jatuh ke A4 bila gagal."""
    try:
        out = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True, check=True).stdout
        for line in out.splitlines():
            if line.startswith("Page size:"):
                nums = line.split()
                return float(nums[2]), float(nums[4])
    except (subprocess.CalledProcessError, IndexError, ValueError):
        pass
    return A4_W_PT, A4_H_PT


def pdf_to_jpegs(pdf, outdir, dpi):
    subprocess.run(["pdftoppm", "-jpeg", "-r", str(dpi), pdf, os.path.join(outdir, "p")], check=True)
    return sorted(glob.glob(os.path.join(outdir, "p-*.jpg")))


def image_size(jpg):
    """Lebar/tinggi dari header JPEG (SOFn), tanpa dependency PIL."""
    with open(jpg, "rb") as f:
        data = f.read()
    i = 2
    while i < len(data):
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
            h = int.from_bytes(data[i + 5:i + 7], "big")
            w = int.from_bytes(data[i + 7:i + 9], "big")
            return w, h
        i += 2 + int.from_bytes(data[i + 2:i + 4], "big")
    raise ValueError(f"tidak bisa membaca dimensi JPEG: {jpg}")


def write_image_pdf(jpgs, out_pdf, page_pt, dpi):
    """Rakit PDF: 1 halaman = 1 XObject gambar JPEG, isi halaman hanya 'Do'."""
    w_pt, h_pt = page_pt
    objects = {}

    kids = " ".join(f"{4 + i * 3} 0 R" for i in range(len(jpgs)))
    objects[1] = b"<< /Type /Catalog /Pages 2 0 R >>"
    objects[2] = f"<< /Type /Pages /Kids [{kids}] /Count {len(jpgs)} >>".encode()

    for i, jpg in enumerate(jpgs):
        img_obj, page_obj, cont_obj = 3 + i * 3, 4 + i * 3, 5 + i * 3
        w, h = image_size(jpg)
        jpeg = open(jpg, "rb").read()

        objects[page_obj] = (
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {w_pt} {h_pt}] "
            f"/Resources << /XObject << /Im0 {img_obj} 0 R >> >> "
            f"/Contents {cont_obj} 0 R >>"
        ).encode()
        objects[img_obj] = (
            f"<< /Type /XObject /Subtype /Image /Width {w} /Height {h} "
            f"/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode "
            f"/Length {len(jpeg)} >>\nstream\n"
        ).encode() + jpeg + b"\nendstream"
        content = f"q {w_pt} 0 0 {h_pt} 0 0 cm /Im0 Do Q"
        objects[cont_obj] = (f"<< /Length {len(content)} >>\nstream\n{content}\nendstream").encode()

    size = len(jpgs) * 3 + 3
    out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = {}
    for num in range(1, size):
        offsets[num] = len(out)
        out += f"{num} 0 obj\n".encode() + objects[num] + b"\nendobj\n"

    xref = len(out)
    out += f"xref\n0 {size}\n0000000000 65535 f \n".encode()
    for num in range(1, size):
        out += f"{offsets[num]:010d} 00000 n \n".encode()
    out += f"trailer\n<< /Size {size} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()

    with open(out_pdf, "wb") as f:
        f.write(out)


def convert(pdf, out_pdf, dpi, keep_jpegs_dir=None):
    page_pt = page_size_pt(pdf)
    tmp = keep_jpegs_dir or tempfile.mkdtemp(prefix="gbr-")
    try:
        jpgs = pdf_to_jpegs(pdf, tmp, dpi)
        if not jpgs:
            raise RuntimeError("pdftoppm tidak menghasilkan gambar")
        write_image_pdf(jpgs, out_pdf, page_pt, dpi)
        return len(jpgs), os.path.getsize(out_pdf)
    finally:
        if keep_jpegs_dir is None:
            shutil.rmtree(tmp, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser(description="PDF -> PDF gambar (teks tidak bisa dicopy)")
    ap.add_argument("inputs", nargs="+", help="file PDF")
    ap.add_argument("-o", "--output", help="nama file hasil (hanya untuk satu input)")
    ap.add_argument("-d", "--dpi", type=int, default=200, help="resolusi render (default 200)")
    ap.add_argument("--suffix", default=".gambar", help="akhiran nama hasil (default .gambar)")
    args = ap.parse_args()

    if args.output and len(args.inputs) > 1:
        sys.exit("--output hanya boleh untuk satu file input")

    gagal = 0
    for pdf in args.inputs:
        out = args.output or os.path.splitext(pdf)[0] + args.suffix + ".pdf"
        try:
            pages, size = convert(pdf, out, args.dpi)
            print(f"OK    {os.path.basename(out)} -- {pages} hal, {size / 1048576:.2f} MB")
        except Exception as exc:  # noqa: BLE001 - lanjutkan batch
            gagal += 1
            print(f"GAGAL {os.path.basename(pdf)} -- {exc}")
    sys.exit(1 if gagal else 0)


if __name__ == "__main__":
    main()