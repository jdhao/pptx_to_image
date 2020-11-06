"""
convert PPT file to images
"""
import os
import sys
import time
import subprocess

from pdf2image import convert_from_bytes


def main():
    img_format = "jpg"
    out_dir = "ppt-previews"
    pptfile_name = sys.argv[1]

    start = time.time()
    print("Start converting your PPT to {} images.".format(img_format))

    filename_base = os.path.basename(pptfile_name)
    filename_bare = os.path.splitext(filename_base)[0]

    # soffice --headless --convert-to pdf demo.pptx
    # convert pptx to PDF
    command_list = ["soffice", "--headless", "--convert-to", "pdf", pptfile_name]
    subprocess.run(command_list)

    pdffile_name = filename_bare + ".pdf"
    with open(pdffile_name, "rb") as f:
        pdf_bytes = f.read()
    images = convert_from_bytes(pdf_bytes, dpi=96)

    if not os.path.exists(out_dir):
        os.mkdir(out_dir)

    for i, img in enumerate(images):
        im_name = os.path.join(out_dir, f"img-{i}.{img_format}")
        img.save(im_name)

    elapse = time.time() - start
    print("Conversion done, images saved in dir {}. Time spent: {}".format(
        out_dir, elapse))


if __name__ == '__main__':
    main()
