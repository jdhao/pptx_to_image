Convert PPT files to images using Python.

This works well on Linux (Ubuntu). We need to install libreoffice,
poppler-utils and pdf2image:

```bash
apt update && apt install libreoffice poppler-utils
pip install pdf2image
```

# How to run?

```bash
python ppt_to_img.py test.pptx
```
