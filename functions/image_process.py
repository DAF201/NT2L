import core.PDF
import core.OCR
from PIL import Image


def pdf_to_images(path: str) -> list[Image.Image]:
    """convert a pdf to a set of images"""
    return core.PDF.global_PDF_convertor.convert_pdf(path)


def read_images(images: list[Image.Image]) -> str:
    """read a set of images from a pdf file, then convert to a chunk of text"""
    OCR_reader = core.OCR.global_ocr_reader
    res = ""
    for image in images:
        res += OCR_reader.read(image)+"\n"
    return res
