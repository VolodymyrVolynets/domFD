from PIL import Image
from numpy.ma.testutils import assert_not_equal
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, TextStringObject
import os
import shutil

class PDFManipulator:
    """Utility class for PDF operations: image conversion, merging, form filling, and mixed file combining."""

    def __init__(self, output_dir: str = "output_pdfs"):
        self.output_dir = output_dir
        os.makedirs(output_dir, exist_ok=True)

    # General helpers -----------------------------------------------------------------
    def convert_to_pdf(self, input_path: str, output_name: str) -> str:
        """Convert a single file (image or PDF) into a PDF stored in ``output_dir``."""
        os.makedirs(self.output_dir, exist_ok=True)
        if not output_name.lower().endswith(".pdf"):
            output_name = f"{output_name}.pdf"
        destination = os.path.join(self.output_dir, output_name)

        if os.path.abspath(input_path) == os.path.abspath(destination):
            print(f"[INFO] Source and destination are the same: {destination}")
            return destination

        ext = os.path.splitext(input_path)[1].lower()
        if ext in {".jpg", ".jpeg", ".png", ".bmp", ".tiff"}:
            image = Image.open(input_path).convert("RGB")
            image.save(destination, "PDF")
        elif ext == ".pdf":
            shutil.copyfile(input_path, destination)
        else:
            raise ValueError(f"Unsupported file format for conversion: {input_path}")

        print(f"[✔] Stored converted PDF: {destination}")
        return destination

    # 1️⃣ Convert images to a single PDF
    def images_to_pdf(self, image_paths: list[str], output_name: str = "images_to_pdf.pdf") -> str:
        pdf_path = os.path.join(self.output_dir, output_name)
        images = [Image.open(img).convert("RGB") for img in image_paths]
        images[0].save(pdf_path, save_all=True, append_images=images[1:])
        print(f"[✔] PDF created from images: {pdf_path}")
        return pdf_path

    # 2️⃣ Merge multiple PDF files
    def merge_pdfs(self, pdf_paths: list[str], output_name: str = "merged.pdf") -> str:
        pdf_path = os.path.join(self.output_dir, output_name)
        writer = PdfWriter()
        for pdf in pdf_paths:
            reader = PdfReader(pdf)
            for page in reader.pages:
                writer.add_page(page)
        with open(pdf_path, "wb") as f:
            writer.write(f)
        print(f"[✔] PDFs merged into: {pdf_path}")
        return pdf_path

    # 3️⃣ Fill a fillable PDF form
    def fill_pdf_form(self, template_path: str, data_dict: dict[str, str], output_name: str = "filled_form.pdf") -> str:
        pdf_path = os.path.join(self.output_dir, output_name)
        print(f"[INFO] Reading PDF template: {template_path}")

        reader = PdfReader(template_path)
        writer = PdfWriter()

        for page_number, page in enumerate(reader.pages, start=1):
            annots = page.get("/Annots")
            if not annots:
                continue

            for annot_ref in annots:
                annot = annot_ref.get_object() if hasattr(annot_ref, "get_object") else annot_ref
                print(annot_ref)

                key = annot.get("/T")
                if not key:
                    mk = annot.get("/MK")
                    if mk and "/CA" in mk:
                        key = mk["/CA"].strip("()")

                if not key:
                    continue

                key = str(key)
                value = None
                if key in data_dict:
                    value = data_dict[key]
                else:
                    if "#" in key:
                        base_key = key.split("#", 1)[0]
                        value = data_dict.get(base_key)
                if value is not None:
                    annot[NameObject("/V")] = TextStringObject(value)
                    print(f"  [UPDATE] Field '{key}' → {value}")

            writer.add_page(page)

        if "/AcroForm" in reader.trailer["/Root"]:
            writer._root_object.update(
                {NameObject("/AcroForm"): reader.trailer["/Root"]["/AcroForm"]}
            )

        with open(pdf_path, "wb") as f:
            writer.write(f)

        print(f"[✔] Filled form saved to: {pdf_path}")
        return pdf_path

    # 4️⃣ Combine images and PDFs into one PDF
    def combine_files(self, file_paths: list[str], output_name: str = "combined.pdf") -> str:
        """
        Combine a mix of images and PDFs into one output PDF.

        Args:
            file_paths: list of image or PDF file paths
            output_name: name of output PDF
        """
        pdf_path = os.path.join(self.output_dir, output_name)
        temp_pdfs = []
        writer = PdfWriter()

        for path in file_paths:
            ext = os.path.splitext(path)[1].lower()
            if ext in [".jpg", ".jpeg", ".png", ".bmp", ".tiff"]:
                # Convert image to a temporary one-page PDF in memory
                image = Image.open(path).convert("RGB")
                temp_path = os.path.join(self.output_dir, f"_temp_{os.path.basename(path)}.pdf")
                image.save(temp_path, "PDF")
                temp_pdfs.append(temp_path)
                print(f"[INFO] Converted image to temporary PDF: {temp_path}")
                reader = PdfReader(temp_path)
            elif ext == ".pdf":
                reader = PdfReader(path)
            else:
                print(f"[WARN] Skipping unsupported file: {path}")
                continue

            for page in reader.pages:
                writer.add_page(page)

        with open(pdf_path, "wb") as f:
            writer.write(f)

        # Clean up temporary PDFs
        for temp in temp_pdfs:
            os.remove(temp)

        print(f"[✔] Combined file created: {pdf_path}")
        return pdf_path
