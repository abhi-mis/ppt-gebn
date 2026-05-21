import os
import base64
import subprocess
import tempfile
import glob


def generate_ppt_images(data, emit_fn=None):
    """
    Convert each slide of a PPTX to a PNG image and stream them
    one by one via the emit_fn (socket emit).

    Uses LibreOffice headless to convert PPTX → PDF → individual PNGs.
    Falls back to pdf2image if available.

    Args:
        data: dict with "ppt_name" key
        emit_fn: callable to emit each slide image (usually socketio emit)
    
    Returns:
        dict with status and total slide count
    """
    ppt_name = data.get("ppt_name", None)
    if not ppt_name:
        print("No ppt name provided in the message.")
        return {"error": "No ppt name provided.", "status": "error"}

    ppt_path = os.path.join("generated", ppt_name)
    if not os.path.exists(ppt_path):
        print("PPT file not found:", ppt_path)
        return {"error": "No ppt found with this name.", "status": "error"}

    try:
        with tempfile.TemporaryDirectory() as tmp_dir:
            # Step 1: Convert PPTX to PDF using LibreOffice headless
            pdf_path = _convert_pptx_to_pdf(ppt_path, tmp_dir)

            if not pdf_path:
                return {"error": "Failed to convert PPTX to PDF.", "status": "error"}

            # Step 2: Convert PDF pages to individual PNG images
            image_paths = _convert_pdf_to_images(pdf_path, tmp_dir)

            if not image_paths:
                return {"error": "Failed to convert PDF to images.", "status": "error"}

            total_slides = len(image_paths)
            print(f"Generated {total_slides} slide images")

            # Step 3: Stream each image via socket
            for index, img_path in enumerate(image_paths):
                slide_number = index + 1

                with open(img_path, "rb") as f:
                    img_bytes = f.read()
                    img_base64 = base64.b64encode(img_bytes).decode("utf-8")

                slide_data = {
                    "slide_number": slide_number,
                    "total_slides": total_slides,
                    "image": img_base64,
                    "mime_type": "image/png",
                    "ppt_name": ppt_name,
                }

                if emit_fn:
                    emit_fn("slide_image", slide_data)
                    print(f"Emitted slide {slide_number}/{total_slides}")

            # Emit completion event
            if emit_fn:
                emit_fn("slide_images_complete", {
                    "total_slides": total_slides,
                    "ppt_name": ppt_name,
                    "status": "complete",
                })

            return {
                "status": "complete",
                "total_slides": total_slides,
                "ppt_name": ppt_name,
            }

    except Exception as e:
        print(f"Error generating slide images: {e}")
        if emit_fn:
            emit_fn("slide_images_error", {
                "error": str(e),
                "ppt_name": ppt_name,
                "status": "error",
            })
        return {"error": str(e), "status": "error"}


def _convert_pptx_to_pdf(pptx_path, output_dir):
    """Convert PPTX to PDF using LibreOffice headless."""
    try:
        subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", output_dir,
                pptx_path,
            ],
            check=True,
            timeout=120,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )

        # Find the generated PDF
        base_name = os.path.splitext(os.path.basename(pptx_path))[0]
        pdf_path = os.path.join(output_dir, f"{base_name}.pdf")

        if os.path.exists(pdf_path):
            return pdf_path

        # Fallback: look for any PDF in the output dir
        pdfs = glob.glob(os.path.join(output_dir, "*.pdf"))
        return pdfs[0] if pdfs else None

    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        print(f"LibreOffice conversion failed: {e}")
        return None


def _convert_pdf_to_images(pdf_path, output_dir, dpi=200):
    """
    Convert PDF pages to PNG images.
    Tries pdf2image (poppler) first, falls back to PyMuPDF (fitz).
    """
    # Try pdf2image (requires poppler)
    try:
        from pdf2image import convert_from_path

        images = convert_from_path(pdf_path, dpi=dpi, output_folder=output_dir, fmt="png")
        image_paths = []
        for i, img in enumerate(images):
            img_path = os.path.join(output_dir, f"slide_{i + 1:03d}.png")
            img.save(img_path, "PNG")
            image_paths.append(img_path)

        return sorted(image_paths)

    except ImportError:
        print("pdf2image not available, trying PyMuPDF...")

    # Fallback: PyMuPDF (fitz)
    try:
        import fitz  # PyMuPDF

        doc = fitz.open(pdf_path)
        image_paths = []

        for page_num in range(len(doc)):
            page = doc[page_num]
            # Scale for good quality
            zoom = dpi / 72
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)

            img_path = os.path.join(output_dir, f"slide_{page_num + 1:03d}.png")
            pix.save(img_path)
            image_paths.append(img_path)

        doc.close()
        return sorted(image_paths)

    except ImportError:
        print("PyMuPDF not available either.")
        return None