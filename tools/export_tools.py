"""
Export tools for PowerPoint MCP Server.
Converts presentation slides to images (PNG) via LibreOffice PDF conversion + PyMuPDF rendering.
Enables AI to visually analyze slide content including layout, charts, shapes, and images.
"""
import os
import subprocess
import tempfile
import base64
import shutil
from typing import Dict, List, Optional
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations


def _find_libreoffice() -> Optional[str]:
    """Find LibreOffice executable path."""
    candidates = [
        "libreoffice",
        "soffice",
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
        "/usr/local/bin/libreoffice",
        "/snap/bin/libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/mnt/c/Program Files/LibreOffice/program/soffice.exe",
    ]
    for candidate in candidates:
        if shutil.which(candidate):
            return candidate
    return None


def _convert_pptx_to_pdf(pptx_path: str, output_dir: str) -> str:
    """Convert PPTX to PDF using LibreOffice headless mode."""
    libreoffice_path = _find_libreoffice()
    if not libreoffice_path:
        raise RuntimeError(
            "LibreOffice not found. Please install LibreOffice:\n"
            "  Ubuntu/Debian: sudo apt install libreoffice\n"
            "  macOS: brew install --cask libreoffice\n"
            "  Windows: download from https://www.libreoffice.org/"
        )

    with tempfile.TemporaryDirectory(prefix="lo_profile_") as profile_dir:
        cmd = [
            libreoffice_path,
            "--headless",
            "--norestore",
            f"-env:UserInstallation=file://{profile_dir}",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            pptx_path,
        ]
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,
        )

    if result.returncode != 0:
        raise RuntimeError(
            f"LibreOffice conversion failed (exit code {result.returncode}):\n"
            f"stdout: {result.stdout}\nstderr: {result.stderr}"
        )

    base_name = os.path.splitext(os.path.basename(pptx_path))[0]
    pdf_path = os.path.join(output_dir, f"{base_name}.pdf")
    if not os.path.exists(pdf_path):
        raise RuntimeError(
            f"PDF file not found after conversion. Expected: {pdf_path}\n"
            f"LibreOffice output: {result.stdout}"
        )
    return pdf_path


def _render_pdf_to_pngs(pdf_path: str, output_dir: str, dpi: int = 150,
                         slide_numbers: Optional[List[int]] = None) -> List[str]:
    """Render PDF pages to PNG images using PyMuPDF."""
    try:
        import fitz  # PyMuPDF
    except ImportError:
        raise RuntimeError(
            "PyMuPDF not found. Please install it:\n"
            "  pip install PyMuPDF"
        )

    doc = fitz.open(pdf_path)
    png_paths = []
    zoom = dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)

    try:
        for page_num in range(len(doc)):
            slide_num = page_num + 1
            if slide_numbers and slide_num not in slide_numbers:
                continue

            page = doc[page_num]
            pix = page.get_pixmap(matrix=matrix)
            png_filename = f"slide_{slide_num:03d}.png"
            png_path = os.path.join(output_dir, png_filename)
            pix.save(png_path)
            png_paths.append(png_path)
    finally:
        doc.close()

    return png_paths


def register_export_tools(app: FastMCP, presentations: Dict, get_current_presentation_id):
    """Register export tools with the FastMCP app"""

    @app.tool(
        annotations=ToolAnnotations(
            title="Export Slides To Images",
            destructiveHint=True,
        ),
    )
    def export_slides_to_images(
        file_path: Optional[str] = None,
        output_dir: Optional[str] = None,
        slide_numbers: Optional[List[int]] = None,
        dpi: int = 150,
        include_base64: bool = False,
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Export presentation slides as PNG images for AI visual analysis.

        Converts each slide to a high-quality PNG image so that AI can visually
        analyze the full slide layout including text positioning, charts, shapes,
        images, and design elements.

        Process: PPTX -> PDF (via LibreOffice) -> PNG per slide (via PyMuPDF)

        Prerequisites:
        - LibreOffice must be installed (sudo apt install libreoffice)
        - PyMuPDF must be installed (pip install PyMuPDF)

        Args:
            file_path: Path to a PPTX file on disk. If not provided, uses the
                       currently loaded presentation (saves to temp file first).
            output_dir: Directory to save PNG files. If not provided, creates a
                        temp directory. The directory will be created if needed.
            slide_numbers: List of 1-based slide numbers to export.
                          If not provided, exports all slides.
            dpi: Resolution for rendering (default: 150). Higher = better quality
                 but larger files. Recommended: 100-200 for analysis, 300 for print.
            include_base64: Whether to include base64-encoded image data in the
                           response. Default is False. Set to True only when
                           explicitly requested, as it makes the response very
                           large. In most cases, file paths are sufficient.
            presentation_id: ID of a loaded presentation to export.
                            Defaults to the current presentation.

        Returns:
            Dictionary with exported image file paths and metadata.
        """
        temp_pptx_path = None
        source_description = ""

        if file_path:
            if not os.path.exists(file_path):
                return {"error": f"File not found: {file_path}"}
            pptx_path = os.path.abspath(file_path)
            source_description = file_path
        else:
            pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
            if pres_id is None or pres_id not in presentations:
                return {
                    "error": "No file_path provided and no presentation is currently loaded. "
                             "Please provide a file_path or load a presentation first."
                }

            pres = presentations[pres_id]
            temp_pptx = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False)
            temp_pptx_path = temp_pptx.name
            temp_pptx.close()
            try:
                pres.save(temp_pptx_path)
            except Exception as e:
                os.unlink(temp_pptx_path)
                return {"error": f"Failed to save presentation to temp file: {str(e)}"}
            pptx_path = temp_pptx_path
            source_description = f"loaded presentation '{pres_id}'"

        if output_dir:
            output_dir = os.path.abspath(output_dir)
            if not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
        else:
            output_dir = tempfile.mkdtemp(prefix="pptx_slides_")

        try:
            with tempfile.TemporaryDirectory(prefix="pptx_pdf_") as pdf_dir:
                pdf_path = _convert_pptx_to_pdf(pptx_path, pdf_dir)
                png_paths = _render_pdf_to_pngs(
                    pdf_path, output_dir, dpi=dpi, slide_numbers=slide_numbers
                )

            if not png_paths:
                return {
                    "error": "No slides were exported. Check slide_numbers parameter.",
                    "source": source_description,
                }

            slides_info = []
            for png_path in png_paths:
                file_size = os.path.getsize(png_path)
                slide_info = {
                    "file_path": png_path,
                    "file_name": os.path.basename(png_path),
                    "file_size_bytes": file_size,
                    "file_size_kb": round(file_size / 1024, 1),
                }

                if include_base64:
                    with open(png_path, "rb") as f:
                        slide_info["base64_data"] = base64.b64encode(f.read()).decode("utf-8")

                slides_info.append(slide_info)

            total_size = sum(s["file_size_bytes"] for s in slides_info)

            return {
                "message": f"Successfully exported {len(png_paths)} slide(s) as PNG images",
                "source": source_description,
                "output_dir": output_dir,
                "dpi": dpi,
                "total_slides_exported": len(png_paths),
                "total_size_bytes": total_size,
                "total_size_mb": round(total_size / (1024 * 1024), 2),
                "slides": slides_info,
                "tip": "Use the file paths with AI vision capabilities to analyze slide content visually.",
            }

        except RuntimeError as e:
            return {"error": str(e)}
        except Exception as e:
            return {"error": f"Unexpected error during export: {str(e)}"}
        finally:
            if temp_pptx_path and os.path.exists(temp_pptx_path):
                os.unlink(temp_pptx_path)
