"""
Export tools for PowerPoint MCP Server.
Converts presentation slides to images (PNG).

Two rendering back ends are used, selected by platform:
- Windows: the PowerPoint COM automation API (Presentation slides are exported
  directly to PNG). PowerPoint is already required to be installed there and it
  renders the deck exactly like PowerPoint does.
- Everything else (Linux, Docker, macOS, WSL): LibreOffice PDF conversion +
  PyMuPDF rendering, which is the original and only portable route.

Enables AI to visually analyze slide content including layout, charts, shapes, and images.
"""
import os
import sys
import subprocess
import tempfile
import base64
import shutil
from typing import Dict, List, Optional
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations


def _find_libreoffice() -> Optional[str]:
    """Find LibreOffice executable path.

    Only used on non-Windows platforms; Windows exports through PowerPoint COM
    and never reaches this function.
    """
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
    """Convert PPTX to PDF using LibreOffice headless mode.

    Only used on non-Windows platforms; Windows exports through PowerPoint COM
    and never reaches this function.
    """
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
        import pymupdf as fitz  # PyMuPDF >= 1.24.3
    except ImportError:
        try:
            import fitz  # older PyMuPDF releases
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


def _same_file_path(left: str, right: str) -> bool:
    """Compare two Windows paths for equality without touching the file system."""
    return os.path.normcase(os.path.normpath(left)) == os.path.normcase(os.path.normpath(right))


def _find_open_presentation(app, pptx_path: str):
    """Return an already-open presentation for pptx_path, or None.

    PowerPoint runs as a single instance per user, so a deck the user already has
    open shows up in our automation session too. Reusing that object instead of
    opening a second copy is what keeps us from closing the user's document at
    cleanup time.
    """
    try:
        count = app.Presentations.Count
    except Exception:
        return None

    for index in range(1, count + 1):
        try:
            candidate = app.Presentations.Item(index)
            full_name = candidate.FullName
        except Exception:
            continue
        if full_name and _same_file_path(full_name, pptx_path):
            return candidate
    return None


def _export_via_powerpoint_com(pptx_path: str, output_dir: str, dpi: int = 150,
                               slide_numbers: Optional[List[int]] = None) -> List[str]:
    """Export slides to PNG through the PowerPoint COM automation API (Windows only).

    Deliberately conservative about the user's own PowerPoint session:
    - Attaches to a running instance instead of fighting it for the single-instance
      server, and only launches PowerPoint when nothing is running.
    - Opens our deck read-only and window-less so it neither steals focus nor
      modifies the source file.
    - Only closes the presentation we opened ourselves, and only quits the
      application we started ourselves, and even then only when no other
      presentation is left open.
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise RuntimeError(
            "pywin32 not found. Windows slide export uses the PowerPoint COM API.\n"
            "  pip install pywin32"
        )

    pptx_path = os.path.abspath(pptx_path)
    output_dir = os.path.abspath(output_dir)

    pythoncom.CoInitialize()
    app = None
    presentation = None
    started_powerpoint = False
    opened_presentation = False
    try:
        try:
            app = win32com.client.GetActiveObject("PowerPoint.Application")
        except pythoncom.com_error:
            try:
                app = win32com.client.Dispatch("PowerPoint.Application")
            except pythoncom.com_error as e:
                raise RuntimeError(
                    "Failed to start Microsoft PowerPoint through COM. PowerPoint must be "
                    f"installed to export slides on Windows.\nCOM error: {e}"
                )
            started_powerpoint = True

        try:
            presentation = _find_open_presentation(app, pptx_path)
            if presentation is None:
                try:
                    presentation = app.Presentations.Open(
                        pptx_path,
                        ReadOnly=True,
                        Untitled=False,
                        WithWindow=False,
                    )
                except pythoncom.com_error as e:
                    raise RuntimeError(
                        f"PowerPoint could not open the presentation: {pptx_path}\n"
                        f"The file may be corrupted, password protected, or not a "
                        f"PowerPoint document.\nCOM error: {e}"
                    )
                opened_presentation = True

            slide_count = presentation.Slides.Count
            page_setup = presentation.PageSetup
            width_px = max(1, int(round(float(page_setup.SlideWidth) * dpi / 72.0)))
            height_px = max(1, int(round(float(page_setup.SlideHeight) * dpi / 72.0)))

            png_paths = []
            for slide_num in range(1, slide_count + 1):
                if slide_numbers and slide_num not in slide_numbers:
                    continue

                png_filename = f"slide_{slide_num:03d}.png"
                png_path = os.path.join(output_dir, png_filename)
                presentation.Slides.Item(slide_num).Export(
                    png_path, "PNG", width_px, height_px
                )
                if not os.path.exists(png_path):
                    raise RuntimeError(
                        f"PowerPoint reported success but no image was written for slide "
                        f"{slide_num}. Expected: {png_path}"
                    )
                png_paths.append(png_path)

            return png_paths
        except pythoncom.com_error as e:
            raise RuntimeError(f"PowerPoint COM export failed: {e}")
    finally:
        # Best effort cleanup: never let a teardown failure mask the real error or
        # skip CoUninitialize, but never leave an orphaned PowerPoint process either.
        try:
            if presentation is not None and opened_presentation:
                try:
                    presentation.Close()
                except Exception:
                    pass
            presentation = None

            if app is not None and started_powerpoint:
                try:
                    if app.Presentations.Count == 0:
                        app.Quit()
                except Exception:
                    pass
            app = None
        finally:
            pythoncom.CoUninitialize()


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

        Process depends on the platform:
        - Windows: PPTX -> PNG per slide directly (via the PowerPoint COM API)
        - Other platforms: PPTX -> PDF (via LibreOffice) -> PNG per slide (via PyMuPDF)

        Prerequisites:
        - Windows: Microsoft PowerPoint and pywin32 (pip install pywin32).
          LibreOffice is not used and does not need to be installed.
        - Other platforms: LibreOffice (sudo apt install libreoffice) and
          PyMuPDF (pip install PyMuPDF)

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
            if sys.platform == "win32":
                renderer = "powerpoint-com"
                png_paths = _export_via_powerpoint_com(
                    pptx_path, output_dir, dpi=dpi, slide_numbers=slide_numbers
                )
            else:
                renderer = "libreoffice+pymupdf"
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
                "renderer": renderer,
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
