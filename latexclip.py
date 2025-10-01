
import io
import os
import re
import tkinter as tk
from tkinter import ttk, messagebox
from ttkthemes import ThemedTk

try:
    from PIL import Image, ImageTk
    # Disable decompression bomb guard explicitly
    Image.MAX_IMAGE_PIXELS = None
except Exception as e:
    raise SystemExit("Missing dependency: Pillow. Install with: pip install pillow") from e

HAVE_PYWIN32 = True
try:
    import win32clipboard
    import win32con
except Exception:
    HAVE_PYWIN32 = False

try:
    import matplotlib
    matplotlib.use("Agg")
    from matplotlib import pyplot as plt
    from matplotlib import rcParams
except Exception as e:
    raise SystemExit("Missing dependency: matplotlib. Install with: pip install matplotlib") from e

# -----------------------------
# Helpers
# -----------------------------
def latex_to_plaintext(s: str) -> str:
    out = s
    # Temporarily replace escaped braces to preserve them
    out = out.replace(r"\{", "__LACE_BRACE__").replace(r"\}", "__RACE_BRACE__")
    # Strip delimiters
    out = re.sub(r"\$\$(.*?)\$\$", r"\1", out, flags=re.DOTALL)
    out = re.sub(r"\$(.*?)\$", r"\1", out, flags=re.DOTALL)

    # Handle common functions
    out = re.sub(r"\\(sin|cos|tan|log|ln|det|dim|lim|exp|deg)\s*", r"\1", out)
    # Handle Greek letters
    greek = r"alpha|beta|gamma|delta|epsilon|zeta|eta|theta|iota|kappa|lambda|mu|nu|xi|pi|rho|sigma|tau|upsilon|phi|chi|psi|omega"
    greek += r"|Gamma|Delta|Theta|Lambda|Xi|Pi|Sigma|Upsilon|Phi|Psi|Omega"
    out = re.sub(r"\\(" + greek + r")", r"\1", out)
    # Handle symbols
    out = re.sub(r"\\(cdot|times)", "*", out)
    out = re.sub(r"\\pm", "+-", out)
    out = re.sub(r"\\mp", "-+", out)

    # Handle \sqrt
    out = re.sub(r"\\sqrt\[([^\]]*)\]\{([^}]*)\}", r"(\2)^(1/\1)", out)
    out = re.sub(r"\\sqrt\{([^}]*)\}", r"sqrt(\1)", out)

    out = re.sub(r"\\text\{([^}]*)\}", r"\1", out)
    def repl_frac(m):
        num = m.group(1); den = m.group(2)
        return f"({num})/({den})"
    out = re.sub(r"\\frac\{([^{}]+|\{[^}]*\})\}\{([^{}]+|\{[^}]*\})\}", repl_frac, out)

    out = re.sub(r"\^\{([^}]*)\}", r"^\1", out)
    out = re.sub(r"_\{([^}]*)\}", r"_\1", out)

    # Replace \left and \right with nothing, as they are for visual grouping
    out = out.replace(r"\left", "").replace(r"\right", "")

    # Now, it's safer to replace remaining braces
    out = out.replace("{", "(").replace("}", ")")
    # Restore the escaped braces
    out = out.replace("__LACE_BRACE__", "{").replace("__RACE_BRACE__", "}")
    out = re.sub(r"\s+", " ", out).strip()
    return out

def sanitize_for_mathtext(s: str) -> str:
    text = s.strip()
    if (text.startswith("$$") and text.endswith("$$")):
        text = text[2:-2].strip()
    if text.startswith(r"\[") and text.endswith(r"\]"):
        text = text[2:-2].strip()
    already_math = text.startswith("$") and text.endswith("$")
    def to_rm(m):
        return r"\mathrm{" + m.group(1) + "}"
    text = re.sub(r"\\text\{([^}]*)\}", to_rm, text)
    text = text.replace(r"\left", "").replace(r"\right", "")
    text = re.sub(r"\s+", " ", text).strip()
    if not already_math and "\n" not in text:
        text = f"${text}$"
    return text

def clamp_image(img, max_megapixels=10, max_side=6000):
    """Downscale overly large images to safe sizes."""
    w, h = img.size
    if w*h <= max_megapixels*1_000_000 and max(w,h) <= max_side:
        return img
    scale = min((max_side / max(w, h)), ( (max_megapixels*1_000_000) / (w*h) ) ** 0.5)
    if scale >= 1.0:
        return img
    new_size = (max(1, int(w*scale)), max(1, int(h*scale)))
    return img.resize(new_size, Image.LANCZOS)

def render_latex_to_png_bytes(latex: str, fontsize: int = 28, usetex: bool = False) -> bytes:
    rcParams["text.usetex"] = bool(usetex)
    rcParams["mathtext.default"] = "regular"
    content = latex.strip()
    if not usetex:
        content = sanitize_for_mathtext(content)
    else:
        if content.startswith("$$") and content.endswith("$$"):
            content = content[2:-2].strip()
        if not (content.startswith(r"\[") and content.endswith(r"\]")) and not (content.startswith("$") and content.endswith("$")):
            content = r"\[" + content + r"\]"

    # Use a sensible figure size and tight bbox instead of manual bbox math
    fig = plt.figure(figsize=(6, 2), dpi=300)
    fig.patch.set_alpha(0.0)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.axis("off")

    try:
        # Left-align to reduce chance of off-canvas metrics
        ax.text(0.01, 0.5, content, ha="left", va="center", fontsize=fontsize)
        buf = io.BytesIO()
        plt.savefig(buf, format="png", dpi=300, transparent=True, bbox_inches='tight', pad_inches=0.02)
        buf.seek(0)
        # Trim transparent padding then clamp to safe size
        img = Image.open(buf).convert("RGBA")
        alpha = img.getchannel("A")
        bbox = alpha.getbbox()
        if bbox:
            img = img.crop(bbox)
        img = clamp_image(img)
        out = io.BytesIO()
        img.save(out, format="PNG")
        out.seek(0)
        return out.getvalue()
    except Exception as e:
        # Add context to the error message
        raise RuntimeError(f"Matplotlib error: {e}\nCheck LaTeX syntax or for missing packages.") from e
    finally:
        # Ensure the figure is always closed to conserve memory
        plt.close(fig)

def png_bytes_to_pil(png_bytes):
    from PIL import Image
    return Image.open(io.BytesIO(png_bytes))

def copy_image_to_windows_clipboard(img):
    if not HAVE_PYWIN32:
        raise RuntimeError("pywin32 not installed. Install with: pip install pywin32")

    # For apps that support transparent PNGs (e.g., Word, OneNote)
    with io.BytesIO() as png_output:
        img.save(png_output, "PNG")
        png_data = png_output.getvalue()

    # For older apps, provide a DIB (BMP) without transparency
    with io.BytesIO() as bmp_output:
        # BMP format requires RGB
        bmp = img.convert("RGB")
        bmp.save(bmp_output, "BMP")
        # The DIB format is the BMP content without the 14-byte file header
        bmp_data = bmp_output.getvalue()[14:]

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        # Register the PNG format
        CF_PNG = win32clipboard.RegisterClipboardFormatW("PNG")
        # Set both formats. Apps can choose which one they prefer.
        win32clipboard.SetClipboardData(CF_PNG, png_data)
        win32clipboard.SetClipboardData(win32con.CF_DIB, bmp_data)
    finally:
        win32clipboard.CloseClipboard()

# -----------------------------
# GUI
# -----------------------------
class App(ThemedTk):
    def __init__(self):
        super().__init__(theme="aquativo")
        self.title("LaTeX Clip")
        self.geometry("900x600")
        self.minsize(820, 520)

        # System-native font for a cleaner look
        font_ui = ("Segoe UI", 10)
        font_editor = ("Consolas", 12)
        self.option_add("*Font", font_ui)

        self.last_image = None
        self.last_photo = None
        self.last_rendered_text = None

        # Use a style to configure the background color of the frames
        style = ttk.Style()
        style.configure("TFrame", background=style.lookup("TFrame", "background"))
        style.configure("Preview.TLabel", background="#f0f0f0")

        # Main container
        main_frame = ttk.Frame(self, padding=(20, 12))
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Input section
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X)
        ttk.Label(input_frame, text="LaTeX Input", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 4))
        self.txt = tk.Text(input_frame, wrap="word", height=8, font=font_editor, relief=tk.FLAT, borderwidth=10)
        self.txt.pack(fill=tk.BOTH, expand=True)

        # Options section
        options_frame = ttk.Frame(main_frame)
        options_frame.pack(fill=tk.X, pady=12)
        self.fontsize_var = tk.IntVar(value=28)
        self.usetex_var = tk.BooleanVar(value=False)

        ttk.Label(options_frame, text="Font Size:", font=font_ui).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Spinbox(options_frame, from_=10, to=96, textvariable=self.fontsize_var, width=5).pack(side=tk.LEFT, padx=(0, 16))
        ttk.Checkbutton(options_frame, text="Use full LaTeX (MiKTeX/TeX Live)", variable=self.usetex_var).pack(side=tk.LEFT)

        # Actions section
        btns_frame = ttk.Frame(main_frame)
        btns_frame.pack(fill=tk.X)
        ttk.Button(btns_frame, text="Preview", command=self.on_preview).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btns_frame, text="Copy as Image", command=self.on_copy_image).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns_frame, text="Copy as Plain Text", command=self.on_copy_text).pack(side=tk.LEFT, padx=8)

        # Status and Preview
        self.status = tk.StringVar(value="Ready")
        ttk.Label(main_frame, textvariable=self.status, foreground="#666").pack(anchor="w", pady=(12, 4))

        self.preview = ttk.Label(main_frame, anchor="center", style="Preview.TLabel")
        self.preview.pack(fill=tk.BOTH, expand=True, pady=4)

    def get_input(self) -> str:
        return self.txt.get("1.0", "end-1c").strip()

    def set_status(self, msg: str):
        self.status.set(msg); self.update_idletasks()

    def render(self):
        latex = self.get_input()
        if not latex:
            raise ValueError("Enter some LaTeX first.")
        png = render_latex_to_png_bytes(latex, fontsize=self.fontsize_var.get(), usetex=self.usetex_var.get())
        img = png_bytes_to_pil(png)
        self.last_image = img
        self.last_rendered_text = latex
        max_w, max_h = 820, 280
        disp = img
        if img.width > max_w or img.height > max_h:
            disp = img.copy(); disp.thumbnail((max_w, max_h))
        self.last_photo = ImageTk.PhotoImage(disp)
        self.preview.configure(image=self.last_photo)

    def on_preview(self):
        try:
            self.render(); self.set_status("Preview updated ✔")
        except Exception as e:
            messagebox.showerror("Render error", str(e)); self.set_status("Render failed")

    def on_copy_image(self):
        try:
            # Re-render if text changed since last render
            if self.last_image is None or self.get_input() != self.last_rendered_text:
                self.render()
            if HAVE_PYWIN32:
                copy_image_to_windows_clipboard(self.last_image); self.set_status("Image (with transparency) copied to clipboard ✔  (Ctrl+V into Word/OneNote)")
            else:
                out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "latex_output.png")
                self.last_image.save(out, "PNG")
                messagebox.showinfo("Saved image", f"pywin32 not found. Saved image to:\n{out}\nYou can insert this file into Word/OneNote.")
                self.set_status(f"Saved image to {out}")
        except Exception as e:
            messagebox.showerror("Copy error", str(e)); self.set_status("Copy failed")

    def on_copy_text(self):
        try:
            raw = self.get_input()
            if not raw: raise ValueError("Enter some LaTeX first.")
            txt = latex_to_plaintext(raw)
            self.clipboard_clear(); self.clipboard_append(txt)
            self.set_status("Plain text copied to clipboard ✔  (Ctrl+V into Word/OneNote)")
        except Exception as e:
            messagebox.showerror("Copy error", str(e)); self.set_status("Copy failed")

if __name__ == "__main__":
    app = App()
    app.mainloop()



