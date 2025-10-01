
import io
import os
import re
import tkinter as tk
from tkinter import ttk, messagebox

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
    out = re.sub(r"\$\$(.*?)\$\$", r"\1", out, flags=re.DOTALL)
    out = re.sub(r"\$(.*?)\$", r"\1", out, flags=re.DOTALL)
    out = re.sub(r"\\text\{([^}]*)\}", r"\1", out)
    def repl_frac(m):
        num = m.group(1); den = m.group(2)
        return f"({num})/({den})"
    out = re.sub(r"\\frac\{([^{}]+|\{[^}]*\})\}\{([^{}]+|\{[^}]*\})\}", repl_frac, out)
    out = re.sub(r"\^\{([^}]*)\}", r"^\1", out)
    out = re.sub(r"_\{([^}]*)\}", r"_\1", out)
    out = out.replace("{", "(").replace("}", ")")
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
        plt.close(fig)
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
        plt.close(fig)
        raise

def png_bytes_to_pil(png_bytes):
    from PIL import Image
    return Image.open(io.BytesIO(png_bytes))

def copy_image_to_windows_clipboard(img):
    if not HAVE_PYWIN32:
        raise RuntimeError("pywin32 not installed. Install with: pip install pywin32")
    with io.BytesIO() as output:
        bmp = img.convert("RGB")
        bmp.save(output, "BMP")
        data = output.getvalue()[14:]
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_DIB, data)
    finally:
        win32clipboard.CloseClipboard()

# -----------------------------
# GUI
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("LaTeX → Image / Text (Copy-Paste) — v3 safe")
        self.geometry("900x600")
        self.minsize(820, 520)
        self.last_image = None
        self.last_photo = None

        top = ttk.Frame(self); top.pack(side=tk.TOP, fill=tk.X, padx=12, pady=8)
        ttk.Label(top, text="LaTeX input (e.g., V_0 = \\frac{D_1}{r_e - g} or $...$):").pack(side=tk.TOP, anchor="w")
        self.txt = tk.Text(self, wrap="word", height=10, font=("Consolas", 12))
        self.txt.pack(side=tk.TOP, fill=tk.BOTH, expand=False, padx=12, pady=(0,8))

        options = ttk.Frame(self); options.pack(side=tk.TOP, fill=tk.X, padx=12, pady=0)
        self.fontsize_var = tk.IntVar(value=28)
        self.usetex_var = tk.BooleanVar(value=False)
        ttk.Label(options, text="Font size:").pack(side=tk.LEFT, padx=(0,6))
        spin = ttk.Spinbox(options, from_=10, to=96, textvariable=self.fontsize_var, width=5); spin.pack(side=tk.LEFT, padx=(0,12))
        ttk.Checkbutton(options, text="Use full LaTeX (MiKTeX/TeX Live)", variable=self.usetex_var).pack(side=tk.LEFT, padx=(0,12))

        btns = ttk.Frame(self); btns.pack(side=tk.TOP, fill=tk.X, padx=12, pady=8)
        ttk.Button(btns, text="Preview", command=self.on_preview).pack(side=tk.LEFT)
        ttk.Button(btns, text="Copy as IMAGE", command=self.on_copy_image).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="Copy as PLAIN TEXT", command=self.on_copy_text).pack(side=tk.LEFT, padx=6)

        self.status = tk.StringVar(value="Ready")
        ttk.Label(self, textvariable=self.status, foreground="#555").pack(side=tk.TOP, anchor="w", padx=12)

        self.preview = ttk.Label(self, anchor="center")
        self.preview.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=12, pady=8)

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
            if self.last_image is None: self.render()
            if HAVE_PYWIN32:
                copy_image_to_windows_clipboard(self.last_image); self.set_status("Image copied to clipboard ✔  (Ctrl+V into Word/OneNote)")
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



