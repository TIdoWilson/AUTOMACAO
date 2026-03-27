#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import io
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from PIL import Image, ImageDraw, ImageFont, ImageTk

# PDF compacto (opcional)
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader


# ----------------- Ajustes do seu template -----------------
# Caixa fixa (área interna onde o texto deve caber) — ajuste se necessário
BOX = {"left": 160, "top": 560, "right": 1288, "bottom": 1860}
PADDING_X = 44
PADDING_Y = 34

NAVY = (10, 28, 55, 255)
DARK = (25, 35, 50, 255)

BASE_FONTS = {
    "title_small": {"size": 30, "bold": True},
    "title_main":  {"size": 44, "bold": True},
    "section":     {"size": 30, "bold": True},
    "body":        {"size": 28, "bold": False},
}

SPACING = {
    "title_small": {"line": 6, "block": 8},
    "title_main":  {"line": 8, "block": 18},
    "section":     {"line": 6, "block": 6},
    "body":        {"line": 8, "block": 14},
    "body_last":   {"line": 8, "block": 0},
}

FONT_CANDIDATES = [
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
     "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
     "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"),
    (r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\arialbd.ttf"),
    ("/System/Library/Fonts/Supplemental/Arial.ttf",
     "/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
]


# ----------------- Renderização (justifica + preenche caixa) -----------------
def load_font(size: int, bold: bool) -> ImageFont.ImageFont:
    for regular, bolder in FONT_CANDIDATES:
        path = bolder if bold else regular
        if os.path.exists(path):
            return ImageFont.truetype(path, size=size)
    return ImageFont.load_default()


def make_fonts(scale: float):
    def s(key: str) -> int:
        base = BASE_FONTS[key]["size"]
        return max(18, int(round(base * scale)))
    return {
        "title_small": load_font(s("title_small"), True),
        "title_main":  load_font(s("title_main"),  True),
        "section":     load_font(s("section"),     True),
        "body":        load_font(s("body"),        False),
    }


def text_w(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont) -> int:
    b = draw.textbbox((0, 0), text, font=font)
    return int(b[2] - b[0])


def wrap_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_width: int):
    lines = []
    for para in (text or "").split("\n"):
        para = para.rstrip()
        if not para.strip():
            lines.append("")
            continue
        words = para.split()
        cur = ""
        for w in words:
            trial = (cur + " " + w).strip()
            if text_w(draw, trial, font) <= max_width:
                cur = trial
            else:
                if cur:
                    lines.append(cur)
                cur = w
        if cur:
            lines.append(cur)
    return lines


def is_last_line_of_paragraph(lines, idx):
    return idx == len(lines) - 1 or lines[idx + 1] == ""


def draw_justified_line(draw, x, y, line, font, max_width, fill):
    line = (line or "").strip()
    if not line:
        return

    bullet_prefix = ""
    rest = line

    if rest.startswith("•"):
        if rest.startswith("• "):
            bullet_prefix = "• "
            rest = rest[2:].strip()
        else:
            bullet_prefix = "•"
            rest = rest[1:].strip()

    x0 = x
    if bullet_prefix:
        draw.text((x0, y), bullet_prefix, font=font, fill=fill)
        x0 += text_w(draw, bullet_prefix, font)

    words = rest.split()
    if len(words) <= 1:
        draw.text((x0, y), rest, font=font, fill=fill)
        return

    word_widths = [text_w(draw, w, font) for w in words]
    base_space = text_w(draw, " ", font)

    total_words = sum(word_widths)
    gaps = len(words) - 1
    current = total_words + base_space * gaps

    target = max_width - (x0 - x)
    extra = max(0, target - current)
    add_per_gap = extra / gaps

    cx = x0
    for i, w in enumerate(words):
        draw.text((cx, y), w, font=font, fill=fill)
        cx += word_widths[i]
        if i < gaps:
            cx += base_space + add_per_gap


def build_blocks(title_small, title_main, intro, sections):
    blocks = []
    blocks.append({"style": "title_small", "text": title_small, "color": NAVY})
    blocks.append({"style": "title_main",  "text": title_main,  "color": NAVY})

    intro = (intro or "").strip()
    if intro:
        blocks.append({"style": "body", "text": intro, "color": DARK})

    # sections = list[(heading, body)]
    valid_sections = [(h.strip(), (b or "").strip()) for h, b in sections if (h or "").strip() or (b or "").strip()]
    for i, (h, b) in enumerate(valid_sections):
        if h:
            blocks.append({"style": "section", "text": h, "color": NAVY})
        if b:
            style = "body_last" if i == len(valid_sections) - 1 else "body"
            blocks.append({"style": style, "text": b, "color": DARK})
    return blocks


def layout_measure(draw, blocks, fonts, max_width):
    y = 0
    laid = []
    for b in blocks:
        style = b["style"]
        if style == "title_small":
            font = fonts["title_small"]; sp = SPACING["title_small"]
        elif style == "title_main":
            font = fonts["title_main"]; sp = SPACING["title_main"]
        elif style == "section":
            font = fonts["section"]; sp = SPACING["section"]
        else:
            font = fonts["body"]; sp = SPACING.get(style, SPACING["body"])

        lines = wrap_text(draw, b["text"], font, max_width)

        for ln in lines:
            if ln == "":
                y += int(font.size * 0.6)
            else:
                y += font.size + sp["line"]
        y += sp["block"]

        laid.append((b, lines, font, sp))
    return y, laid


def find_best_scale(img, blocks, max_text_width, max_text_height, lo=0.85, hi=1.70, iters=22):
    draw = ImageDraw.Draw(img)

    def fits(scale):
        fonts = make_fonts(scale)
        content_h, _ = layout_measure(draw, blocks, fonts, max_text_width)
        return content_h <= max_text_height

    while lo > 0.5 and not fits(lo):
        lo *= 0.9
    while hi < 3.0 and fits(hi):
        hi *= 1.1

    best = lo
    for _ in range(iters):
        mid = (lo + hi) / 2
        if fits(mid):
            best = mid
            lo = mid
        else:
            hi = mid
    return best


def render_to_image(template_path, title_small, title_main, intro, sections, justify=True, center_vertically=True):
    img = Image.open(template_path).convert("RGBA")
    draw = ImageDraw.Draw(img)

    left, top, right, bottom = BOX["left"], BOX["top"], BOX["right"], BOX["bottom"]
    text_left = left + PADDING_X
    text_top = top + PADDING_Y
    text_right = right - PADDING_X
    text_bottom = bottom - PADDING_Y

    max_w = text_right - text_left
    max_h = text_bottom - text_top

    blocks = build_blocks(title_small, title_main, intro, sections)
    scale = find_best_scale(img, blocks, max_w, max_h)

    fonts = make_fonts(scale)
    content_h, laid = layout_measure(draw, blocks, fonts, max_w)

    y = text_top + (max(0, (max_h - content_h) // 2) if center_vertically else 0)

    for b, lines, font, sp in laid:
        style = b["style"]
        color = b["color"]

        for i, ln in enumerate(lines):
            if ln == "":
                y += int(font.size * 0.6)
                continue

            if justify and style in ("body", "body_last"):
                if (not is_last_line_of_paragraph(lines, i)) and (len(ln.split()) > 1):
                    draw_justified_line(draw, text_left, y, ln, font, max_w, color)
                else:
                    draw.text((text_left, y), ln, font=font, fill=color)
            else:
                draw.text((text_left, y), ln, font=font, fill=color)

            y += font.size + sp["line"]
        y += sp["block"]

    return img


def save_compact_pdf(pil_img: Image.Image, pdf_path: str, max_width: int = 1080, jpeg_quality: int = 78):
    img = pil_img.convert("RGB")
    if img.width > max_width:
        new_h = int(img.height * (max_width / img.width))
        img = img.resize((max_width, new_h), Image.LANCZOS)

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=jpeg_quality, optimize=True, progressive=True)
    buf.seek(0)

    w_pt, h_pt = img.width, img.height
    c = canvas.Canvas(pdf_path, pagesize=(w_pt, h_pt))
    c.drawImage(ImageReader(buf), 0, 0, width=w_pt, height=h_pt)
    c.showPage()
    c.save()


# ----------------- GUI -----------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gerador de Comunicados (WhatsApp)")
        self.geometry("1100x720")

        self.template_path = tk.StringVar(value="")
        self.out_png = tk.StringVar(value="saida.png")
        self.out_pdf = tk.StringVar(value="saida.pdf")

        self.justify = tk.BooleanVar(value=True)
        self.center_vert = tk.BooleanVar(value=True)
        self.make_pdf = tk.BooleanVar(value=True)
        self.pdf_width = tk.IntVar(value=1080)
        self.pdf_quality = tk.IntVar(value=78)

        self.title_small = tk.StringVar(value="COMUNICADO IMPORTANTE")
        self.title_main = tk.StringVar(value="Título principal")

        self._preview_imgtk = None

        self._build_ui()

    def _build_ui(self):
        main = ttk.Frame(self, padding=12)
        main.pack(fill="both", expand=True)

        left = ttk.Frame(main)
        left.pack(side="left", fill="both", expand=True)

        right = ttk.Frame(main, width=420)
        right.pack(side="right", fill="both")
        right.pack_propagate(False)

        # Template
        row = ttk.Frame(left)
        row.pack(fill="x")
        ttk.Label(row, text="Template (PNG):").pack(side="left")
        ttk.Entry(row, textvariable=self.template_path).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row, text="Escolher...", command=self.pick_template).pack(side="left")

        # Titles
        f1 = ttk.LabelFrame(left, text="Cabeçalho")
        f1.pack(fill="x", pady=10)

        r1 = ttk.Frame(f1); r1.pack(fill="x", padx=10, pady=6)
        ttk.Label(r1, text="Linha 1 (menor):", width=16).pack(side="left")
        ttk.Entry(r1, textvariable=self.title_small).pack(side="left", fill="x", expand=True)

        r2 = ttk.Frame(f1); r2.pack(fill="x", padx=10, pady=6)
        ttk.Label(r2, text="Título (maior):", width=16).pack(side="left")
        ttk.Entry(r2, textvariable=self.title_main).pack(side="left", fill="x", expand=True)

        # Intro
        f2 = ttk.LabelFrame(left, text="Texto inicial")
        f2.pack(fill="both", pady=10)
        self.intro_txt = tk.Text(f2, height=4, wrap="word")
        self.intro_txt.pack(fill="both", expand=True, padx=10, pady=8)
        self.intro_txt.insert("1.0", "Digite aqui o texto inicial...")

        # Sections
        f3 = ttk.LabelFrame(left, text="Seções (Heading + Body)")
        f3.pack(fill="both", expand=True, pady=10)

        self.sections_frame = ttk.Frame(f3)
        self.sections_frame.pack(fill="both", expand=True, padx=10, pady=8)

        self.section_widgets = []
        for _ in range(3):
            self.add_section()

        ttk.Button(f3, text="Adicionar seção", command=self.add_section).pack(anchor="w", padx=10, pady=(0,10))

        # Options + Outputs
        opt = ttk.LabelFrame(right, text="Opções e Saída")
        opt.pack(fill="x", pady=8)

        ttk.Checkbutton(opt, text="Justificar corpo do texto", variable=self.justify).pack(anchor="w", padx=10, pady=4)
        ttk.Checkbutton(opt, text="Centralizar verticalmente na caixa", variable=self.center_vert).pack(anchor="w", padx=10, pady=4)

        out = ttk.LabelFrame(right, text="Arquivos")
        out.pack(fill="x", pady=8)

        r3 = ttk.Frame(out); r3.pack(fill="x", padx=10, pady=6)
        ttk.Label(r3, text="PNG:", width=6).pack(side="left")
        ttk.Entry(r3, textvariable=self.out_png).pack(side="left", fill="x", expand=True)
        ttk.Button(r3, text="Salvar como...", command=self.pick_out_png).pack(side="left", padx=6)

        ttk.Checkbutton(out, text="Gerar PDF compacto", variable=self.make_pdf).pack(anchor="w", padx=10, pady=4)

        r4 = ttk.Frame(out); r4.pack(fill="x", padx=10, pady=6)
        ttk.Label(r4, text="PDF:", width=6).pack(side="left")
        ttk.Entry(r4, textvariable=self.out_pdf).pack(side="left", fill="x", expand=True)
        ttk.Button(r4, text="Salvar como...", command=self.pick_out_pdf).pack(side="left", padx=6)

        r5 = ttk.Frame(out); r5.pack(fill="x", padx=10, pady=6)
        ttk.Label(r5, text="PDF largura:", width=10).pack(side="left")
        ttk.Entry(r5, textvariable=self.pdf_width, width=6).pack(side="left")
        ttk.Label(r5, text="Qualidade:", width=9).pack(side="left", padx=(12,0))
        ttk.Entry(r5, textvariable=self.pdf_quality, width=6).pack(side="left")

        # Buttons
        btns = ttk.Frame(right)
        btns.pack(fill="x", pady=10)

        ttk.Button(btns, text="Pré-visualizar", command=self.preview).pack(side="left", fill="x", expand=True, padx=(0,6))
        ttk.Button(btns, text="Gerar arquivo(s)", command=self.generate).pack(side="left", fill="x", expand=True)

        # Preview area
        pv = ttk.LabelFrame(right, text="Preview")
        pv.pack(fill="both", expand=True, pady=8)
        self.preview_label = ttk.Label(pv)
        self.preview_label.pack(fill="both", expand=True, padx=10, pady=10)

    def add_section(self):
        if len(self.section_widgets) >= 6:
            messagebox.showinfo("Limite", "Máximo de 6 seções.")
            return

        frame = ttk.Frame(self.sections_frame)
        frame.pack(fill="x", pady=6)

        heading_var = tk.StringVar(value="")
        ttk.Label(frame, text="Heading:", width=9).grid(row=0, column=0, sticky="w")
        heading_entry = ttk.Entry(frame, textvariable=heading_var)
        heading_entry.grid(row=0, column=1, sticky="ew", padx=6)

        ttk.Label(frame, text="Body:", width=9).grid(row=1, column=0, sticky="nw")
        body_txt = tk.Text(frame, height=3, wrap="word")
        body_txt.grid(row=1, column=1, sticky="ew", padx=6, pady=(4,0))
        frame.columnconfigure(1, weight=1)

        self.section_widgets.append((heading_var, body_txt))

    def pick_template(self):
        path = filedialog.askopenfilename(filetypes=[("PNG", "*.png")])
        if path:
            self.template_path.set(path)

    def pick_out_png(self):
        path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG", "*.png")])
        if path:
            self.out_png.set(path)

    def pick_out_pdf(self):
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if path:
            self.out_pdf.set(path)

    def _collect(self):
        template = self.template_path.get().strip()
        if not template or not os.path.exists(template):
            raise ValueError("Selecione um template PNG válido.")

        title_small = self.title_small.get().strip()
        title_main = self.title_main.get().strip()
        intro = self.intro_txt.get("1.0", "end").strip()

        sections = []
        for hv, bt in self.section_widgets:
            h = hv.get().strip()
            b = bt.get("1.0", "end").strip()
            sections.append((h, b))

        return template, title_small, title_main, intro, sections

    def preview(self):
        try:
            template, t1, t2, intro, sections = self._collect()
            img = render_to_image(
                template, t1, t2, intro, sections,
                justify=self.justify.get(),
                center_vertically=self.center_vert.get()
            )

            # reduz para preview
            pv_w = 380
            scale = pv_w / img.width
            pv_h = int(img.height * scale)
            pv = img.convert("RGB").resize((pv_w, pv_h), Image.LANCZOS)
            self._preview_imgtk = ImageTk.PhotoImage(pv)
            self.preview_label.configure(image=self._preview_imgtk)

        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def generate(self):
        try:
            template, t1, t2, intro, sections = self._collect()
            img = render_to_image(
                template, t1, t2, intro, sections,
                justify=self.justify.get(),
                center_vertically=self.center_vert.get()
            )

            out_png = self.out_png.get().strip() or "saida.png"
            img.convert("RGB").save(out_png, quality=95)

            if self.make_pdf.get():
                out_pdf = self.out_pdf.get().strip() or "saida.pdf"
                save_compact_pdf(
                    img,
                    out_pdf,
                    max_width=int(self.pdf_width.get()),
                    jpeg_quality=int(self.pdf_quality.get())
                )

            messagebox.showinfo("OK", "Arquivos gerados com sucesso.")

        except Exception as e:
            messagebox.showerror("Erro", str(e))


if __name__ == "__main__":
    App().mainloop()
