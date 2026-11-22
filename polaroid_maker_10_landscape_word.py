"""
A4 Polaroid Template Maker
Requirements: Python 3.10+, Pillow, and pywin32 (for direct Windows printing)
Install packages: python -m pip install pillow pywin32
Run: python polaroid_maker.py
"""
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, colorchooser, messagebox
import subprocess
import sys
from PIL import Image, ImageOps, ImageTk, ImageDraw
from pathlib import Path

# A4 landscape at 150 DPI (11.69 x 8.27 inches).
DPI = 150
PAGE_W, PAGE_H = 1754, 1240
PREVIEW_W = 1120
SCALE = PREVIEW_W / PAGE_W
PREVIEW_H = int(PAGE_H * SCALE)

# Ten Polaroids: 5 columns x 2 rows. Each frame is exactly
# 2.1 inches wide x 3.4 inches high at the export resolution.
COLS, ROWS = 5, 2
CARD_W, CARD_H = round(2.1 * DPI), round(3.4 * DPI)
# Photo window inside each frame; the larger lower border is the caption area.
PHOTO_W, PHOTO_H = round(1.8 * DPI), round(2.6 * DPI)
PHOTO_TOP = round(0.12 * DPI)
GAP_X = 15
GAP_Y = 50
MARGIN_X = (PAGE_W - (COLS * CARD_W + (COLS - 1) * GAP_X)) // 2
MARGIN_Y = (PAGE_H - (ROWS * CARD_H + (ROWS - 1) * GAP_Y)) // 2

class Polaroid:
    def __init__(self, x, y):
        self.x, self.y = x, y
        self.image = None
        self.angle = 0
        self.frame_color = "#ffffff"
        self.visible = True

class PolaroidMaker:
    def __init__(self, root):
        self.root = root
        root.title("A4 Polaroid Template Maker")
        # Start the application maximized/full-screen while retaining the
        # normal Windows title bar and controls.
        try:
            root.state("zoomed")
        except tk.TclError:
            root.geometry("1280x900")
        root.minsize(1020, 760)
        root.configure(bg="#E8E8E8")
        self.cards = [
            Polaroid(MARGIN_X + c * (CARD_W + GAP_X), MARGIN_Y + r * (CARD_H + GAP_Y))
            for r in range(ROWS) for c in range(COLS)
        ]
        self.selected = 0
        self.page_color = "#FFFFFF"
        self.page_fill_mode = "solid"  # solid, gradient, pattern
        self.page_gradient_colors = ("#FFFFFF", "#FCE8E8")
        self.page_pattern = "dots"
        self.tk_page = None
        # Undo/Redo stacks
        self.undo_stack = []
        self.redo_stack = []
        self.max_undo = 50
        self.build_ui()
        self.redraw()

    def build_ui(self):
        # J&V Printing Services — professional red/black interface.
        self.ui = {
            "red": "#C01818",
            "red_hover": "#A01414",
            "red_pressed": "#801010",
            "dark_red": "#500808",
            "text": "#1A1A1A",
            "muted": "#555555",
            "page": "#E8E8E8",
            "white": "#FFFFFF",
            "line": "#CCCCCC",
            "workspace": "#F0F0F0",
            "button_bg": "#FCE8E8",
            "button_hover": "#F8D0D0",
            "button_pressed": "#F0B8B8",
            "button_text": "#701A1A",
            "red_soft": "#FCE8E8",
            "red_border": "#E0A8A8",
            "danger": "#E5484D",
            "danger_hover": "#D63A40",
        }
        self.root.configure(bg=self.ui["page"])

        # Thin red top rail.
        tk.Frame(self.root, bg=self.ui["red"], height=6).pack(fill="x")

        # ------------------------------------------------------------------
        # BRAND HEADER — closely follows the supplied reference image.
        # ------------------------------------------------------------------
        header = tk.Frame(self.root, bg=self.ui["white"], height=150)
        header.pack(fill="x")
        header.pack_propagate(False)

        # Soft decorative red background sweep.
        art = tk.Canvas(
            header, bg=self.ui["white"], highlightthickness=0, bd=0
        )
        art.place(relx=0.45, rely=0, relwidth=0.55, relheight=1.0)

        w = max(self.root.winfo_width(), 1200)
        art.create_oval(
            520, -330, 1180, 220,
            fill="#FCE8E8", outline=""
        )
        art.create_oval(
            680, -270, 1240, 175,
            fill="#F8D0D0", outline=""
        )
        art.create_oval(
            820, -220, 1320, 140,
            fill="#F0B8B8", outline=""
        )

        header_content = tk.Frame(header, bg=self.ui["white"])
        header_content.place(x=46, y=19, relwidth=0.94, height=112)

        # Printer/photo brand mark recreated with Tkinter shapes.
        logo = tk.Canvas(
            header_content, width=164, height=112,
            bg=self.ui["white"], highlightthickness=0, bd=0
        )
        logo.pack(side="left", padx=(0, 36))

        # Printer body.
        logo.create_round_rect = None
        logo.create_rectangle(25, 38, 128, 83, fill="#C01818", outline="")
        logo.create_rectangle(42, 17, 111, 50, fill="#E02020", outline="")
        logo.create_rectangle(51, 24, 103, 43, fill="#FCE8E8", outline="")
        logo.create_rectangle(38, 65, 115, 93, fill="#FFFFFF", outline="")
        logo.create_rectangle(49, 72, 85, 80, fill="#C01818", outline="")

        # Small photo cards.
        logo.create_polygon(
            [87, 54, 133, 47, 145, 91, 98, 99],
            fill="#FFFFFF", outline="#C01818", width=3
        )
        logo.create_oval(105, 60, 119, 74, fill="#E02020", outline="")
        logo.create_polygon(
            [102, 88, 114, 76, 122, 83, 129, 72, 140, 88],
            fill="#E02020", outline=""
        )

        # Vertical brand divider.
        tk.Frame(
            header_content, bg="#E06060", width=3, height=108
        ).pack(side="left", padx=(0, 34))

        tk.Label(
            header_content,
            text="J&V Printing Services",
            bg=self.ui["white"],
            fg="#500808",
            font=("Segoe UI", 27, "bold")
        ).pack(side="left", anchor="center")

        # ------------------------------------------------------------------
        # MAIN CONTENT — LEFT EDITING SIDEBAR + LARGE WORKSPACE
        # ------------------------------------------------------------------
        main_area = tk.Frame(
            self.root,
            bg=self.ui["page"]
        )
        main_area.pack(
            fill="both", expand=True,
            padx=12, pady=(10, 8)
        )

        # ================================================================
        # LEFT EDITING TOOLS SIDEBAR — top to bottom
        # ================================================================
        sidebar = tk.Frame(
            main_area,
            bg=self.ui["white"],
            width=275,
            highlightthickness=1,
            highlightbackground="#CCCCCC"
        )
        sidebar.pack(
            side="left",
            fill="y",
            padx=(0, 10)
        )
        sidebar.pack_propagate(False)

        # Sidebar header
        side_title = tk.Frame(
            sidebar,
            bg=self.ui["white"]
        )
        side_title.pack(fill="x", padx=16, pady=(18, 8))

        gear = tk.Canvas(
            side_title,
            width=36, height=36,
            bg=self.ui["white"],
            highlightthickness=0, bd=0
        )
        gear.pack(side="left", padx=(0, 9))
        gear.create_oval(7, 7, 29, 29, fill=self.ui["red"], outline="")
        gear.create_oval(14, 14, 22, 22, fill=self.ui["white"], outline="")
        for x1, y1, x2, y2 in [
            (17, 0, 19, 7), (17, 29, 19, 36),
            (0, 17, 7, 19), (29, 17, 36, 19)
        ]:
            gear.create_rectangle(
                x1, y1, x2, y2,
                fill=self.ui["red"], outline=""
            )

        tk.Label(
            side_title,
            text="EDITING\nTOOLS",
            justify="left",
            bg=self.ui["white"],
            fg=self.ui["red"],
            font=("Segoe UI", 14, "bold")
        ).pack(side="left")

        tk.Frame(
            sidebar,
            bg=self.ui["line"],
            height=1
        ).pack(fill="x", padx=16, pady=(2, 10))

        # Scrollable vertical button area so every tool remains accessible
        # even on smaller displays.
        tool_canvas = tk.Canvas(
            sidebar,
            bg=self.ui["white"],
            highlightthickness=0,
            bd=0
        )
        tool_scroll = ttk.Scrollbar(
            sidebar,
            orient="vertical",
            command=tool_canvas.yview
        )
        tool_frame = tk.Frame(
            tool_canvas,
            bg=self.ui["white"]
        )

        tool_window = tool_canvas.create_window(
            (0, 0),
            window=tool_frame,
            anchor="nw"
        )
        tool_canvas.configure(yscrollcommand=tool_scroll.set)

        tool_canvas.pack(
            side="left", fill="both", expand=True,
            padx=(10, 0), pady=(0, 10)
        )
        tool_scroll.pack(
            side="right", fill="y",
            padx=(0, 6), pady=(0, 10)
        )

        def update_tool_scroll(_event=None):
            tool_canvas.configure(scrollregion=tool_canvas.bbox("all"))
            tool_canvas.itemconfigure(
                tool_window,
                width=max(1, tool_canvas.winfo_width())
            )

        tool_frame.bind("<Configure>", update_tool_scroll)
        tool_canvas.bind("<Configure>", update_tool_scroll)

        # Every editing tool is now stacked vertically, top to bottom.
        vertical_tools = [
            ("↶  Undo", self.undo, False),
            ("↷  Redo", self.redo, False),
            ("▣  Add / Replace Image", self.add_image, False),
            ("↶  Rotate Left", lambda: self.rotate(-90), False),
            ("↷  Rotate Right", lambda: self.rotate(90), False),
            ("●  Frame Color", self.choose_frame_color, False),
            ("●  All Frame Color", self.choose_all_frame_color, False),
            ("◯  Page Color", self.choose_page_color, False),
            ("◆  Page Fill Style", self.open_page_fill_dialog, False),
            ("×  Remove Selected", self.remove_selected, False),
            ("↺  Restore All", self.restore_all, False),
            ("⇩  Export", self.export_pdf, False),
            ("🖨  Print", self.open_print_dialog, True),
        ]

        self.toolbar_buttons = []
        for label, command, accent in vertical_tools:
            button = self.add_button(
                tool_frame,
                label,
                command,
                accent=accent,
                width=236
            )
            # Force each button to occupy its own row.
            button.pack_forget()
            button.pack(
                side="top",
                fill="x",
                padx=2,
                pady=3
            )
            self.toolbar_buttons.append(button)

        # ================================================================
        # LARGE POLAROID WORKSPACE — right side
        # ================================================================
        holder = tk.Frame(
            main_area,
            bg=self.ui["workspace"],
            padx=10, pady=8
        )
        holder.pack(
            side="left",
            fill="both",
            expand=True
        )

        canvas_wrap = tk.Frame(
            holder,
            bg=self.ui["white"],
            highlightthickness=1,
            highlightbackground="#CCCCCC"
        )
        canvas_wrap.pack(
            expand=True,
            fill="both"
        )

        self.canvas = tk.Canvas(
            canvas_wrap,
            width=PREVIEW_W,
            height=PREVIEW_H,
            bg=self.ui["white"],
            highlightthickness=0,
            bd=0
        )
        self.canvas.pack(
            expand=True,
            padx=10, pady=10
        )
        self.canvas.bind("<Button-1>", self.on_click)

        # Small bottom status bar, intentionally subtle.
        status_bar = tk.Frame(self.root, bg=self.ui["white"], height=30)
        status_bar.pack(fill="x", side="bottom")
        status_bar.pack_propagate(False)

        self.status = tk.Label(
            status_bar,
            text="Ready",
            anchor="w",
            bg=self.ui["white"],
            fg=self.ui["muted"],
            font=("Segoe UI", 8)
        )
        self.status.pack(side="left", padx=18)

        tk.Label(
            status_bar,
            text="J&V Printing Services  •  A4 Landscape",
            bg=self.ui["white"],
            fg=self.ui["red"],
            font=("Segoe UI", 8, "bold")
        ).pack(side="right", padx=18)

    def add_button(self, parent, text, command, accent=False, width=None, icon_only=False):
        """Large rounded reference-style red/white button."""
        bg = self.ui["red"] if accent else self.ui["button_bg"]
        fg = "#FFFFFF" if accent else self.ui["button_text"]
        hover_bg = self.ui["red_hover"] if accent else self.ui["button_hover"]
        pressed_bg = self.ui["red_pressed"] if accent else self.ui["button_pressed"]
        edge = self.ui["red"] if accent else self.ui["red_border"]

        font = tkfont.Font(
            family="Segoe UI",
            size=10 if not icon_only else 13,
            weight="bold"
        )
        button_width = width or max(100, font.measure(text) + 42)
        height = 52
        radius = 12

        button = tk.Canvas(
            parent,
            width=button_width,
            height=height,
            bg="#FFFFFF",
            highlightthickness=0,
            bd=0,
            cursor="hand2"
        )
        button.pack(side="left", padx=5)

        def rounded_rect(x1, y1, x2, y2, r, fill, outline=None):
            points = [
                x1+r, y1, x2-r, y1,
                x2, y1, x2, y1+r,
                x2, y2-r, x2, y2,
                x2-r, y2, x1+r, y2,
                x1, y2, x1, y2-r,
                x1, y1+r, x1, y1
            ]
            return button.create_polygon(
                points, smooth=True, splinesteps=28,
                fill=fill, outline=outline or fill
            )

        # Soft shadow/depth.
        rounded_rect(
            1, 4, button_width-1, height-1, radius,
            "#F8D0D0" if not accent else "#E0A8A8"
        )

        shape = rounded_rect(
            1, 1, button_width-1, height-4, radius,
            bg, edge
        )

        label = button.create_text(
            button_width / 2,
            height / 2 - 1,
            text=text,
            fill=fg,
            font=font
        )

        def set_state(color, label_color):
            button.itemconfig(shape, fill=color, outline=color)
            button.itemconfig(label, fill=label_color)

        def on_enter(_event):
            set_state(hover_bg, "#FFFFFF" if accent else self.ui["dark_red"])

        def on_leave(_event):
            set_state(bg, fg)

        def on_press(_event):
            button.move(shape, 0, 2)
            button.move(label, 0, 2)
            set_state(pressed_bg, "#FFFFFF" if accent else self.ui["dark_red"])

        def on_release(event):
            button.move(shape, 0, -2)
            button.move(label, 0, -2)
            inside = 0 <= event.x <= button_width and 0 <= event.y <= height
            if inside:
                set_state(hover_bg, "#FFFFFF" if accent else self.ui["dark_red"])
                command()
            else:
                set_state(bg, fg)

        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)
        button.bind("<ButtonPress-1>", on_press)
        button.bind("<ButtonRelease-1>", on_release)

        return button

    def render_page(self, show_selection=False):
        # Create page with background fill based on mode
        if self.page_fill_mode == "gradient":
            page = self.create_gradient_background()
        elif self.page_fill_mode == "pattern":
            page = self.create_pattern_background()
        else:
            page = Image.new("RGB", (PAGE_W, PAGE_H), self.page_color)

        draw = ImageDraw.Draw(page)
        for i, card in enumerate(self.cards):
            if not card.visible:
                continue
            x, y = card.x, card.y
            # Clean frame with a dashed cut guide; no drop shadow.
            draw.rectangle((x, y, x+CARD_W-1, y+CARD_H-1), fill=card.frame_color,
                           outline="#a1a1aa", width=1)
            self.draw_cut_guides(draw, x, y, CARD_W-1, CARD_H-1)
            if card.image:
                img = ImageOps.exif_transpose(card.image.copy()).convert("RGB")
                img = img.rotate(card.angle, expand=True, resample=Image.Resampling.BICUBIC)
                # Cover-crop the rotated photo to the photo window.
                img = ImageOps.fit(img, (PHOTO_W, PHOTO_H), method=Image.Resampling.LANCZOS,
                                   centering=(0.5, 0.5))
                page.paste(img, (x + (CARD_W-PHOTO_W)//2, y + PHOTO_TOP))
            else:
                px, py = x + (CARD_W-PHOTO_W)//2, y + PHOTO_TOP
                draw.rectangle((px, py, px+PHOTO_W, py+PHOTO_H), fill="#e8edf2")
                # simple plus sign placeholder
                cx, cy = px + PHOTO_W//2, py + PHOTO_H//2
                draw.line((cx-28, cy, cx+28, cy), fill="#94a3b8", width=8)
                draw.line((cx, cy-28, cx, cy+28), fill="#94a3b8", width=8)
            # selected outline
            if show_selection and i == self.selected:
                draw.rectangle((x-5, y-5, x+CARD_W+5, y+CARD_H+5), outline="#2563eb", width=7)
        return page

    def create_gradient_background(self):
        """Create a linear gradient background."""
        page = Image.new("RGB", (PAGE_W, PAGE_H))
        c1 = self.hex_to_rgb(self.page_gradient_colors[0])
        c2 = self.hex_to_rgb(self.page_gradient_colors[1])
        for y in range(PAGE_H):
            ratio = y / PAGE_H
            r = int(c1[0] * (1 - ratio) + c2[0] * ratio)
            g = int(c1[1] * (1 - ratio) + c2[1] * ratio)
            b = int(c1[2] * (1 - ratio) + c2[2] * ratio)
            for x in range(PAGE_W):
                page.putpixel((x, y), (r, g, b))
        return page

    def create_pattern_background(self):
        """Create a pattern background."""
        page = Image.new("RGB", (PAGE_W, PAGE_H), self.page_color)
        draw = ImageDraw.Draw(page)
        pattern_color = self.adjust_color_brightness(self.page_color, 0.85)

        if self.page_pattern == "dots":
            spacing = 40
            radius = 3
            for y in range(0, PAGE_H, spacing):
                for x in range(0, PAGE_W, spacing):
                    draw.ellipse([x-radius, y-radius, x+radius, y+radius], fill=pattern_color)

        elif self.page_pattern == "grid":
            spacing = 60
            for x in range(0, PAGE_W, spacing):
                draw.line([(x, 0), (x, PAGE_H)], fill=pattern_color, width=1)
            for y in range(0, PAGE_H, spacing):
                draw.line([(0, y), (PAGE_W, y)], fill=pattern_color, width=1)

        elif self.page_pattern == "diagonal":
            spacing = 50
            for i in range(-PAGE_H, PAGE_W + PAGE_H, spacing):
                draw.line([(i, 0), (i + PAGE_H, PAGE_H)], fill=pattern_color, width=1)

        elif self.page_pattern == "crosshatch":
            spacing = 50
            for i in range(-PAGE_H, PAGE_W + PAGE_H, spacing):
                draw.line([(i, 0), (i + PAGE_H, PAGE_H)], fill=pattern_color, width=1)
                draw.line([(i + PAGE_H, 0), (i, PAGE_H)], fill=pattern_color, width=1)

        return page

    def hex_to_rgb(self, hex_color):
        """Convert hex color to RGB tuple."""
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

    def adjust_color_brightness(self, hex_color, factor):
        """Adjust color brightness."""
        rgb = self.hex_to_rgb(hex_color)
        return tuple(int(c * factor) for c in rgb)

    def draw_cut_guides(self, draw, x, y, right, bottom):
        """Draw dashed cutting guides around a Polaroid frame."""
        dash, gap = 12, 7
        # Top and bottom edges
        for start in range(x, x + right, dash + gap):
            end = min(start + dash, x + right)
            draw.line((start, y, end, y), fill="#555555", width=2)
            draw.line((start, y + bottom, end, y + bottom), fill="#555555", width=2)
        # Left and right edges
        for start in range(y, y + bottom, dash + gap):
            end = min(start + dash, y + bottom)
            draw.line((x, start, x, end), fill="#555555", width=2)
            draw.line((x + right, start, x + right, end), fill="#555555", width=2)

    def choose_all_frame_color(self):
        color = colorchooser.askcolor(
            color=self.cards[self.selected].frame_color,
            title="Choose color for all Polaroid frames"
        )[1]
        if color:
            self.save_state()
            for card in self.cards:
                if card.visible:
                    card.frame_color = color
            self.redraw()

    def export_word(self):
        """Export the visible layout as a print-ready A4 landscape DOCX."""
        path = filedialog.asksaveasfilename(
            title="Save A4 Polaroid Word document",
            defaultextension=".docx",
            filetypes=[("Word document", "*.docx")]
        )
        if not path:
            return

        try:
            import zipfile
            import io

            # Use the same print rendering as PDF: no editor selection border/X controls.
            rendered = self.render_page(show_selection=False)
            image_buffer = io.BytesIO()
            rendered.save(image_buffer, format="PNG", dpi=(DPI, DPI))
            image_data = image_buffer.getvalue()

            # Verify that the page image is valid and contains visible artwork.
            check_image = Image.open(io.BytesIO(image_data)).convert("RGB")
            if check_image.getbbox() is None:
                raise ValueError("The rendered page is empty. Restore at least one frame and try again.")

            # A4 landscape in twips, with 0.10-inch margins to prevent Word
            # from moving a full-page inline image onto a blank/extra page.
            page_w, page_h = 16838, 11906
            margin = 144
            content_w, content_h = page_w - 2 * margin, page_h - 2 * margin
            cx = int(content_w / 1440 * 914400)
            cy = int(content_h / 1440 * 914400)

            content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 <Default Extension="xml" ContentType="application/xml"/>
 <Default Extension="png" ContentType="image/png"/>
 <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

            package_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

            document_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/page.png"/>
</Relationships>"""

            document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
 <w:body>
  <w:p>
   <w:pPr>
    <w:jc w:val="center"/>
    <w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/>
    <w:ind w:left="0" w:right="0"/>
   </w:pPr>
   <w:r><w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0">
     <wp:extent cx="{cx}" cy="{cy}"/>
     <wp:effectExtent l="0" t="0" r="0" b="0"/>
     <wp:docPr id="1" name="A4 Polaroid Print Layout"/>
     <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
     <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
      <pic:pic>
       <pic:nvPicPr><pic:cNvPr id="0" name="Polaroid Print Layout"/><pic:cNvPicPr/></pic:nvPicPr>
       <pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
       <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
      </pic:pic>
     </a:graphicData></a:graphic>
    </wp:inline>
   </w:drawing></w:r>
  </w:p>
  <w:sectPr>
   <w:pgSz w:w="{page_w}" w:h="{page_h}" w:orient="landscape"/>
   <w:pgMar w:top="{margin}" w:right="{margin}" w:bottom="{margin}" w:left="{margin}" w:header="0" w:footer="0" w:gutter="0"/>
  </w:sectPr>
 </w:body>
</w:document>"""

            with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as docx:
                docx.writestr("[Content_Types].xml", content_types)
                docx.writestr("_rels/.rels", package_rels)
                docx.writestr("word/document.xml", document_xml)
                docx.writestr("word/_rels/document.xml.rels", document_rels)
                docx.writestr("word/media/page.png", image_data)

            # Validate the saved package and embedded image before reporting success.
            with zipfile.ZipFile(path, "r") as docx:
                embedded = Image.open(io.BytesIO(docx.read("word/media/page.png")))
                if embedded.size != (PAGE_W, PAGE_H):
                    raise ValueError("The embedded Word page image failed validation.")

            messagebox.showinfo(
                "Word export complete",
                "The Polaroid layout has been embedded in the Word document. "
                "It is formatted for A4 landscape printing."
            )
        except Exception as exc:
            messagebox.showerror("Word export failed", str(exc))

    def redraw(self):
        page = self.render_page(show_selection=True)
        preview = page.resize((PREVIEW_W, PREVIEW_H), Image.Resampling.LANCZOS)
        self.tk_page = ImageTk.PhotoImage(preview)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, image=self.tk_page, anchor="nw")

        # Editor-only remove controls. They are canvas widgets, not part of
        # render_page(), so they never appear in PDF, PNG, or Word exports.
        self.remove_hitboxes = {}
        for i, card in enumerate(self.cards):
            if not card.visible:
                continue
            bx = (card.x + CARD_W - 22) * SCALE
            by = (card.y + 22) * SCALE
            radius = 10
            circle = self.canvas.create_oval(
                bx-radius, by-radius, bx+radius, by+radius,
                fill=self.ui["danger"], outline="#FFFFFF", width=2
            )
            cross = self.canvas.create_text(
                bx, by-1, text="×", fill="#FFFFFF",
                font=("Segoe UI", 12, "bold")
            )
            self.remove_hitboxes[i] = (bx-radius-3, by-radius-3,
                                       bx+radius+3, by+radius+3)

        visible_count = sum(card.visible for card in self.cards)
        if not self.cards[self.selected].visible:
            visible_indices = [i for i, card in enumerate(self.cards) if card.visible]
            if visible_indices:
                self.selected = visible_indices[0]
        selected_card = self.cards[self.selected]
        self.status.config(
            text=f"Selected Polaroid {self.selected+1}  •  "
                 f"{'Image loaded' if selected_card.image else 'Empty photo slot'}  •  "
                 f"{visible_count}/10 frames visible. Click the red × to remove a frame."
        )

    def on_click(self, event):
        # First check the editor-only red X controls.
        for i, bounds in getattr(self, "remove_hitboxes", {}).items():
            x1, y1, x2, y2 = bounds
            if x1 <= event.x <= x2 and y1 <= event.y <= y2:
                self.selected = i
                self.remove_card(i)
                return

        x, y = event.x / SCALE, event.y / SCALE
        for i, card in enumerate(self.cards):
            if not card.visible:
                continue
            if card.x <= x <= card.x+CARD_W and card.y <= y <= card.y+CARD_H:
                self.selected = i
                photo_x1 = card.x + (CARD_W-PHOTO_W)//2
                photo_y1 = card.y + PHOTO_TOP
                if photo_x1 <= x <= photo_x1+PHOTO_W and photo_y1 <= y <= photo_y1+PHOTO_H:
                    self.add_image()
                else:
                    self.redraw()
                return

    def remove_card(self, index):
        if not self.cards[index].visible:
            return
        self.save_state()
        self.cards[index].visible = False
        remaining = [i for i, card in enumerate(self.cards) if card.visible]
        if remaining:
            self.selected = remaining[0] if index == self.selected else self.selected
        self.redraw()

    def remove_selected(self):
        self.save_state()
        self.remove_card(self.selected)

    def restore_all(self):
        self.save_state()
        for card in self.cards:
            card.visible = True
        self.redraw()

    def save_state(self):
        """Save current state to undo stack."""
        state = {
            "cards": [
                {
                    "image": card.image,
                    "angle": card.angle,
                    "frame_color": card.frame_color,
                    "visible": card.visible,
                }
                for card in self.cards
            ],
            "selected": self.selected,
            "page_color": self.page_color,
            "page_fill_mode": self.page_fill_mode,
            "page_gradient_colors": self.page_gradient_colors,
            "page_pattern": self.page_pattern,
        }
        self.undo_stack.append(state)
        if len(self.undo_stack) > self.max_undo:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def undo(self):
        """Undo the last action."""
        if not self.undo_stack:
            return
        # Save current state to redo stack
        current_state = {
            "cards": [
                {
                    "image": card.image,
                    "angle": card.angle,
                    "frame_color": card.frame_color,
                    "visible": card.visible,
                }
                for card in self.cards
            ],
            "selected": self.selected,
            "page_color": self.page_color,
            "page_fill_mode": self.page_fill_mode,
            "page_gradient_colors": self.page_gradient_colors,
            "page_pattern": self.page_pattern,
        }
        self.redo_stack.append(current_state)

        # Restore previous state
        state = self.undo_stack.pop()
        self._restore_state(state)

    def redo(self):
        """Redo the last undone action."""
        if not self.redo_stack:
            return
        # Save current state to undo stack
        current_state = {
            "cards": [
                {
                    "image": card.image,
                    "angle": card.angle,
                    "frame_color": card.frame_color,
                    "visible": card.visible,
                }
                for card in self.cards
            ],
            "selected": self.selected,
            "page_color": self.page_color,
            "page_fill_mode": self.page_fill_mode,
            "page_gradient_colors": self.page_gradient_colors,
            "page_pattern": self.page_pattern,
        }
        self.undo_stack.append(current_state)

        # Restore next state
        state = self.redo_stack.pop()
        self._restore_state(state)

    def _restore_state(self, state):
        """Restore state from a saved state dict."""
        for i, card_state in enumerate(state["cards"]):
            self.cards[i].image = card_state["image"]
            self.cards[i].angle = card_state["angle"]
            self.cards[i].frame_color = card_state["frame_color"]
            self.cards[i].visible = card_state["visible"]
        self.selected = state["selected"]
        self.page_color = state["page_color"]
        self.page_fill_mode = state["page_fill_mode"]
        self.page_gradient_colors = state["page_gradient_colors"]
        self.page_pattern = state["page_pattern"]
        self.redraw()

    def add_image(self):
        self.save_state()
        path = filedialog.askopenfilename(
            title="Choose a photo",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.webp *.tif"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            img = Image.open(path)
            ImageOps.exif_transpose(img).load()
            self.cards[self.selected].image = img.copy()
            self.cards[self.selected].angle = 0
            self.redraw()
        except Exception as exc:
            messagebox.showerror("Could not open image", str(exc))

    def rotate(self, degrees):
        card = self.cards[self.selected]
        if card.image is None:
            messagebox.showinfo("No image", "Add an image to the selected Polaroid first.")
            return
        self.save_state()
        card.angle = (card.angle + degrees) % 360
        self.redraw()

    def choose_frame_color(self):
        color = colorchooser.askcolor(color=self.cards[self.selected].frame_color,
                                      title="Choose Polaroid frame color")[1]
        if color:
            self.save_state()
            self.cards[self.selected].frame_color = color
            self.redraw()

    def choose_page_color(self):
        color = colorchooser.askcolor(color=self.page_color, title="Choose A4 page color")[1]
        if color:
            self.save_state()
            self.page_color = color
            self.redraw()

    def open_page_fill_dialog(self):
        """Open dialog to select page background fill style."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Page Background Fill")
        dialog.configure(bg=self.ui["page"])
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        panel = tk.Frame(
            dialog,
            bg="#FFFFFF",
            padx=24,
            pady=22,
            highlightthickness=1,
            highlightbackground="#CCCCCC"
        )
        panel.pack(fill="both", expand=True, padx=14, pady=14)

        tk.Label(
            panel,
            text="PAGE BACKGROUND FILL",
            bg="#FFFFFF",
            fg="#C01818",
            font=("Segoe UI", 15, "bold")
        ).pack(anchor="w")

        tk.Label(
            panel,
            text="Choose the background fill style for the A4 page.",
            bg="#FFFFFF",
            fg="#555555",
            font=("Segoe UI", 9)
        ).pack(anchor="w", pady=(4, 16))

        # Fill mode selector
        tk.Label(
            panel,
            text="FILL STYLE",
            bg="#FFFFFF",
            fg="#333333",
            font=("Segoe UI", 8, "bold")
        ).pack(anchor="w")

        fill_var = tk.StringVar(value=self.page_fill_mode)
        fill_options = [
            ("solid", "Solid Color"),
            ("gradient", "Gradient"),
            ("pattern", "Pattern"),
        ]

        fill_frame = tk.Frame(panel, bg="#FFFFFF")
        fill_frame.pack(fill="x", pady=(5, 14))

        for value, label in fill_options:
            rb = tk.Radiobutton(
                fill_frame,
                text=label,
                variable=fill_var,
                value=value,
                bg="#FFFFFF",
                fg="#333333",
                font=("Segoe UI", 9),
                activebackground="#FFFFFF",
                activeforeground="#C01818",
                selectcolor="#FCE8E8",
                command=lambda: self.update_fill_options(fill_var.get(), options_frame)
            )
            rb.pack(anchor="w", pady=2)

        # Dynamic options frame
        options_frame = tk.Frame(panel, bg="#FFFFFF")
        options_frame.pack(fill="x", pady=(0, 14))

        # Initialize options
        self.update_fill_options(self.page_fill_mode, options_frame)

        # Buttons
        buttons = tk.Frame(panel, bg="#FFFFFF")
        buttons.pack(fill="x")

        def apply_and_close():
            self.save_state()
            self.page_fill_mode = fill_var.get()
            self.redraw()
            dialog.destroy()

        action_row = tk.Frame(buttons, bg="#FFFFFF")
        action_row.pack(fill="x")

        self.add_button(
            action_row,
            "Cancel",
            dialog.destroy,
            width=90
        )
        self.add_button(
            action_row,
            "Apply",
            apply_and_close,
            accent=True,
            width=120
        )

        dialog.update_idletasks()
        x = self.root.winfo_rootx() + (
            self.root.winfo_width() - dialog.winfo_width()
        ) // 2
        y = self.root.winfo_rooty() + (
            self.root.winfo_height() - dialog.winfo_height()
        ) // 2
        dialog.geometry(
            f"+{max(0, x)}+{max(0, y)}"
        )

    def update_fill_options(self, mode, parent):
        """Update the options frame based on selected fill mode."""
        # Clear existing widgets
        for widget in parent.winfo_children():
            widget.destroy()

        if mode == "solid":
            tk.Label(
                parent,
                text="Solid color uses the Page Color setting.",
                bg="#FFFFFF",
                fg="#555555",
                font=("Segoe UI", 9)
            ).pack(anchor="w")

        elif mode == "gradient":
            tk.Label(
                parent,
                text="GRADIENT COLORS",
                bg="#FFFFFF",
                fg="#333333",
                font=("Segoe UI", 8, "bold")
            ).pack(anchor="w", pady=(8, 4))

            grad_frame = tk.Frame(parent, bg="#FFFFFF")
            grad_frame.pack(fill="x", pady=(0, 8))

            # Color 1
            tk.Label(
                grad_frame,
                text="Start Color:",
                bg="#FFFFFF",
                fg="#333333",
                font=("Segoe UI", 9)
            ).pack(side="left", padx=(0, 8))

            color1_btn = tk.Button(
                grad_frame,
                text="  ",
                bg=self.page_gradient_colors[0],
                width=3,
                command=lambda: self.pick_gradient_color(0, color1_btn, color2_btn)
            )
            color1_btn.pack(side="left", padx=(0, 16))

            # Color 2
            tk.Label(
                grad_frame,
                text="End Color:",
                bg="#FFFFFF",
                fg="#333333",
                font=("Segoe UI", 9)
            ).pack(side="left", padx=(0, 8))

            color2_btn = tk.Button(
                grad_frame,
                text="  ",
                bg=self.page_gradient_colors[1],
                width=3,
                command=lambda: self.pick_gradient_color(1, color1_btn, color2_btn)
            )
            color2_btn.pack(side="left")

        elif mode == "pattern":
            tk.Label(
                parent,
                text="PATTERN TYPE",
                bg="#FFFFFF",
                fg="#333333",
                font=("Segoe UI", 8, "bold")
            ).pack(anchor="w", pady=(8, 4))

            pattern_var = tk.StringVar(value=self.page_pattern)
            patterns = [
                ("dots", "Dots"),
                ("grid", "Grid"),
                ("diagonal", "Diagonal Lines"),
                ("crosshatch", "Crosshatch"),
            ]

            for value, label in patterns:
                rb = tk.Radiobutton(
                    parent,
                    text=label,
                    variable=pattern_var,
                    value=value,
                    bg="#FFFFFF",
                    fg="#333333",
                    font=("Segoe UI", 9),
                    activebackground="#FFFFFF",
                    activeforeground="#C01818",
                    selectcolor="#FCE8E8",
                    command=lambda v=value: setattr(self, 'page_pattern', v)
                )
                rb.pack(anchor="w", pady=2)

    def pick_gradient_color(self, index, btn1, btn2):
        """Pick a gradient color."""
        color = colorchooser.askcolor(
            color=self.page_gradient_colors[index],
            title=f"Choose gradient color {index + 1}"
        )[1]
        if color:
            self.save_state()
            colors = list(self.page_gradient_colors)
            colors[index] = color
            self.page_gradient_colors = tuple(colors)
            btn1.config(bg=self.page_gradient_colors[0])
            btn2.config(bg=self.page_gradient_colors[1])

    def open_print_dialog(self):
        """Show printer selection and print preferences before printing."""
        if not sys.platform.startswith("win"):
            messagebox.showerror(
                "Printing not supported",
                "Direct printer selection is currently implemented for Windows."
            )
            return

        try:
            import win32print
            printers = [
                item[2] for item in win32print.EnumPrinters(
                    win32print.PRINTER_ENUM_LOCAL |
                    win32print.PRINTER_ENUM_CONNECTIONS
                )
            ]
        except ImportError:
            messagebox.showerror(
                "Printing package required",
                "Direct printing needs pywin32.\n\n"
                "Install it with:\npython -m pip install pywin32"
            )
            return
        except Exception as exc:
            messagebox.showerror(
                "Printer error",
                f"Could not list printers:\n{exc}"
            )
            return

        if not printers:
            messagebox.showwarning(
                "No printers found",
                "Windows did not report any installed printers."
            )
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("Print Preferences")
        dialog.configure(bg=self.ui["page"])
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        panel = tk.Frame(
            dialog,
            bg="#FFFFFF",
            padx=24,
            pady=22,
            highlightthickness=1,
            highlightbackground="#CCCCCC"
        )
        panel.pack(fill="both", expand=True, padx=14, pady=14)

        tk.Label(
            panel,
            text="PRINT YOUR POLAROID LAYOUT",
            bg="#FFFFFF",
            fg="#C01818",
            font=("Segoe UI", 15, "bold")
        ).pack(anchor="w")

        tk.Label(
            panel,
            text="Select the printer and preferences before sending the A4 layout to print.",
            bg="#FFFFFF",
            fg="#555555",
            font=("Segoe UI", 9)
        ).pack(anchor="w", pady=(4, 16))

        # ---------------------------------------------------------------
        # Printer
        # ---------------------------------------------------------------
        tk.Label(
            panel,
            text="PRINTER",
            bg="#FFFFFF",
            fg="#333333",
            font=("Segoe UI", 8, "bold")
        ).pack(anchor="w")

        printer_var = tk.StringVar()
        try:
            default_printer = win32print.GetDefaultPrinter()
        except Exception:
            default_printer = printers[0]
        printer_var.set(
            default_printer if default_printer in printers else printers[0]
        )

        printer_box = ttk.Combobox(
            panel,
            textvariable=printer_var,
            values=printers,
            state="readonly",
            width=52
        )
        printer_box.pack(fill="x", pady=(5, 14))

        # ---------------------------------------------------------------
        # Buttons
        # ---------------------------------------------------------------
        def open_preferences():
            selected_printer = printer_var.get().strip()
            if not selected_printer:
                return
            try:
                subprocess.Popen([
                    "rundll32.exe",
                    "printui.dll",
                    "PrintUIEntry",
                    "/e",
                    "/n",
                    selected_printer
                ])
            except Exception as exc:
                messagebox.showerror(
                    "Printer preferences",
                    f"Could not open printer preferences:\n{exc}",
                    parent=dialog
                )

        def do_print():
            selected_printer = printer_var.get().strip()
            if not selected_printer:
                messagebox.showwarning(
                    "Select printer",
                    "Please select a printer first.",
                    parent=dialog
                )
                return

            if orientation_var.get() != "Landscape":
                messagebox.showinfo(
                    "Orientation",
                    "The Polaroid template is designed for A4 Landscape. "
                    "The print job will use Landscape orientation.",
                    parent=dialog
                )

            # Keep the existing A4 landscape print defaults after removing
            # the preference controls from this popup.
            preferences = {
                "paper": "A4",
                "orientation": "Landscape",
                "copies": 1,
                "color": "Color",
                "quality": "High",
                "scaling": "Fit to printable area",
            }

            dialog.destroy()
            self.print_to_printer(
                selected_printer,
                preferences
            )

        buttons = tk.Frame(panel, bg="#FFFFFF")
        buttons.pack(fill="x")

        tk.Label(
            buttons,
            text="Use Printer Preferences to open the selected printer's driver settings.",
            bg="#FFFFFF",
            fg="#555555",
            font=("Segoe UI", 8),
            anchor="w"
        ).pack(fill="x", pady=(0, 8))

        action_row = tk.Frame(buttons, bg="#FFFFFF")
        action_row.pack(fill="x")

        self.add_button(
            action_row,
            "⚙  Printer Preferences",
            open_preferences,
            width=170
        )
        self.add_button(
            action_row,
            "Cancel",
            dialog.destroy,
            width=90
        )
        self.add_button(
            action_row,
            "🖨  Print",
            do_print,
            accent=True,
            width=120
        )

        dialog.update_idletasks()
        x = self.root.winfo_rootx() + (
            self.root.winfo_width() - dialog.winfo_width()
        ) // 2
        y = self.root.winfo_rooty() + (
            self.root.winfo_height() - dialog.winfo_height()
        ) // 2
        dialog.geometry(
            f"+{max(0, x)}+{max(0, y)}"
        )

    def print_to_printer(self, printer_name, preferences=None):
        """Print the current page directly through the selected Windows printer."""
        preferences = preferences or {
            "paper": "A4", "orientation": "Landscape", "copies": 1,
            "color": "Color", "quality": "High",
            "scaling": "Fit to printable area"
        }
        try:
            import win32print
            import win32ui
            from PIL import ImageWin

            page = self.render_page(show_selection=False).convert("RGB")
            printer_dc = win32ui.CreateDC()
            printer_dc.CreatePrinterDC(printer_name)

            # Use the printer's actual printable area and hardware margins.
            import win32con
            printable_w = printer_dc.GetDeviceCaps(win32con.HORZRES)
            printable_h = printer_dc.GetDeviceCaps(win32con.VERTRES)
            offset_x = printer_dc.GetDeviceCaps(win32con.PHYSICALOFFSETX)
            offset_y = printer_dc.GetDeviceCaps(win32con.PHYSICALOFFSETY)
            physical_w = printer_dc.GetDeviceCaps(win32con.PHYSICALWIDTH)
            physical_h = printer_dc.GetDeviceCaps(win32con.PHYSICALHEIGHT)

            if printable_w <= 0 or printable_h <= 0:
                raise RuntimeError("The selected printer reported an invalid printable area.")

            if preferences.get("scaling") == "Actual size":
                scale = min(
                    1.0,
                    printable_w / page.width,
                    printable_h / page.height
                )
            else:
                scale = min(
                    printable_w / page.width,
                    printable_h / page.height
                )

            draw_w = int(page.width * scale)
            draw_h = int(page.height * scale)
            left = offset_x + max(0, (printable_w - draw_w) // 2)
            top = offset_y + max(0, (printable_h - draw_h) // 2)
            right, bottom = left + draw_w, top + draw_h

            copies = max(1, int(preferences.get("copies", 1)))
            dib = ImageWin.Dib(page)

            printer_dc.StartDoc(
                f"A4 Polaroid Layout - {copies} copy"
                + ("ies" if copies != 1 else "")
            )

            for _ in range(copies):
                printer_dc.StartPage()
                dib.draw(
                    printer_dc.GetHandleOutput(),
                    (left, top, right, bottom)
                )
                printer_dc.EndPage()

            printer_dc.EndDoc()
            printer_dc.DeleteDC()

            messagebox.showinfo(
                "Print job sent",
                f"Your A4 landscape Polaroid layout was sent to:\n"
                f"{printer_name}\n\n"
                f"Copies: {copies}\n"
                f"Color: {preferences.get('color', 'Color')}\n"
                f"Quality: {preferences.get('quality', 'High')}\n"
                f"Scaling: {preferences.get('scaling', 'Fit to printable area')}"
            )
        except ImportError:
            messagebox.showerror(
                "Printing package required",
                "Direct printing needs pywin32.\n\nInstall it with:\npython -m pip install pywin32"
            )
            return
        except Exception as exc:
            messagebox.showerror("Print failed", str(exc))
            return

    def export_pdf(self):
        path = filedialog.asksaveasfilename(
            title="Save A4 Polaroid PDF", defaultextension=".pdf",
            filetypes=[("PDF file", "*.pdf"), ("PNG image", "*.png")]
        )
        if not path:
            return
        try:
            page = self.render_page()
            if path.lower().endswith(".png"):
                page.save(path, dpi=(DPI, DPI))
            else:
                # Pillow writes a print-ready A4-sized raster PDF at 150 DPI.
                page.save(path, "PDF", resolution=float(DPI))
            messagebox.showinfo("Export complete", f"Your A4 layout was saved:\n{path}")
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))

if __name__ == "__main__":
    root = tk.Tk()
    app = PolaroidMaker(root)
    root.mainloop()
