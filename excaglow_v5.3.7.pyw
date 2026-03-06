import wmi
import struct
import time
import threading
import ctypes
import sys
import math
import json
import os
import tkinter as tk
from mss import mss
from PIL import Image, ImageEnhance

# ── pystray opsiyonel ────────────────────────────────────────────
try:
    import pystray
    from pystray import MenuItem as TrayItem
    from PIL import Image as PILImage
    TRAY_OK = True
except ImportError:
    TRAY_OK = False

# ── Ayar Dosyası ─────────────────────────────────────────────────
SETTINGS_PATH = os.path.join(os.path.expanduser("~"), ".excaglow_settings.json")

def load_settings():
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}

def save_settings(data: dict):
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        print(f"Ayar kaydedilemedi: {e}")

# ── Donanım Sabitleri ────────────────────────────────────────────
ZONE_LEFT  = 3
ZONE_MID   = 4
ZONE_RIGHT = 5
ZONE_ALL   = 6
WMI_CLASS  = "RW_GMWMI"

SMOOTHING  = 0.55   # 0.0 = anlık, 1.0 = hiç değişmez — 0.55 = hızlı ama yumuşak

# ── UI Renk Paleti ───────────────────────────────────────────────
C = {
    "bg":       "#050508",
    "surface":  "#0d0d14",
    "surface2": "#13131e",
    "border":   "#1e1e30",
    "accent":   "#4f8ef7",
    "accent2":  "#7c3aed",
    "cyan":     "#00d4ff",
    "green":    "#00ff88",
    "red":      "#ff355e",
    "orange":   "#ff6b35",
    "text":     "#e2e8f0",
    "muted":    "#475569",
    "dim":      "#1e293b",
}

QUICK_COLORS = [
    (0,212,255),(0,255,136),(255,53,94),
    (255,107,53),(147,51,234),(255,220,0),(255,255,255)
]

# ── Yardımcı Fonksiyonlar ────────────────────────────────────────
def build_packet(zone, r, g, b, brightness, mode=1):
    color_data = (mode << 28) | (brightness << 24) | (r << 16) | (g << 8) | b
    return struct.pack("<HHIIIIIII", 0xFB00, 0x0100, zone, color_data, 0, 0, 0, 0, 0)

def hsv_to_rgb(h, s, v):
    h = h % 360
    c = v * s
    x = c * (1 - abs((h / 60) % 2 - 1))
    m = v - c
    if   h < 60:  r,g,b = c,x,0
    elif h < 120: r,g,b = x,c,0
    elif h < 180: r,g,b = 0,c,x
    elif h < 240: r,g,b = 0,x,c
    elif h < 300: r,g,b = x,0,c
    else:         r,g,b = c,0,x
    return int((r+m)*255), int((g+m)*255), int((b+m)*255)

def lerp(a, b, t):
    """Tek kanal lineer interpolasyon"""
    return a + (b - a) * t

def lerp_color(c1, c2, t):
    t = max(0.0, min(1.0, t))
    return (int(lerp(c1[0],c2[0],t)),
            int(lerp(c1[1],c2[1],t)),
            int(lerp(c1[2],c2[2],t)))

def smooth_color(current, target, factor=SMOOTHING):
    """
    Linear interpolation smoothing:
    current += (target - current) * factor
    """
    return (
        int(current[0] + (target[0] - current[0]) * factor),
        int(current[1] + (target[1] - current[1]) * factor),
        int(current[2] + (target[2] - current[2]) * factor),
    )

def sample_palette(colors, t):
    if len(colors) == 1:
        return colors[0]
    t = t % 1.0
    n = len(colors)
    scaled = t * n
    idx = int(scaled)
    frac = scaled - idx
    return lerp_color(colors[idx % n], colors[(idx+1) % n], frac)

def make_tray_icon():
    """16x16 basit ikon oluştur"""
    img = PILImage.new("RGB", (64, 64), (5, 5, 8))
    # Basit gradient daire
    from PIL import ImageDraw
    draw = ImageDraw.Draw(img)
    draw.ellipse([4,4,60,60], fill=(0, 212, 255))
    draw.ellipse([16,16,48,48], fill=(79, 142, 247))
    return img

# ── Renk Çarkı ───────────────────────────────────────────────────
class ColorWheelPicker(tk.Toplevel):
    def __init__(self, parent, initial=(0,212,255), callback=None):
        super().__init__(parent)
        self.title("")
        self.configure(bg=C["bg"])
        self.resizable(False, False)
        self.callback  = callback
        self.sel_color = initial
        self.sel_h     = 195.0
        self.sel_s     = 1.0
        self.bv        = tk.DoubleVar(value=1.0)
        self.SIZE      = 240
        self.CENTER    = self.SIZE // 2
        self.RADIUS    = 105
        self._build()
        self._draw_wheel()
        self.grab_set()
        self._center()

    def _center(self):
        self.update_idletasks()
        x = self.master.winfo_x() + (self.master.winfo_width()  - self.winfo_width())  // 2
        y = self.master.winfo_y() + (self.master.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _build(self):
        tk.Label(self, text="RENK SEÇİCİ", font=("Courier", 11, "bold"),
                 fg=C["cyan"], bg=C["bg"]).pack(pady=(18,8))
        self.canvas = tk.Canvas(self, width=self.SIZE, height=self.SIZE,
                                bg=C["bg"], highlightthickness=0, cursor="crosshair")
        self.canvas.pack(padx=24)
        self.canvas.bind("<Button-1>",  self._click)
        self.canvas.bind("<B1-Motion>", self._click)
        br_f = tk.Frame(self, bg=C["bg"])
        br_f.pack(fill="x", padx=24, pady=(10,4))
        tk.Label(br_f, text="PARLAKLIK", fg=C["muted"], bg=C["bg"],
                 font=("Courier", 8)).pack(anchor="w")
        tk.Scale(br_f, from_=0.0, to=1.0, resolution=0.01, orient="horizontal",
                 variable=self.bv, bg=C["surface"], fg=C["cyan"],
                 troughcolor=C["dim"], highlightthickness=0, bd=0,
                 command=lambda _: self._refresh()).pack(fill="x")
        self.prev = tk.Label(self, bg='#%02x%02x%02x' % self.sel_color,
                             width=22, height=2, relief="flat")
        self.prev.pack(pady=8)
        self.hex_lbl = tk.Label(self, text='#%02x%02x%02x' % self.sel_color,
                                fg=C["text"], bg=C["bg"], font=("Courier", 13, "bold"))
        self.hex_lbl.pack()
        bf = tk.Frame(self, bg=C["bg"])
        bf.pack(pady=14)
        tk.Button(bf, text="✔  UYGULA", bg=C["green"], fg="#000",
                  font=("Courier", 10, "bold"), relief="flat", padx=16, pady=6,
                  cursor="hand2", command=self._confirm).pack(side="left", padx=6)
        tk.Button(bf, text="✘  İPTAL", bg=C["surface2"], fg=C["muted"],
                  font=("Courier", 10), relief="flat", padx=16, pady=6,
                  cursor="hand2", command=self.destroy).pack(side="left", padx=6)

    def _draw_wheel(self):
        self.canvas.delete("all")
        img = tk.PhotoImage(width=self.SIZE, height=self.SIZE)
        cx, cy, r = self.CENTER, self.CENTER, self.RADIUS
        rows = []
        for y in range(self.SIZE):
            row = []
            for x in range(self.SIZE):
                dx, dy = x-cx, y-cy
                dist = math.hypot(dx, dy)
                if dist <= r:
                    angle = math.degrees(math.atan2(dy, dx)) % 360
                    rv, gv, bv = hsv_to_rgb(angle, dist/r, 1.0)
                    row.append(f"#{rv:02x}{gv:02x}{bv:02x}")
                else:
                    row.append(C["bg"])
            rows.append("{" + " ".join(row) + "}")
        img.put(" ".join(rows))
        self.canvas.create_image(0, 0, anchor="nw", image=img)
        self._img_ref = img
        self.canvas.create_oval(cx-r-1, cy-r-1, cx+r+1, cy+r+1,
                                outline=C["border"], width=2)
        self._sel_ring = self.canvas.create_oval(0,0,1,1, outline="white", width=2)

    def _click(self, e):
        dx, dy = e.x-self.CENTER, e.y-self.CENTER
        dist = math.hypot(dx, dy)
        if dist > self.RADIUS: return
        self.sel_h = math.degrees(math.atan2(dy, dx)) % 360
        self.sel_s = dist / self.RADIUS
        self._refresh()
        self.canvas.coords(self._sel_ring, e.x-8, e.y-8, e.x+8, e.y+8)

    def _refresh(self):
        r, g, b = hsv_to_rgb(self.sel_h, self.sel_s, self.bv.get())
        self.sel_color = (r, g, b)
        hx = '#%02x%02x%02x' % (r, g, b)
        self.prev.config(bg=hx)
        self.hex_lbl.config(text=hx.upper())

    def _confirm(self):
        if self.callback: self.callback(self.sel_color)
        self.destroy()


# ── Dalga Paleti Editörü ─────────────────────────────────────────
class WavePaletteEditor(tk.Toplevel):
    def __init__(self, parent, colors, callback):
        super().__init__(parent)
        self.title("")
        self.configure(bg=C["bg"])
        self.resizable(False, False)
        self.callback     = callback
        self.colors       = list(colors)
        self._pending_idx = None

        # grad_canvas ÖNCE oluştur
        tk.Label(self, text="DALGA PALETİ", font=("Courier", 12, "bold"),
                 fg=C["cyan"], bg=C["bg"]).pack(pady=(18,4))
        tk.Label(self, text="renklere tıkla · düzenlemek için",
                 font=("Courier", 8), fg=C["muted"], bg=C["bg"]).pack(pady=(0,10))

        self.slots_frame = tk.Frame(self, bg=C["bg"])
        self.slots_frame.pack(padx=24, pady=(0,4))

        self.grad_canvas = tk.Canvas(self, height=20, bg=C["bg"], highlightthickness=0)
        self.grad_canvas.pack(fill="x", padx=24, pady=(8,4))
        self.grad_canvas.bind("<Configure>", lambda _: self._draw_gradient())

        self._render_slots()   # grad_canvas artık mevcut

        btn_row = tk.Frame(self, bg=C["bg"])
        btn_row.pack(pady=8)
        tk.Button(btn_row, text="＋  RENK EKLE", bg=C["dim"], fg=C["cyan"],
                  font=("Courier", 9, "bold"), relief="flat", padx=12, pady=6,
                  cursor="hand2", command=self._add_color).pack(side="left", padx=4)
        tk.Button(btn_row, text="－  SON SİL", bg=C["dim"], fg=C["muted"],
                  font=("Courier", 9), relief="flat", padx=12, pady=6,
                  cursor="hand2", command=self._remove_last).pack(side="left", padx=4)

        bf = tk.Frame(self, bg=C["bg"])
        bf.pack(pady=14)
        tk.Button(bf, text="✔  UYGULA", bg=C["green"], fg="#000",
                  font=("Courier", 10, "bold"), relief="flat", padx=16, pady=6,
                  cursor="hand2", command=self._confirm).pack(side="left", padx=6)
        tk.Button(bf, text="✘  İPTAL", bg=C["surface2"], fg=C["muted"],
                  font=("Courier", 10), relief="flat", padx=16, pady=6,
                  cursor="hand2", command=self.destroy).pack(side="left", padx=6)

        self.grab_set()
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width()  - self.winfo_width())  // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _render_slots(self):
        for w in self.slots_frame.winfo_children():
            w.destroy()
        for i, col in enumerate(self.colors):
            hx = '#%02x%02x%02x' % col
            f = tk.Frame(self.slots_frame, bg=C["surface"], padx=6, pady=6)
            f.pack(side="left", padx=4)
            cv = tk.Canvas(f, width=52, height=52, bg=hx,
                           highlightthickness=2, highlightbackground=C["border"],
                           cursor="hand2")
            cv.pack()
            tk.Label(f, text=hx.upper(), font=("Courier", 7),
                     fg=C["muted"], bg=C["surface"]).pack(pady=(3,0))
            cv.bind("<Button-1>", lambda e, idx=i: self._edit_color(idx))
        self._draw_gradient()

    def _draw_gradient(self):
        self.grad_canvas.delete("all")
        w = self.grad_canvas.winfo_width()
        h = self.grad_canvas.winfo_height() or 20
        if w < 2 or not self.colors: return
        for x in range(w):
            r, g, b = sample_palette(self.colors, x / w)
            self.grad_canvas.create_line(x, 0, x, h, fill='#%02x%02x%02x'%(r,g,b))

    def _edit_color(self, idx):
        self._pending_idx = idx
        ColorWheelPicker(self, initial=self.colors[idx], callback=self._on_picked)

    def _on_picked(self, color):
        if self._pending_idx is not None:
            self.colors[self._pending_idx] = color
            self._render_slots()

    def _add_color(self):
        if len(self.colors) >= 6: return
        self._pending_idx = len(self.colors)
        self.colors.append((255, 255, 255))
        self._render_slots()
        ColorWheelPicker(self, initial=(255,255,255), callback=self._on_picked)

    def _remove_last(self):
        if len(self.colors) > 1:
            self.colors.pop()
            self._render_slots()

    def _confirm(self):
        self.callback(self.colors)
        self.destroy()


# ── Mod Kartı ────────────────────────────────────────────────────
class ModeCard(tk.Frame):
    def __init__(self, parent, label, icon, value, variable):
        super().__init__(parent, bg=C["surface"], cursor="hand2", relief="flat", bd=0)
        self.value = value
        self.var   = variable
        self.inner = tk.Frame(self, bg=C["surface"], padx=10, pady=10)
        self.inner.pack(fill="both", expand=True)
        self.icon_lbl = tk.Label(self.inner, text=icon, font=("Segoe UI Emoji", 17),
                                  bg=C["surface"], fg=C["muted"])
        self.icon_lbl.pack()
        self.text_lbl = tk.Label(self.inner, text=label, font=("Courier", 7, "bold"),
                                  bg=C["surface"], fg=C["muted"])
        self.text_lbl.pack(pady=(2,0))
        for w in [self, self.inner, self.icon_lbl, self.text_lbl]:
            w.bind("<Button-1>", self._select)
        self.var.trace_add("write", lambda *_: self._refresh())
        self._refresh()

    def _select(self, _=None): self.var.set(self.value)

    def _refresh(self):
        sel = self.var.get() == self.value
        bg  = C["dim"]    if sel else C["surface"]
        fg  = C["cyan"]   if sel else C["muted"]
        bdr = C["accent"] if sel else C["border"]
        for w in [self, self.inner, self.icon_lbl, self.text_lbl]:
            w.config(bg=bg)
        self.icon_lbl.config(fg=fg)
        self.text_lbl.config(fg=fg)
        self.config(highlightbackground=bdr,
                    highlightthickness=2 if sel else 1,
                    highlightcolor=bdr)


# ── Gradient Bar ─────────────────────────────────────────────────
class GradientBar(tk.Canvas):
    def __init__(self, parent, colors, **kw):
        h = kw.pop("height", 3)
        super().__init__(parent, height=h, highlightthickness=0, bg=C["bg"], **kw)
        self.colors = colors
        self.bind("<Configure>", self._draw)

    def _draw(self, _=None):
        self.delete("all")
        w = self.winfo_width(); h = self.winfo_height()
        segs = len(self.colors) - 1
        if segs < 1 or w < 2: return
        sw = w / segs
        for i in range(segs):
            c1, c2 = self.colors[i], self.colors[i+1]
            for x in range(int(sw)):
                t  = x / sw
                r  = int(c1[0]*(1-t)+c2[0]*t)
                g  = int(c1[1]*(1-t)+c2[1]*t)
                b  = int(c1[2]*(1-t)+c2[2]*t)
                self.create_line(int(i*sw+x), 0, int(i*sw+x), h,
                                 fill='#%02x%02x%02x'%(r,g,b))


# ── Ana Uygulama ─────────────────────────────────────────────────
class ExcaGlowApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excalibur RGB Controller")
        self.root.geometry("560x860")
        self.root.configure(bg=C["bg"])
        self.root.resizable(False, False)

        self.running   = False
        self.wmi_ok    = self._test_wmi()

        # Smooth renk durumları (her zone için ayrı)
        self._smooth = {ZONE_LEFT: (0,0,0), ZONE_MID: (0,0,0), ZONE_RIGHT: (0,0,0)}

        # Ayarları yükle
        cfg = load_settings()
        self.mode_var       = tk.StringVar(value=cfg.get("mode", "ambient"))
        self.brightness_var = tk.IntVar(value=cfg.get("brightness", 2))
        self.fps_var        = tk.IntVar(value=cfg.get("fps", 50))
        sc = cfg.get("static_color", [0, 212, 255])
        self.static_color   = tuple(sc)
        wc = cfg.get("wave_colors", [[0,212,255],[147,51,234]])
        self.wave_colors    = [tuple(c) for c in wc]
        self.smooth_var     = tk.DoubleVar(value=cfg.get("smooth", 0.55))

        # Efekt zamanları
        self._breath_t = 0.0
        self._wave_t   = 0.0
        self._cycle_h  = 0.0

        # Sistem tepsisi
        self._tray      = None
        self._tray_thread = None

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # --startup: direkt tepside başla ve ışıkları yak
        if "--startup" in sys.argv:
            self.root.withdraw()
            self._start_tray()
            self.root.after(500, self._toggle)

    # ── Ayar Kayıt ─────────────────────────────────────────────
    def _save_settings(self, *_):
        wave_speed = 1.0
        if hasattr(self, "wave_speed_var"):
            try: wave_speed = self.wave_speed_var.get()
            except: pass
        save_settings({
            "mode":         self.mode_var.get(),
            "brightness":   self.brightness_var.get(),
            "fps":          self.fps_var.get(),
            "static_color": list(self.static_color),
            "wave_colors":  [list(c) for c in self.wave_colors],
            "wave_speed":   wave_speed,
            "smooth":       self.smooth_var.get(),
        })

    def _on_close(self):
        self.running = False
        self._save_settings()
        if self._tray:
            try: self._tray.stop()
            except: pass
        self.root.destroy()

    # ── WMI ────────────────────────────────────────────────────
    def _test_wmi(self):
        try:
            c = wmi.WMI(namespace="root\\wmi")
            return len(c.query(f"SELECT * FROM {WMI_CLASS}")) > 0
        except: return False

    # ── UI ─────────────────────────────────────────────────────
    def _build_ui(self):
        # Header
        hdr = tk.Frame(self.root, bg=C["bg"])
        hdr.pack(fill="x", padx=28, pady=(28,0))
        lhdr = tk.Frame(hdr, bg=C["bg"])
        lhdr.pack(side="left")
        tk.Label(lhdr, text="EXCAGLOW", font=("Courier", 22, "bold"),
                 fg=C["text"], bg=C["bg"]).pack(anchor="w")
        tk.Label(lhdr, text="V5.3.7  ·  CASPER EDITION", font=("Courier", 9),
                 fg=C["muted"], bg=C["bg"]).pack(anchor="w")

        right_hdr = tk.Frame(hdr, bg=C["bg"])
        right_hdr.pack(side="right", anchor="n")

        pill_c = C["green"] if self.wmi_ok else C["red"]
        pill_t = "● BAĞLANDI" if self.wmi_ok else "● BAĞLANTI YOK"
        tk.Label(right_hdr, text=pill_t, font=("Courier", 8, "bold"),
                 fg=pill_c, bg=C["surface2"], padx=10, pady=4).pack(anchor="e")

        # Tepsi butonu (pystray varsa)
        if TRAY_OK:
            tk.Button(right_hdr, text="⬇ Gizle", bg=C["dim"], fg=C["muted"],
                      font=("Courier", 8), relief="flat", padx=6, pady=2,
                      cursor="hand2", command=self._hide_to_tray).pack(anchor="e", pady=(4,0))

        GradientBar(self.root, [(79,142,247),(124,58,237),(0,212,255)],
                    height=3).pack(fill="x", padx=28, pady=(14,20))

        self._build_keyboard_preview()

        tk.Frame(self.root, bg=C["border"], height=1).pack(fill="x", padx=28, pady=14)

        # Mod kartları
        tk.Label(self.root, text="MOD", font=("Courier", 9, "bold"),
                 fg=C["muted"], bg=C["bg"]).pack(anchor="w", padx=28)
        mf = tk.Frame(self.root, bg=C["bg"])
        mf.pack(fill="x", padx=28, pady=(6,0))
        modes = [
            ("AMBİANT","🌅","ambient"), ("3 BÖLGE","🎨","zones"),
            ("SABİT","💡","static"),   ("NEFES","🫁","breathe"),
            ("DALGA","🌊","wave"),      ("DÖNGÜ","🌈","cycle"),
        ]
        for i, (lbl, ico, val) in enumerate(modes):
            card = ModeCard(mf, lbl, ico, val, self.mode_var)
            card.grid(row=0, column=i, padx=4, sticky="nsew")
            mf.columnconfigure(i, weight=1)

        self.mode_var.trace_add("write",       lambda *_: (self._on_mode_change(), self._save_settings()))
        self.brightness_var.trace_add("write", self._save_settings)
        self.fps_var.trace_add("write",        self._save_settings)

        # Dinamik panel
        self.dynamic_frame = tk.Frame(self.root, bg=C["bg"])
        self.dynamic_frame.pack(fill="x", padx=28, pady=(10,0))
        self.static_panel = tk.Frame(self.dynamic_frame, bg=C["bg"])
        self._build_static_picker(self.static_panel)
        self.wave_panel = tk.Frame(self.dynamic_frame, bg=C["bg"])
        self._build_wave_panel(self.wave_panel)

        # Başlangıç paneli göster
        self._on_mode_change()

        tk.Frame(self.root, bg=C["border"], height=1).pack(fill="x", padx=28, pady=14)

        # Slider'lar
        sf = tk.Frame(self.root, bg=C["bg"])
        sf.pack(fill="x", padx=28)
        self._make_slider(sf, "PARLAKLIK", self.brightness_var, 0, 2, 0)
        self._make_slider(sf, "GÜNCELLEME HIZI  (ms)", self.fps_var, 10, 200, 1)
        self._make_slider_float(sf, "GEÇİŞ YUMUŞAKLIĞI", self.smooth_var, 0.05, 0.95, 2)
        self.smooth_var.trace_add("write", self._save_settings)

        tk.Frame(self.root, bg=C["border"], height=1).pack(fill="x", padx=28, pady=14)

        ctrl = tk.Frame(self.root, bg=C["bg"])
        ctrl.pack(fill="x", padx=28, pady=(0,24))
        self.btn_start = tk.Button(ctrl, text="▶   BAŞLAT",
                                   bg=C["accent"], fg="white",
                                   font=("Courier", 12, "bold"),
                                   relief="flat", pady=12, cursor="hand2",
                                   activebackground=C["accent2"],
                                   activeforeground="white",
                                   command=self._toggle)
        self.btn_start.pack(side="left", fill="both", expand=True, padx=(0,8))
        tk.Button(ctrl, text="■   IŞIĞI KAP",
                  bg=C["surface2"], fg=C["muted"],
                  font=("Courier", 12), relief="flat", pady=12,
                  cursor="hand2", activebackground=C["dim"],
                  activeforeground=C["text"],
                  command=self._turn_off).pack(side="left", fill="both", expand=True, padx=(8,0))

        self.status_var = tk.StringVar(value="Hazır")
        sb = tk.Frame(self.root, bg=C["surface"], pady=8)
        sb.pack(fill="x", side="bottom")
        tk.Label(sb, textvariable=self.status_var,
                 font=("Courier", 9), fg=C["muted"], bg=C["surface"]).pack()

    def _build_keyboard_preview(self):
        kf = tk.Frame(self.root, bg=C["surface"], padx=16, pady=14)
        kf.pack(fill="x", padx=28)
        tk.Label(kf, text="KLAVYE ÖNİZLEME", font=("Courier", 8),
                 fg=C["muted"], bg=C["surface"]).pack(anchor="w", pady=(0,8))
        row = tk.Frame(kf, bg=C["surface"])
        row.pack(fill="x")
        self.zone_canvases = {}
        for i, (lbl, zid) in enumerate([("SOL", ZONE_LEFT), ("ORTA", ZONE_MID), ("SAĞ", ZONE_RIGHT)]):
            zf = tk.Frame(row, bg=C["surface"])
            zf.grid(row=0, column=i, sticky="nsew", padx=(0 if i==0 else 6, 0))
            row.columnconfigure(i, weight=1)
            cv = tk.Canvas(zf, height=38, bg="#111118",
                           highlightthickness=1, highlightbackground=C["border"])
            cv.pack(fill="x")
            tk.Label(zf, text=lbl, font=("Courier", 7), fg=C["muted"],
                     bg=C["surface"]).pack(pady=(3,0))
            self.zone_canvases[zid] = cv

    def _build_static_picker(self, parent):
        inner = tk.Frame(parent, bg=C["surface"], padx=16, pady=12)
        inner.pack(fill="x")
        tk.Label(inner, text="SABİT RENK", font=("Courier", 8),
                 fg=C["muted"], bg=C["surface"]).pack(anchor="w", pady=(0,8))
        row = tk.Frame(inner, bg=C["surface"])
        row.pack(fill="x")
        self.static_swatch = tk.Canvas(row, width=42, height=42,
                                        highlightthickness=0, bg=C["surface"])
        self.static_swatch.pack(side="left", padx=(0,12))
        self._draw_swatch()
        info = tk.Frame(row, bg=C["surface"])
        info.pack(side="left", fill="x", expand=True)
        self.hex_lbl = tk.Label(info, text='#%02x%02x%02x' % self.static_color,
                                 font=("Courier", 14, "bold"), fg=C["text"],
                                 bg=C["surface"])
        self.hex_lbl.pack(anchor="w")
        tk.Label(info, text="tıkla veya aşağıdan seç",
                 font=("Courier", 8), fg=C["muted"], bg=C["surface"]).pack(anchor="w")
        tk.Button(row, text="🎨", font=("Segoe UI Emoji", 16),
                  bg=C["dim"], fg=C["text"], relief="flat", padx=8,
                  cursor="hand2", command=self._open_static_wheel).pack(side="right")
        qr = tk.Frame(inner, bg=C["surface"])
        qr.pack(fill="x", pady=(10,0))
        tk.Label(qr, text="HIZLI", font=("Courier", 7),
                 fg=C["muted"], bg=C["surface"]).pack(side="left", padx=(0,8))
        for qc in QUICK_COLORS:
            hx = '#%02x%02x%02x' % qc
            tk.Button(qr, bg=hx, width=2, height=1, relief="flat", cursor="hand2",
                      command=lambda c=qc: self._set_static(c)).pack(side="left", padx=2)
        self.static_swatch.bind("<Button-1>", lambda _: self._open_static_wheel())

    def _build_wave_panel(self, parent):
        inner = tk.Frame(parent, bg=C["surface"], padx=16, pady=12)
        inner.pack(fill="x")
        tk.Label(inner, text="DALGA PALETİ", font=("Courier", 8),
                 fg=C["muted"], bg=C["surface"]).pack(anchor="w", pady=(0,6))
        self.wave_grad = tk.Canvas(inner, height=20, bg=C["bg"], highlightthickness=0)
        self.wave_grad.pack(fill="x", pady=(0,8))
        self.wave_grad.bind("<Configure>", lambda _: self._draw_wave_grad())
        self.wave_swatches_row = tk.Frame(inner, bg=C["surface"])
        self.wave_swatches_row.pack(fill="x")
        self._render_wave_swatches()
        tk.Button(inner, text="🎨  Paleti Düzenle", bg=C["dim"], fg=C["cyan"],
                  font=("Courier", 9, "bold"), relief="flat", padx=12, pady=6,
                  cursor="hand2", command=self._open_wave_editor).pack(pady=(10,0))
        spd_f = tk.Frame(inner, bg=C["surface"])
        spd_f.pack(fill="x", pady=(10,0))
        tk.Label(spd_f, text="DALGA HIZI", font=("Courier", 8),
                 fg=C["muted"], bg=C["surface"]).pack(anchor="w")
        saved_spd = load_settings().get("wave_speed", 1.0)
        self.wave_speed_var = tk.DoubleVar(value=saved_spd)
        tk.Scale(spd_f, from_=0.1, to=5.0, resolution=0.1, orient="horizontal",
                 variable=self.wave_speed_var, bg=C["surface"], fg=C["cyan"],
                 troughcolor=C["dim"], highlightthickness=0, bd=0,
                 showvalue=True, command=lambda _: self._save_settings()).pack(fill="x")

    def _render_wave_swatches(self):
        for w in self.wave_swatches_row.winfo_children():
            w.destroy()
        for i, col in enumerate(self.wave_colors):
            hx = '#%02x%02x%02x' % col
            f  = tk.Frame(self.wave_swatches_row, bg=C["surface"])
            f.pack(side="left", padx=(0 if i==0 else 4, 0))
            tk.Canvas(f, width=32, height=32, bg=hx,
                      highlightthickness=1,
                      highlightbackground=C["border"]).pack()
        self._draw_wave_grad()

    def _draw_wave_grad(self):
        self.wave_grad.delete("all")
        w = self.wave_grad.winfo_width()
        h = self.wave_grad.winfo_height() or 20
        if w < 2: return
        for x in range(w):
            r, g, b = sample_palette(self.wave_colors, x/w)
            self.wave_grad.create_line(x, 0, x, h, fill='#%02x%02x%02x'%(r,g,b))

    def _draw_swatch(self):
        self.static_swatch.delete("all")
        r, g, b = self.static_color
        self.static_swatch.create_rectangle(0, 0, 42, 42,
                                             fill='#%02x%02x%02x'%(r,g,b), outline="")

    def _make_slider(self, parent, label, var, from_, to_, row):
        tk.Label(parent, text=label, font=("Courier", 8, "bold"),
                 fg=C["muted"], bg=C["bg"]).grid(row=row*2, column=0, sticky="w", pady=(8,2))
        tk.Label(parent, textvariable=var, font=("Courier", 10, "bold"),
                 fg=C["cyan"], bg=C["bg"], width=4).grid(row=row*2, column=1, sticky="e")
        tk.Scale(parent, from_=from_, to=to_, orient="horizontal",
                 variable=var, bg=C["bg"], fg=C["text"],
                 troughcolor=C["dim"], activebackground=C["accent"],
                 highlightthickness=0, bd=0, showvalue=False, length=490
                 ).grid(row=row*2+1, column=0, columnspan=2, sticky="ew", pady=(0,4))
        parent.columnconfigure(0, weight=1)

    def _make_slider_float(self, parent, label, var, from_, to_, row):
        """0.0-1.0 arası float slider — sağda 2 ondalık gösterir"""
        # Label
        tk.Label(parent, text=label, font=("Courier", 8, "bold"),
                 fg=C["muted"], bg=C["bg"]).grid(row=row*2, column=0, sticky="w", pady=(8,2))
        # Değer etiketi (formatlanmış)
        val_lbl = tk.Label(parent, text=f"{var.get():.2f}",
                           font=("Courier", 10, "bold"),
                           fg=C["cyan"], bg=C["bg"], width=5)
        val_lbl.grid(row=row*2, column=1, sticky="e")

        def _on_change(_=None):
            val_lbl.config(text=f"{var.get():.2f}")

        tk.Scale(parent, from_=from_, to=to_, resolution=0.01, orient="horizontal",
                 variable=var, bg=C["bg"], fg=C["text"],
                 troughcolor=C["dim"], activebackground=C["accent"],
                 highlightthickness=0, bd=0, showvalue=False, length=490,
                 command=_on_change
                 ).grid(row=row*2+1, column=0, columnspan=2, sticky="ew", pady=(0,4))
        parent.columnconfigure(0, weight=1)

    # ── Mod ────────────────────────────────────────────────────
    def _on_mode_change(self):
        mode = self.mode_var.get()
        self.static_panel.pack_forget()
        self.wave_panel.pack_forget()
        if mode == "static":
            self.static_panel.pack(fill="x")
        elif mode == "wave":
            self.wave_panel.pack(fill="x")

    # ── Renk seçimi ────────────────────────────────────────────
    def _set_static(self, color):
        self.static_color = color
        self._draw_swatch()
        self.hex_lbl.config(text='#%02x%02x%02x' % color)
        self._update_all_zones(color)
        self._save_settings()

    def _open_static_wheel(self):
        ColorWheelPicker(self.root, initial=self.static_color, callback=self._set_static)

    def _open_wave_editor(self):
        WavePaletteEditor(self.root, self.wave_colors,
                          callback=self._on_wave_palette_changed)

    def _on_wave_palette_changed(self, colors):
        self.wave_colors = colors
        self._render_wave_swatches()
        self._save_settings()

    # ── Klavye önizleme ────────────────────────────────────────
    def _update_zone_preview(self, zone_id, r, g, b):
        cv = self.zone_canvases.get(zone_id)
        if not cv: return
        hx = '#%02x%02x%02x' % (r, g, b)
        w  = cv.winfo_width()  or 140
        h  = cv.winfo_height() or 38
        cv.delete("all")
        cv.create_rectangle(0, 0, w, h, fill=hx, outline="")
        cx, cy = w//2, h//2
        rad = min(cx, cy)
        for i in range(rad, 0, -3):
            t  = i / rad
            cr = int(r + (255-r)*(1-t)*0.22)
            cg = int(g + (255-g)*(1-t)*0.22)
            cb = int(b + (255-b)*(1-t)*0.22)
            cv.create_oval(cx-i, cy-i, cx+i, cy+i,
                           fill='#%02x%02x%02x'%(min(255,cr),min(255,cg),min(255,cb)),
                           outline="")

    def _update_all_zones(self, color):
        for zid in [ZONE_LEFT, ZONE_MID, ZONE_RIGHT]:
            self.root.after(0, lambda z=zid, c=color: self._update_zone_preview(z, *c))

    # ── Kontrol ────────────────────────────────────────────────
    def _toggle(self):
        if not self.wmi_ok:
            self.status_var.set("⚠ WMI bağlantısı yok")
            return
        if self.running:
            self.running = False
            self.btn_start.config(text="▶   BAŞLAT", bg=C["accent"])
            self.status_var.set("Durduruldu")
        else:
            self.running   = True
            self._breath_t = 0.0
            self._wave_t   = 0.0
            self._cycle_h  = 0.0
            # Smooth durumlarını sıfırla
            self._smooth   = {ZONE_LEFT: (0,0,0), ZONE_MID: (0,0,0), ZONE_RIGHT: (0,0,0)}
            self.btn_start.config(text="■   DURDUR", bg=C["red"])
            self.status_var.set("Çalışıyor…")
            threading.Thread(target=self._loop, daemon=True).start()

    def _turn_off(self):
        self.running = False
        self.btn_start.config(text="▶   BAŞLAT", bg=C["accent"])
        self.status_var.set("Kapatıldı")
        try:
            import pythoncom; pythoncom.CoInitialize()
            c    = wmi.WMI(namespace="root\\wmi")
            inst = c.query(f"SELECT * FROM {WMI_CLASS}")[0]
            inst.BufferBytes = list(build_packet(ZONE_ALL, 0, 0, 0, 0))
            inst.put()
            pythoncom.CoUninitialize()
        except: pass
        self._update_all_zones((0,0,0))

    # ── Sistem Tepsisi ─────────────────────────────────────────
    def _start_tray(self):
        if not TRAY_OK: return
        if self._tray: return

        icon_img = make_tray_icon()
        menu = pystray.Menu(
            TrayItem("Göster",    lambda: self.root.after(0, self._show_from_tray)),
            TrayItem("Başlat",    lambda: self.root.after(0, self._toggle)
                     if not self.running else None),
            TrayItem("Durdur",    lambda: self.root.after(0, self._turn_off)),
            pystray.Menu.SEPARATOR,
            TrayItem("Çıkış",     lambda: self.root.after(0, self._on_close)),
        )
        self._tray = pystray.Icon("ExcaGlow", icon_img, "ExcaGlow V6", menu)
        self._tray_thread = threading.Thread(target=self._tray.run, daemon=True)
        self._tray_thread.start()

    def _hide_to_tray(self):
        if not TRAY_OK:
            self.status_var.set("pystray kurulu değil: pip install pystray")
            return
        self._start_tray()
        self.root.withdraw()

    def _show_from_tray(self):
        self.root.deiconify()
        self.root.lift()

    # ── Ekran örnekleme ────────────────────────────────────────
    def _grab(self):
        """Ekranın ortasındaki %50'lik alanı örnekle"""
        with mss() as sct:
            mon = sct.monitors[1]
            reg = {"left":   mon["left"] + mon["width"]//4,
                   "top":    mon["top"]  + mon["height"]//4,
                   "width":  mon["width"]//2,
                   "height": mon["height"]//2}
            img = sct.grab(reg)
            pil = Image.frombytes("RGB", img.size, img.bgra, "raw", "BGRX")
            pil = ImageEnhance.Color(pil).enhance(1.5)
            return pil.resize((1,1), resample=Image.BILINEAR).getpixel((0,0))

    def _grab_zones(self):
        """3 bölge — klavye yönüne göre ters (ekran sağ = klavye sol)"""
        with mss() as sct:
            mon = sct.monitors[1]
            w   = mon["width"] // 3
            res = []
            for i in [2, 1, 0]:
                reg = {"left": mon["left"]+i*w, "top": mon["top"],
                       "width": w, "height": mon["height"]}
                img = sct.grab(reg)
                pil = Image.frombytes("RGB", img.size, img.bgra, "raw", "BGRX")
                pil = ImageEnhance.Color(pil).enhance(1.5)
                res.append(pil.resize((1,1), resample=Image.BILINEAR).getpixel((0,0)))
            return res  # [klavye_sol, klavye_mid, klavye_sag]

    # ── WMI yazma ──────────────────────────────────────────────
    def _send(self, inst, zone, r, g, b, bri):
        """inst: döngü başında bir kez alınmış WMI instance"""
        inst.BufferBytes = list(build_packet(zone, r, g, b, bri))
        inst.put()

    def _send_all(self, inst, colors_by_zone, bri):
        """3 bölgeyi araya sleep koymadan gönder"""
        for zone, (r, g, b) in colors_by_zone.items():
            self._send(inst, zone, r, g, b, bri)

    # ── Ana döngü ──────────────────────────────────────────────
    def _loop(self):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            wc   = wmi.WMI(namespace="root\\wmi")
            inst = wc.query(f"SELECT * FROM {WMI_CLASS}")[0]  # bir kez al, yeniden kullan

            while self.running:
                m   = self.mode_var.get()
                bri = self.brightness_var.get()
                dt  = self.fps_var.get() / 1000.0

                # ── Ambient ──────────────────────────────────
                if m == "ambient":
                    target = self._grab()
                    # Tüm bölgeler için aynı smooth
                    s = self._smooth[ZONE_LEFT]
                    sf = self.smooth_var.get()
                    new_s = smooth_color(s, target, sf)
                    for z in [ZONE_LEFT, ZONE_MID, ZONE_RIGHT]:
                        self._smooth[z] = new_s
                    r, g, b = new_s
                    self._send(inst, ZONE_ALL, r, g, b, bri)
                    col = (r, g, b)
                    self.root.after(0, lambda c=col: self._update_all_zones(c))

                # ── 3 Bölge ──────────────────────────────────
                elif m == "zones":
                    targets = self._grab_zones()
                    zone_ids = [ZONE_LEFT, ZONE_MID, ZONE_RIGHT]
                    to_send  = {}
                    sf = self.smooth_var.get()
                    for zid, target in zip(zone_ids, targets):
                        new_s = smooth_color(self._smooth[zid], target, sf)
                        self._smooth[zid] = new_s
                        to_send[zid] = new_s

                    # Araya sleep koymadan 3 paketi ardı ardına gönder
                    self._send_all(inst, to_send, bri)

                    for zid, col in to_send.items():
                        self.root.after(0, lambda z=zid, c=col:
                                        self._update_zone_preview(z, *c))

                # ── Sabit ────────────────────────────────────
                elif m == "static":
                    target = self.static_color
                    s  = self._smooth[ZONE_LEFT]
                    sf = self.smooth_var.get()
                    new_s = smooth_color(s, target, sf)
                    for z in [ZONE_LEFT, ZONE_MID, ZONE_RIGHT]:
                        self._smooth[z] = new_s
                    r, g, b = new_s
                    self._send(inst, ZONE_ALL, r, g, b, bri)

                # ── Nefes ────────────────────────────────────
                elif m == "breathe":
                    self._breath_t += dt * 0.7
                    factor = ((math.sin(self._breath_t * math.pi * 2) + 1) / 2) ** 1.6
                    sc     = self.static_color
                    target = (int(sc[0]*factor), int(sc[1]*factor), int(sc[2]*factor))
                    new_s  = smooth_color(self._smooth[ZONE_LEFT], target, min(0.9, self.smooth_var.get()*1.5))
                    for z in [ZONE_LEFT, ZONE_MID, ZONE_RIGHT]:
                        self._smooth[z] = new_s
                    r, g, b = new_s
                    self._send(inst, ZONE_ALL, r, g, b, bri)
                    col = (r, g, b)
                    self.root.after(0, lambda c=col: self._update_all_zones(c))

                # ── Dalga ────────────────────────────────────
                elif m == "wave":
                    spd = self.wave_speed_var.get() if hasattr(self, "wave_speed_var") else 1.0
                    self._wave_t += dt * spd * 0.4

                    # Her bölge paletin farklı noktasından — smooth ile akıcı
                    offsets  = [0.0, 1/3, 2/3]
                    to_send  = {}
                    for zid, off in zip([ZONE_LEFT, ZONE_MID, ZONE_RIGHT], offsets):
                        pos    = (self._wave_t + off) % 1.0
                        target = sample_palette(self.wave_colors, pos)
                        new_s  = smooth_color(self._smooth[zid], target, self.smooth_var.get())
                        self._smooth[zid] = new_s
                        to_send[zid] = new_s

                    # Araya sleep koymadan 3 paketi gönder
                    self._send_all(inst, to_send, bri)

                    for zid, col in to_send.items():
                        self.root.after(0, lambda z=zid, c=col:
                                        self._update_zone_preview(z, *c))

                # ── Döngü (Rainbow) ──────────────────────────
                elif m == "cycle":
                    self._cycle_h = (self._cycle_h + dt * 45) % 360
                    target = hsv_to_rgb(self._cycle_h, 1.0, 1.0)
                    new_s  = smooth_color(self._smooth[ZONE_LEFT], target, self.smooth_var.get())
                    for z in [ZONE_LEFT, ZONE_MID, ZONE_RIGHT]:
                        self._smooth[z] = new_s
                    r, g, b = new_s
                    self._send(inst, ZONE_ALL, r, g, b, bri)
                    col = (r, g, b)
                    self.root.after(0, lambda c=col: self._update_all_zones(c))

                time.sleep(dt)

        except Exception as e:
            print(f"Döngü Hatası: {e}")
            self.running = False
            self.root.after(0, lambda: [
                self.btn_start.config(text="▶   BAŞLAT", bg=C["accent"]),
                self.status_var.set("Hata — yeniden dene")
            ])
        finally:
            import pythoncom
            pythoncom.CoUninitialize()


# ── Giriş Noktası ────────────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    ExcaGlowApp(root)
    root.mainloop()