"""
GUI Tkinter – v2
- Import CSV/Excel
- Recherche Nom/Prénom
- Grille visible : ✓, Nom, Prénom, Dernière impression, # Impressions
- DDN et Expire = données cachées (BDD & logique métier)
- Impression -> génère ZPL + enregistre en SQLite (status='printed')
- Tri sur Nom/Prénom via clic sur l'en-tête (asc/desc), conservation des coches
- Fermeture propre via la croix (WM_DELETE_WINDOW)

Linux : sudo apt install -y python3-tk
Lancement : python gui.py
"""
from __future__ import annotations

import sys
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

# --- Impression Brother QL ---
from brother_ql.backends import backend_factory
from brother_ql.conversion import convert
from brother_ql.raster import BrotherQLRaster
from PIL import Image, ImageDraw, ImageFont

# Import des modules applicatifs
ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from pathlib import Path
import sys
if getattr(sys, "frozen", False):
    ROOT = Path(sys.executable).resolve().parent
else:
    ROOT = Path(__file__).resolve().parent
SRC_DIR = ROOT / "src"
if SRC_DIR.exists() and str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from app.io_utils import lire_tableau
from app.zpl import genere_zpl, ecrire_sorties
from app.db import init_db, connect, record_print


import configparser

CONFIG_PATH = ROOT / "config.ini"

def load_config() -> dict[str, str]:
    cfg = configparser.ConfigParser()
    if CONFIG_PATH.exists():
        cfg.read(CONFIG_PATH, encoding="utf-8")
    return {
        "backend": cfg.get("impression", "backend", fallback="win32print"),
        "device": cfg.get("impression", "device", fallback="Brother QL-570"),
        "label": cfg.get("impression", "label", fallback="62"),
        "default_expire": cfg.get("app", "default_expire", fallback=""),
        "auto_import_file": cfg.get("app", "auto_import_file", fallback="deja_imprimes.csv"),
        "rotate": cfg.get("impression", "rotate", fallback="0"),
    }


def build_ddn_lookup_from_rows(rows: list[dict]) -> dict[tuple[str, str], str | None]:
    """
    Construit un mapping (nom_lower, prenom_lower) -> ddn (str) ou None s'il y a conflit / inconnu.
    Prend les DDN depuis le tableau importé (colonnes 'Nom', 'Prénom', 'Date_de_naissance').
    """
    tmp: dict[tuple[str, str], set[str]] = {}
    for r in rows or []:
        nom = (r.get("Nom") or "").strip().lower()
        prenom = (r.get("Prénom") or "").strip().lower()
        ddn = (r.get("Date_de_naissance") or "").strip()
        if not nom and not prenom:
            continue
        key = (nom, prenom)
        tmp.setdefault(key, set())
        if ddn:
            tmp[key].add(ddn)

    out: dict[tuple[str, str], str | None] = {}
    for key, s in tmp.items():
        if len(s) == 1:
            out[key] = next(iter(s))           # DDN unique depuis les lignes
        elif len(s) == 0:
            out[key] = None                    # inconnu
        else:
            out[key] = None                    # conflit : plusieurs DDN pour le même Nom/Prénom
    return out

import csv
from app.db import connect, record_print

def import_already_printed_csv(csv_path: Path, expire: str, rows_ddn_lookup: dict[tuple[str, str], str | None] | None = None) -> tuple[int, int]:
    """
    Importe un CSV 'nom;prenom' et crée 1 ligne 'printed' par personne pour l'expiration donnée.
    Résout la DDN ainsi :
      1) DB: si une et une seule DDN non vide existe pour (nom, prenom), on la prend
      2) Fallback: si rows_ddn_lookup en propose une (unique), on la prend
      3) Sinon: DDN vide
    Retourne (importés, ignorés).
    """
    if not csv_path.exists():
        return (0, 0)

    imported, skipped = 0, 0
    with connect(DB_PATH) as cn:
        for nom, prenom in csv.reader(open(csv_path, "r", encoding="utf-8", newline=""), delimiter=";"):
            nom = (nom or "").strip()
            prenom = (prenom or "").strip()
            if not nom and not prenom:
                skipped += 1
                continue

            key = (nom.lower(), prenom.lower())

            # 1) Cherche DDN unique en base
            ddns = cn.execute(
                """
                SELECT DISTINCT ddn
                FROM prints
                WHERE LOWER(nom)=? AND LOWER(prenom)=?
                """,
                key,
            ).fetchall()
            ddn_candidates = [ (row[0] or "").strip() for row in ddns if (row[0] or "").strip() ]
            if len(ddn_candidates) == 1:
                ddn = ddn_candidates[0]
            elif len(ddn_candidates) > 1:
                ddn = ""  # ambigu → joker
            else:
                # 2) Fallback: lookup depuis les lignes déjà importées dans la GUI
                ddn = (rows_ddn_lookup.get(key) or "") if rows_ddn_lookup else ""
                ddn = ddn or ""  # None -> ""

            record_print(cn, nom, prenom, ddn=ddn, expire=expire, zpl=None, status="printed")
            imported += 1

    return (imported, skipped)

MODEL = "QL-570"
DEFAULT_CANVAS = {    # largeur px à 300 dpi
    62: (696, 300),
    38: (413, 300),
    29: (306, 300),
    12: (118, 300),
}

def saison_from_expire(expire: str) -> str:
    """Calcule 'AAAA / AAAA+1' à partir d'une date JJ/MM/AAAA (expire)."""
    from datetime import datetime
    try:
        dt = datetime.strptime(expire, "%d/%m/%Y")
        return f"{dt.year-1} / {dt.year}"
    except Exception:
        return ""

def _find_font() -> ImageFont.ImageFont | None:
    """Essaie de charger une TTF lisible, sinon fallback PIL par défaut."""
    for p in [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",  # Linux
        "C:\\Windows\\Fonts\\arial.ttf",                    # Windows
    ]:
        try:
            if Path(p).exists():
                return ImageFont.truetype(p, 36)
        except Exception:
            pass
    return None

def make_label_image_simple(nom: str, prenom: str, expire: str, label_mm: int = 62) -> Image.Image:
    """Étiquette simple : Nom, Prénom, Saison."""
    W, H = DEFAULT_CANVAS.get(label_mm, DEFAULT_CANVAS[62])
    img = Image.new("1", (W, H), 1)
    d = ImageDraw.Draw(img)
    f_big = _find_font()
    f_small = None
    try:
        if f_big and hasattr(f_big, "path"):
            f_small = ImageFont.truetype(f_big.path, 28)   # type: ignore[attr-defined]
    except Exception:
        pass

    line1 = (nom or "").upper()
    line2 = (prenom or "").capitalize()
    saison = saison_from_expire(expire) or expire
    line3 = f"Saison : {saison}"

    y = 10
    d.text((10, y), line1, fill=0, font=f_big);   y += 70
    d.text((10, y), line2, fill=0, font=f_big);   y += 70
    d.text((10, y), line3, fill=0, font=f_small)
    return img

def _open_bql_handle(backend_name: str, device: str | None):
    """
    Ouvre le handle d'impression brother_ql en fonction du backend.
    - pyusb : essaie enumerate() si device non fourni
    - linux_kernel, network, file : exigent 'device'
    - dummy : toujours ok
    """
    be = backend_factory(backend_name)

    backend_name = backend_name.lower().strip()
    # Backends qui DEMANDENT un device explicite
    needs_device = {"linux_kernel", "network", "file"}
    if backend_name in needs_device:
        if not device:
            raise RuntimeError(
                f"Le backend '{backend_name}' requiert 'device' dans config.ini :\n"
                f" - linux_kernel : /dev/usb/lp0\n"
                f" - network      : 192.168.1.50:9100\n"
                f" - file         : /chemin/sortie.bin"
            )
        return be.open(device)

    # Backends qui savent énumérer (pyusb, dummy)
    if hasattr(be, "enumerate"):
        targets = be.enumerate()
        if device:
            # On respecte la cible fournie (chemin/identifiant dépend du backend)
            return be.open(device)
        if not targets:
            raise RuntimeError(
                "Aucun périphérique Brother détecté pour backend='pyusb'.\n"
                "Vérifie le câble/driver/permissions (libusb)."
            )
        # Sur pyusb, open() accepte l'élément retourné par enumerate()
        return be.open(targets[0])

    # Cas improbable : backend sans enumerate() et sans device
    if not device:
        raise RuntimeError(
            f"Le backend '{backend_name}' ne fournit pas enumerate() et aucun 'device' n'a été donné."
        )
    return be.open(device)

def _render_label_bytes(nom: str, prenom: str, expire: str, label_mm: int, rotate_val: str | int = 0):
    img = make_label_image_simple(nom, prenom, expire, label_mm)

    from brother_ql.raster import BrotherQLRaster
    qlr = BrotherQLRaster(MODEL)        # ex. "QL-570"
    qlr.exception_on_warning = True

    try:
        rot = int(rotate_val)
    except Exception:
        rot = 0  # défaut : pas de rotation

    res = convert(
        qlr=qlr,
        images=[img],
        label=str(label_mm),
        rotate=rot,
        threshold=70,
        dither=False,
        compress=False,
        red=False,
        dpi_600=False,
        hq=True,
        cut=False,
    )

    # convert() renvoie soit un objet avec .output, soit directement des bytes
    payload = res.output if hasattr(res, "output") else res

    return payload, img

def _print_via_brotherql(backend_name: str, device: str | None, payload: bytes):
    h = _open_bql_handle(backend_name, device)
    h.write(payload)

def _print_via_win32_driver(device_name: str, pil_image):
    # Impression via driver Windows (win32print) : pas de brother_ql ici
    import win32print, win32ui
    from PIL import ImageWin
    if not device_name:
        device_name = win32print.GetDefaultPrinter()
    hDC = win32ui.CreateDC()
    hDC.CreatePrinterDC(device_name)
    hDC.StartDoc("Etiquette CSN"); hDC.StartPage()
    dib = ImageWin.Dib(pil_image.convert("RGB"))
    # Ajuste la zone si besoin :
    w, h = pil_image.size
    dib.draw(hDC.GetHandleOutput(), (0, 0, w, h))
    hDC.EndPage(); hDC.EndDoc(); hDC.DeleteDC()

def print_ql570_direct(nom: str, prenom: str, ddn: str, expire: str,
                       label: str = "62", backend_name: str = "pyusb",
                       device: str | None = None, rotate: str = "0"):
    label_mm = int(label or "62")
    payload, img = _render_label_bytes(nom, prenom, expire, label_mm, rotate_val=rotate)

    if backend_name.lower() in {"pyusb", "linux_kernel", "network", "file", "dummy"}:
        _print_via_brotherql(backend_name, device, payload)
    elif backend_name.lower() == "win32print":
        _print_via_win32_driver(device or "", img)
    else:
        raise RuntimeError(f"Backend inconnu : {backend_name}")

DEFAULT_EXPIRATION = "31/12/2026"
SORTIES_DIR = ROOT / "data" / "sorties"
DB_PATH = ROOT / "data" / "app.db"

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Étiquettes – GUI")
        self.geometry("1180x650")
        self.minsize(1000, 560)

        # État
        self.rows: list[dict] = []      # lignes importées (avec DDN/Expire cachés)
        self.view_rows: list[dict] = [] # lignes affichées après filtre
        self.checked: set[int] = set()  # indices cochés dans view_rows
        self.sort_col: str | None = None
        self.sort_asc: bool = True

        self._build_ui()
        # Gestion fermeture propre (croix)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.mode_var = tk.StringVar(value="tout")  # "a_imprimer" | "deja" | "tout"
        self.per_expire_count = {}  # (nom, prenom, ddn, expire) -> nb 'printed' pour CETTE expiration

        self.cfg = load_config()
        if self.cfg.get("default_expire"):
            self.exp_var.set(self.cfg["default_expire"])

    # ---------------- UI ----------------
    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=8)

        ttk.Label(top, text="Afficher :").pack(side=tk.LEFT, padx=(16, 4))
        cb = ttk.Combobox(top, width=16, state="readonly",
                  values=["À imprimer", "Déjà imprimées", "Tout"])
        cb.current(2)  # "Tout"
        cb.pack(side=tk.LEFT)
        def _on_mode_change(event=None):
            v = cb.get()
            self.mode_var.set("a_imprimer" if v == "À imprimer"
                      else "deja" if v == "Déjà imprimées"
                      else "tout")
            self.apply_filter()
        cb.bind("<<ComboboxSelected>>", _on_mode_change)
        
        ttk.Button(top, text="Importer CSV/Excel", command=self.on_import).pack(side=tk.LEFT, padx=(6, 0))
        self.cfg = load_config()
        if self.cfg.get("backend") != "win32print":
            from app.zpl import genere_zpl, ecrire_sorties
            ttk.Button(top, text="Imprimer ZPL", command=self.on_print).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(top, text="Imprimer QL-570", command=self.on_print_ql570).pack(side=tk.LEFT, padx=(6, 0))

        ttk.Label(top, text="Expiration à garder :").pack(side=tk.LEFT, padx=(16, 4))
        self.exp_var = tk.StringVar(value=DEFAULT_EXPIRATION)
        ttk.Entry(top, width=12, textvariable=self.exp_var).pack(side=tk.LEFT)

        ttk.Label(top, text="Nom :").pack(side=tk.LEFT, padx=(16, 4))
        self.nom_var = tk.StringVar()
        ttk.Entry(top, width=16, textvariable=self.nom_var).pack(side=tk.LEFT)

        ttk.Label(top, text="Prénom :").pack(side=tk.LEFT, padx=(8, 4))
        self.prenom_var = tk.StringVar()
        ttk.Entry(top, width=16, textvariable=self.prenom_var).pack(side=tk.LEFT)

        ttk.Button(top, text="Rechercher", command=self.apply_filter).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(top, text="Réinitialiser", command=self.reset_filter).pack(side=tk.LEFT, padx=(6, 0))
    
        ttk.Button(top, text="Init DB", command=self.on_init_db).pack(side=tk.LEFT)
        ttk.Button(top, text="Réinitialiser la base", command=self.on_reset_db).pack(side=tk.LEFT, padx=(6, 0))

        mid = ttk.Frame(self)
        mid.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        cols = ("sel", "Nom", "Prénom", "Derniere", "Compteur")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="extended")
        self.tree.heading("sel", text="✓", command=self.toggle_all)
        self.tree.heading("Nom", text="Nom", command=lambda: self.sort_by("Nom"))
        self.tree.heading("Prénom", text="Prénom", command=lambda: self.sort_by("Prénom"))
        self.tree.heading("Derniere", text="Dernière impression")
        self.tree.heading("Compteur", text="# Impressions")

        self.tree.column("sel", width=48, anchor=tk.CENTER, stretch=False)
        self.tree.column("Nom", width=240)
        self.tree.column("Prénom", width=240)
        self.tree.column("Derniere", width=200, anchor=tk.CENTER)
        self.tree.column("Compteur", width=130, anchor=tk.E)

        self.tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=sb.set)
        sb.pack(fill=tk.Y, side=tk.RIGHT)

        self.tree.bind("<Button-1>", self.on_tree_click)

        self.status = tk.StringVar(value="Prêt")
        ttk.Label(self, textvariable=self.status, anchor=tk.W).pack(fill=tk.X, padx=8, pady=(0, 6))

        self._update_headers()

    def on_print_ql570(self):
        # Vérifs basiques
        if not self.view_rows:
            messagebox.showinfo("Info", "Aucune donnée à imprimer.")
            return
        if not self.checked:
            messagebox.showinfo("Info", "Sélectionnez au moins une ligne (colonne ✓).")
            return

        expiration = (self.exp_var.get() or "").strip()
        if not expiration:
            messagebox.showerror("Erreur", "Entrez une date d'expiration (JJ/MM/AAAA)")
            return

        # Sélection (on respecte Expire == expiration)
        selected = []
        for idx in sorted(self.checked):
            r = self.view_rows[idx]
            if (r.get("Expire_le") or "").strip() == expiration:
               selected.append(r)
        if not selected:
            messagebox.showinfo("Info", f"Aucune ligne sélectionnée avec Expire le = {expiration}.")
            return

        # Impression directe (Windows: backend 'win32' + nom d’imprimante; Linux: 'pyusb')
        # Adapte backend/device si besoin (ex: backend_name='win32print', device='Brother QL-570')
        backend_name = self.cfg.get("backend", "win32print")
        device_name = self.cfg.get("device") or None
        label = self.cfg.get("label", "62")

        # Sur Windows, tu mettras :
        # backend_name = "win32print; device_name = "Brother QL-570"

        old_cursor = self["cursor"]
        self.config(cursor="watch")
        self.toast("Impression QL-570 en cours…")
        self.update_idletasks()

        try:
        # Envoie 1 étiquette par personne
            for r in selected:
                nom     = (r.get("Nom") or "").strip()
                prenom  = (r.get("Prénom") or "").strip()
                ddn     = (r.get("Date_de_naissance") or "").strip()
            expire  = (r.get("Expire_le") or "").strip()

            print_ql570_direct(nom, prenom, ddn, expire,
                               label=label, backend_name=backend_name, device=device_name)

            # Journalise l'impression (status=printed)
            cn = connect(DB_PATH)
            with cn:
                from app.db import record_print
                record_print(cn, nom, prenom, ddn, expire, zpl=None, status="printed")

            # Rafraîchit stats + vue
            self.refresh_from_db_stats()
            self.apply_filter()
            self.toast("Impressions QL-570 ok")
            messagebox.showinfo("Succès", f"{len(selected)} étiquette(s) envoyée(s) à la QL-570.")

        except Exception as e:
            messagebox.showerror("Erreur", f"Impression QL-570 : {e}")
        finally:
            self.config(cursor=old_cursor)
            self.update_idletasks()

    def on_reset_db(self):
        ok = messagebox.askyesno(
            "Confirmation",
            "Réinitialiser la base ?\n\n"
            "Cette action SUPPRIME le fichier data/app.db puis recrée le schéma.",
            icon=messagebox.WARNING,
        )
        if not ok:
            return
        try:
            db_file = DB_PATH
            if db_file.exists():
                db_file.unlink()  # supprime la base
            from app.db import init_db
            init_db(DB_PATH)     # recrée tables / index / vues
            self.refresh_from_db_stats()
            self.apply_filter()
            self.toast("Base réinitialisée.")
            messagebox.showinfo("OK", "Base réinitialisée.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Reset DB : {e}")


    # -------------- Helpers --------------
    def _row_key(self, r: dict) -> tuple[str, str, str]:
        """Clé unique par personne pour préserver la sélection (Nom+Prénom+DDN)."""
        return (
            (r.get("Nom", "") or "").strip().lower(),
            (r.get("Prénom", "") or "").strip().lower(),
            (r.get("Date_de_naissance", "") or "").strip(),
        )

    def _fmt_dt(self, s: str | None) -> str:
        if not s:
            return ""
        try:
            dt = datetime.fromisoformat(s)
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return s

    def _update_headers(self):
        arrow_nom = " ▲" if (self.sort_col == "Nom" and self.sort_asc) else (" ▼" if self.sort_col == "Nom" else "")
        arrow_prenom = " ▲" if (self.sort_col == "Prénom" and self.sort_asc) else (" ▼" if self.sort_col == "Prénom" else "")
        self.tree.heading("Nom", text="Nom" + arrow_nom, command=lambda: self.sort_by("Nom"))
        self.tree.heading("Prénom", text="Prénom" + arrow_prenom, command=lambda: self.sort_by("Prénom"))

    # -------------- Actions --------------
    def on_refresh(self):
        # Rafraîchit depuis la BDD (stats) puis réapplique le filtre (incluant la validité)
        self.refresh_from_db_stats()
        self.apply_filter()
        self.toast("Données rafraîchies")

    def on_close(self):
        # Fermeture propre de la fenêtre
        try:
            self.destroy()
        except Exception:
            pass

    def on_init_db(self):
        ok = messagebox.askyesno(
            "Confirmation",
            "Initialiser / mettre à jour la base SQLite ?\n\n"
            "Cette opération crée / met à jour le schéma (tables, index, vues)\n"
            "sans supprimer les données existantes.",
            icon=messagebox.WARNING,
        )
        if not ok:
            self.toast("Init DB annulé")
            return
        try:
            init_db(DB_PATH)

            # Auto-import si on trouve deja_imprimes.csv à la racine
            try:
                
                csv_init = ROOT / self.cfg.get("auto_import_file", "deja_imprimes.csv")
                exp = (self.exp_var.get() or "").strip()
                imp = 0
                skip = 0
                if csv_init.exists() and exp:
                    # construit un lookup DDN à partir des données déjà chargées dans la GUI
                    ddn_lookup = build_ddn_lookup_from_rows(self.rows)
                    imp, skip = import_already_printed_csv(csv_init, exp, rows_ddn_lookup=ddn_lookup)
                if imp or skip:
                    self.toast(f"Import init: {imp} ajout(s), {skip} ignoré(s)")
                    # rafraîchir la vue après import
                    self.refresh_from_db_stats()
                    self.apply_filter()
            except Exception as e:
                messagebox.showerror("Erreur", f"Import init (deja_imprimes.csv) : {e}")

            self.toast("Base SQLite initialisée / mise à jour")
            if self.rows:
                self.refresh_from_db_stats()
                self.apply_filter()
        except Exception as e:
            messagebox.showerror("Erreur", f"Init DB: {e}")

    def on_import(self):
        path = filedialog.askopenfilename(
            title="Choisir un fichier CSV/Excel",
            filetypes=[("Fichiers CSV", "*.csv"), ("Fichiers Excel", "*.xlsx *.xls"), ("Tous", "*.*")],
        )
        if not path:
            return
        try:
            df = lire_tableau(path)
            self.rows = df.to_dict(orient="records")
            self.refresh_from_db_stats()
            self.apply_filter()
            self.toast(f"Fichier chargé: {Path(path).name} ({len(self.rows)} lignes)")
        except Exception as e:
            messagebox.showerror("Erreur", f"Import: {e}")

    def refresh_from_db_stats(self):
        """Complète chaque row avec Derniere/Compteur (toutes expirations) + map per-expire."""
        try:
            cn = connect(DB_PATH)
        except Exception:
            return

        # ---- Stats par personne (toutes expirations confondues) ----
        stats_person = {}
        try:
            cur = cn.execute("SELECT nom, prenom, ddn, last_print, cnt FROM v_person_stats")
        except Exception:
            cur = cn.execute("""
                SELECT
                    nom, prenom, ddn,
                    MAX(CASE WHEN status='printed' THEN printed_at END) AS last_print,
                    SUM(CASE WHEN status='printed' THEN 1 ELSE 0 END)   AS cnt
                FROM prints
                GROUP BY nom, prenom, ddn
            """)
        for row in cur.fetchall():
            key = (
                (row["nom"] or "").strip().lower(),
                (row["prenom"] or "").strip().lower(),
                (row["ddn"] or "").strip(),
            )
            stats_person[key] = (row["last_print"], int(row["cnt"] or 0))

        # ---- Compteur par personne+expiration (pour filtrer À imprimer / Déjà imprimées) ----
        per_expire = {}
        cur2 = cn.execute("""
            SELECT nom, prenom, ddn, expire,
                   SUM(CASE WHEN status='printed' THEN 1 ELSE 0 END) AS cnt_exp
            FROM prints
            GROUP BY nom, prenom, ddn, expire
        """)
        for row in cur2.fetchall():
            keyx = (
                (row["nom"] or "").strip().lower(),
                (row["prenom"] or "").strip().lower(),
                (row["ddn"] or "").strip(),
                (row["expire"] or "").strip(),
            )
            per_expire[keyx] = int(row["cnt_exp"] or 0)
        self.per_expire_count = per_expire  # <-- stocké pour apply_filter()

        # ---- Injection sur les lignes importées ----
        for r in self.rows:
            pkey = (
                (r.get("Nom", "") or "").strip().lower(),
                (r.get("Prénom", "") or "").strip().lower(),
                (r.get("Date_de_naissance", "") or "").strip(),
            )
            last, cnt_total = stats_person.get(pkey, (None, 0))
            r["Derniere"] = last or ""
            r["Compteur"] = cnt_total


    def apply_filter(self):
        nom = self.nom_var.get().strip().lower()
        prenom = self.prenom_var.get().strip().lower()
        expiration = (self.exp_var.get() or "").strip()
        mode = self.mode_var.get()  # "a_imprimer" | "deja" | "tout"

        def ok(r: dict) -> bool:
            # Filtre nom/prénom
            if nom and nom not in (r.get("Nom", "") or "").strip().lower():
                return False
            if prenom and prenom not in (r.get("Prénom", "") or "").strip().lower():
                return False
            # Ne garder que la validité choisie
            if expiration and ((r.get("Expire_le", "") or "").strip() != expiration):
                return False
            # Filtre de statut d'impression pour CETTE expiration
            if mode != "tout":
                keyx_exact = (
                        (r.get("Nom", "") or "").strip().lower(),
                        (r.get("Prénom", "") or "").strip().lower(),
                        (r.get("Date_de_naissance", "") or "").strip(),
                expiration,
                )
                keyx_wild = (
                    (r.get("Nom", "") or "").strip().lower(),
                    (r.get("Prénom", "") or "").strip().lower(),
                    "",  # joker pour imports sans DDN
                    expiration,
                )               
                cnt_exp = self.per_expire_count.get(keyx_exact, self.per_expire_count.get(keyx_wild, 0))

                if mode == "a_imprimer" and cnt_exp > 0:
                    return False
                if mode == "deja" and cnt_exp == 0:
                    return False
            return True

        self.view_rows = [r for r in self.rows if ok(r)]
        self.checked.clear()
        self.render_tree()
        self._update_headers()

    def reset_filter(self):
        # Réinitialise uniquement Nom/Prénom (on conserve la validité choisie)
        self.nom_var.set("")
        self.prenom_var.set("")
        self.apply_filter()
        self.toast("Filtres nom/prénom réinitialisés")

    def render_tree(self):
        self.tree.delete(*self.tree.get_children())
        for idx, r in enumerate(self.view_rows):
            sel_txt = "[x]" if idx in self.checked else "[ ]"
            values = (sel_txt, r.get("Nom", ""), r.get("Prénom", ""), self._fmt_dt(r.get("Derniere")), r.get("Compteur", 0))
            self.tree.insert("", tk.END, iid=str(idx), values=values)
        self.status.set(f"Affichées: {len(self.view_rows)} (sélectionnées: {len(self.checked)})")

    def on_tree_click(self, event):
        # Clic sur la colonne 1 pour cocher/décocher
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        colid = self.tree.identify_column(event.x)
        if colid != "#1":
            return
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        idx = int(row_id)
        if idx in self.checked:
            self.checked.remove(idx)
        else:
            self.checked.add(idx)
        self.render_tree()

    def toggle_all(self):
        if len(self.checked) < len(self.view_rows):
            self.checked = set(range(len(self.view_rows)))
        else:
            self.checked.clear()
        self.render_tree()

    def sort_by(self, col: str):
        # Sauvegarde des éléments cochés via la clé unique
        checked_keys = {self._row_key(self.view_rows[i]) for i in self.checked}

        # Détermine le sens
        if self.sort_col == col:
            self.sort_asc = not self.sort_asc
        else:
            self.sort_col = col
            self.sort_asc = True

        def keyfunc(r: dict):
            v = r.get(col, "")
            if col == "Derniere" and v:
                try:
                    return datetime.fromisoformat(v)
                except Exception:
                    return datetime.min
            if col == "Compteur":
                try:
                    return int(v)
                except Exception:
                    return 0
            return (v or "").lower()

        self.view_rows.sort(key=keyfunc, reverse=not self.sort_asc)

        # Reconstitue la sélection
        self.checked.clear()
        for idx, r in enumerate(self.view_rows):
            if self._row_key(r) in checked_keys:
                self.checked.add(idx)

        self.render_tree()
        self._update_headers()

    def on_print(self):
        if not self.view_rows:
            messagebox.showinfo("Info", "Aucune donnée à imprimer.")
            return
        if not self.checked:
            messagebox.showinfo("Info", "Sélectionnez au moins une ligne (colonne ✓).")
            return

        expiration = self.exp_var.get().strip()
        if not expiration:
            messagebox.showerror("Erreur", "Entrez une date d'expiration (JJ/MM/AAAA)")
            return

        # Applique la règle métier Expire == expiration (même si Expire est caché)
        selected_records = []
        for idx in sorted(self.checked):
            r = self.view_rows[idx]
            if r.get("Expire_le", "").strip() == expiration:
                selected_records.append(r)
        if not selected_records:
            messagebox.showinfo("Info", f"Aucune ligne sélectionnée avec Expire le = {expiration}.")
            return

        # Génère ZPL
        try:
            self.toast("Génération des étiquettes…")
            fichiers = genere_zpl(selected_records)
            ecrire_sorties(SORTIES_DIR, fichiers)
        except Exception as e:
            messagebox.showerror("Erreur", f"Génération ZPL : {e}")
            return

        # Journalise en DB
        try:
            cn = connect(DB_PATH)
            with cn:
                for r, (fname, contenu) in zip(selected_records, fichiers):
                    record_print(
                        cn,
                        r.get("Nom", ""),
                        r.get("Prénom", ""),
                        r.get("Date_de_naissance", ""),
                        r.get("Expire_le", ""),
                        contenu,
                        status="printed",
                    )
        except Exception as e:
            messagebox.showerror("Erreur", f"Enregistrement DB : {e}")
            return

        # Rafraîchit stats + vue
        self.refresh_from_db_stats()
        self.apply_filter()
        self.toast("Impressions enregistrées et grille mise à jour")
        messagebox.showinfo("Succès", f"{len(selected_records)} étiquette(s) générée(s) dans {SORTIES_DIR}.")

    def toast(self, msg: str):
        self.status.set(msg)


if __name__ == "__main__":
    # Prépare l'arborescence
    (ROOT / "data" / "sorties").mkdir(parents=True, exist_ok=True)
    app = App()
    app.mainloop()
