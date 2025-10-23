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

from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable

import sqlite3
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from src.app.config import (
    ATTESTATIONS_DIR,
    DEFAULT_EXPIRATION,
    DB_PATH,
    ROOT,
    SORTIES_DIR,
    load_config,
)
from src.app.db import connect, fetch_latest_contact, init_db, record_print
from src.app.imports import (
    build_ddn_lookup_from_rows,
    import_already_printed_csv,
    load_last_import,
    parse_validation_three_line_file,
    persist_last_import,
)
from src.app.validation import (
    build_validation_lookup,
    compute_validation_status,
    find_latest_validation_export,
    load_latest_expiration_by_person,
    load_validation_export,
    parse_validator_names,
)
try:  # Compatibilité exécutable PyInstaller : l'import peut varier selon le contexte.
    from src.app.io_utils import lire_tableau, normalize_name
except ModuleNotFoundError:  # pragma: no cover - dépend du packaging Windows
    try:
        from app.io_utils import lire_tableau, normalize_name  # type: ignore
    except ModuleNotFoundError:  # pragma: no cover - dernier recours
        from io_utils import lire_tableau, normalize_name  # type: ignore
from src.app.attestations import (
    AttestationData,
    generate_attestation_pdf,
    load_attestation_settings,
    send_attestation_email,
)
from src.app.printing import print_ql570_direct
from src.app.zpl import ecrire_sorties, genere_zpl


__all__ = ["App"]


class App(tk.Tk):
    def __init__(
        self,
        *,
        config_loader: Callable[[], dict[str, str]] = load_config,
        csv_importer: Callable[[Path, str, dict[tuple[str, str], str | None] | None], tuple[int, int]] = import_already_printed_csv,
        ddn_lookup_builder: Callable[[Iterable[dict]], dict[tuple[str, str], str | None]] = build_ddn_lookup_from_rows,
        printer: Callable[..., None] = print_ql570_direct,
        db_path: Path = DB_PATH,
        sorties_dir: Path = SORTIES_DIR,
        default_expiration: str = DEFAULT_EXPIRATION,
    ):
        super().__init__()
        self.title("Étiquettes – GUI")
        self.geometry("1180x650")
        self.minsize(1000, 560)

        self._config_loader = config_loader
        self._csv_importer = csv_importer
        self._ddn_lookup_builder = ddn_lookup_builder
        self._printer = printer
        self.db_path = Path(db_path)
        self.sorties_dir = Path(sorties_dir)
        self.attestations_dir = Path(ATTESTATIONS_DIR)
        self.default_expiration = default_expiration
        self._attestation_sender = send_attestation_email

        self.cfg = self._config_loader()
        self.expiration_default_value = self.cfg.get("default_expire") or self.default_expiration
        self._validators = parse_validator_names(self.cfg.get("ffessm_validators") or "")

        self.validation_rows: list[dict[str, str]] = []
        self._validation_lookup: dict[tuple[str, str], dict[str, str]] = {}
        self._status_colors: dict[str, str] = {}
        self._latest_validation_path: Path | None = None

        self._load_latest_validation_export()

        # État
        self.rows: list[dict] = []      # lignes importées (avec DDN/Expire cachés)
        self.view_rows: list[dict] = [] # lignes affichées après filtre
        self.checked: set[int] = set()  # indices cochés dans view_rows
        self.sort_col: str | None = None
        self.sort_asc: bool = True

        self._build_ui()
        self._init_status_styles()
        # Gestion fermeture propre (croix)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.mode_var = tk.StringVar(value="tout")  # "a_imprimer" | "deja" | "tout"
        self.per_expire_count = {}  # (nom, prenom, ddn, expire) -> nb 'printed' pour CETTE expiration
        self._splash_window: tk.Toplevel | None = None
        self._splash_image = None

        if self.expiration_default_value:
            self.exp_var.set(self.expiration_default_value)

        self.after_idle(self._show_splash_screen)
        self._load_last_import_if_available()

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
        
        ttk.Button(top, text="Fichier Profil", command=self.on_import).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(
            top,
            text="Fichier Validation",
            command=self.on_import_validation_file,
        ).pack(side=tk.LEFT, padx=(6, 0))
        if self.cfg.get("backend") != "win32print":
            ttk.Button(top, text="Imprimer ZPL", command=self.on_print).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(top, text="Imprimer QL-570", command=self.on_print_ql570).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(
            top,
            text="Envoyer Mail Attestation",
            command=self.on_send_attestation,
        ).pack(side=tk.LEFT, padx=(6, 0))

        ttk.Label(top, text="Expiration à garder :").pack(side=tk.LEFT, padx=(16, 4))
        self.exp_var = tk.StringVar(value=self.default_expiration)
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
        show_reset = (self.cfg.get("show_reset_db_button") or "").strip().lower() == "true"
        if show_reset:
            ttk.Button(top, text="Réinitialiser la base", command=self.on_reset_db).pack(side=tk.LEFT, padx=(6, 0))

        mid = ttk.Frame(self)
        mid.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        cols = ("sel", "Nom", "Prénom", "Derniere", "Compteur", "ErreurValide")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="extended")
        self.tree.heading("sel", text="✓", command=self.toggle_all)
        self.tree.heading("Nom", text="Nom", command=lambda: self.sort_by("Nom"))
        self.tree.heading("Prénom", text="Prénom", command=lambda: self.sort_by("Prénom"))
        self.tree.heading("Derniere", text="Dernière impression")
        self.tree.heading("Compteur", text="# Impressions")
        self.tree.heading(
            "ErreurValide",
            text="Erreur valide",
            command=lambda: self.sort_by("ErreurValide"),
        )

        self.tree.column("sel", width=48, anchor=tk.CENTER, stretch=False)
        self.tree.column("Nom", width=240)
        self.tree.column("Prénom", width=240)
        self.tree.column("Derniere", width=200, anchor=tk.CENTER)
        self.tree.column("Compteur", width=130, anchor=tk.E)
        self.tree.column("ErreurValide", width=120, anchor=tk.CENTER, stretch=False)

        self.tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=sb.set)
        sb.pack(fill=tk.Y, side=tk.RIGHT)

        self.tree.bind("<Button-1>", self.on_tree_click)

        self.status = tk.StringVar(value="Prêt")
        ttk.Label(self, textvariable=self.status, anchor=tk.W).pack(fill=tk.X, padx=8, pady=(0, 6))

        self._update_headers()

    def _init_status_styles(self):
        self._status_colors = {
            "status-red": "#c0392b",
            "status-orange": "#f39c12",
            "status-green": "#27ae60",
        }
        status_column = "ErreurValide"
        for tag, color in self._status_colors.items():
            # Configure the tag so only the "Erreur Valide" column text is tinted when supported.
            column_spec = f"{{{status_column} {color}}}"
            try:
                self.tree.tag_configure(tag, foreground=column_spec)
            except tk.TclError:
                # Older Tk versions (< 8.7) do not support per-column colors. Fall back to
                # classic colouring so that the UI stays functional instead of crashing.
                self.tree.tag_configure(tag, foreground=color)

    def _load_latest_validation_export(self, *, silent: bool = True):
        try:
            latest = find_latest_validation_export()
        except Exception as exc:
            self.validation_rows = []
            self._validation_lookup = {}
            self._latest_validation_path = None
            if not silent:
                self.toast(f"Validation : {exc}")
            return

        if not latest:
            self.validation_rows = []
            self._validation_lookup = {}
            self._latest_validation_path = None
            return

        try:
            rows = load_validation_export(latest)
        except Exception as exc:
            self.validation_rows = []
            self._validation_lookup = {}
            self._latest_validation_path = latest
            if not silent:
                self.toast(f"Validation : {exc}")
            return

        self.validation_rows = rows
        self._validation_lookup = build_validation_lookup(rows)
        self._latest_validation_path = latest

    def _set_validation_rows(self, rows: list[dict[str, str]], export_path: Path | None = None):
        self.validation_rows = rows or []
        self._validation_lookup = build_validation_lookup(self.validation_rows)
        if export_path:
            self._latest_validation_path = Path(export_path)
        self.refresh_from_db_stats()
        if self.rows:
            self.apply_filter()

    def _show_splash_screen(self):
        splash_cfg = (self.cfg.get("splash_image") or "").strip()
        if not splash_cfg:
            return

        splash_path = Path(splash_cfg)
        if not splash_path.is_absolute():
            splash_path = ROOT / splash_path

        if not splash_path.exists():
            self.toast(f"Illustration introuvable: {splash_path}")
            return

        try:
            from PIL import Image, ImageTk  # type: ignore
        except Exception as exc:  # pragma: no cover - optional dependency at runtime
            self.toast(f"Impossible d'afficher l'illustration (Pillow manquant): {exc}")
            return

        try:
            image = Image.open(splash_path)
        except Exception as exc:  # pragma: no cover - depends on local file
            self.toast(f"Illustration illisible: {exc}")
            return

        splash = tk.Toplevel(self)
        splash.title("Bienvenue")
        splash.transient(self)
        splash.resizable(False, False)

        try:
            photo = ImageTk.PhotoImage(image)
        except Exception as exc:  # pragma: no cover - depends on Tk capabilities
            splash.destroy()
            self.toast(f"Affichage de l'illustration impossible: {exc}")
            return

        label = ttk.Label(splash, image=photo)
        label.image = photo  # Conserve une référence pour Tkinter
        label.pack()

        self._splash_window = splash
        self._splash_image = photo

        splash.update_idletasks()
        self.update_idletasks()

        try:
            if not self.winfo_viewable():
                self.wait_visibility()
        except Exception:
            # La fenêtre principale peut ne pas encore être mappée (ex: premier affichage)
            # Dans ce cas, on s'appuie sur les dimensions requises/écran ci-dessous.
            pass

        self.update_idletasks()

        width = splash.winfo_width()
        height = splash.winfo_height()

        parent_w = self.winfo_width()
        parent_h = self.winfo_height()
        if parent_w <= 1 or parent_h <= 1:
            parent_w = max(parent_w, self.winfo_reqwidth())
            parent_h = max(parent_h, self.winfo_reqheight())
        if parent_w <= 1 or parent_h <= 1:
            parent_w = self.winfo_screenwidth()
            parent_h = self.winfo_screenheight()

        parent_x = self.winfo_rootx()
        parent_y = self.winfo_rooty()
        if parent_x <= 0 and parent_y <= 0:
            screen_w = self.winfo_screenwidth()
            screen_h = self.winfo_screenheight()
            parent_x = max((screen_w - parent_w) // 2, 0)
            parent_y = max((screen_h - parent_h) // 2, 0)

        x = parent_x + max((parent_w - width) // 2, 0)
        y = parent_y + max((parent_h - height) // 2, 0)
        splash.geometry(f"{width}x{height}+{x}+{y}")

        splash.after(3000, self._hide_splash_screen)

    def _hide_splash_screen(self):
        if self._splash_window is None:
            return
        try:
            self._splash_window.destroy()
        except Exception:
            pass
        finally:
            self._splash_window = None
            self._splash_image = None

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
        rotate = self.cfg.get("rotate", "0")

        # Sur Windows, tu mettras :
        # backend_name = "win32print; device_name = "Brother QL-570"

        old_cursor = self["cursor"]
        self.config(cursor="watch")
        self.toast("Impression QL-570 en cours…")
        self.update_idletasks()

        try:
            with connect(self.db_path) as cn:
                for r in selected:
                    nom = (r.get("Nom") or "").strip()
                    prenom = (r.get("Prénom") or "").strip()
                    ddn = (r.get("Date_de_naissance") or "").strip()
                    expire_val = (r.get("Expire_le") or "").strip()
                    email = (r.get("Email") or "").strip()

                    self._printer(
                        nom,
                        prenom,
                        ddn,
                        expire_val,
                        label=label,
                        backend_name=backend_name,
                        device=device_name,
                        rotate=rotate,
                    )

                    record_print(
                        cn,
                        nom,
                        prenom,
                        ddn,
                        expire_val,
                        email,
                        zpl=None,
                        status="printed",
                    )

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

    def on_send_attestation(self):
        if not self.view_rows:
            messagebox.showinfo("Info", "Aucune donnée à traiter.")
            return
        if not self.checked:
            messagebox.showinfo("Info", "Sélectionnez au moins une ligne (colonne ✓).")
            return

        settings = load_attestation_settings(self.cfg)
        if not settings.is_configured:
            messagebox.showerror(
                "Configuration email",
                "Configuration SMTP incomplète (voir config.ini).",
            )
            return

        selected = [self.view_rows[idx] for idx in sorted(self.checked)]
        if not selected:
            messagebox.showinfo("Info", "Aucune ligne sélectionnée.")
            return

        old_cursor = self["cursor"]
        self.config(cursor="watch")
        self.update_idletasks()

        self.attestations_dir.mkdir(parents=True, exist_ok=True)

        successes = 0
        failures: list[str] = []

        try:
            cn = connect(self.db_path)
        except Exception as exc:
            self.config(cursor=old_cursor)
            self.update_idletasks()
            messagebox.showerror("Erreur", f"Connexion DB : {exc}")
            return

        try:
            for row in selected:
                nom = (row.get("Nom") or "").strip()
                prenom = (row.get("Prénom") or "").strip()
                ddn = (row.get("Date_de_naissance") or "").strip()
                expire = (row.get("Expire_le") or "").strip()
                montant = str(row.get("Montant") or "").strip()
                if not montant:
                    failures.append(f"{nom} {prenom} : montant introuvable dans le fichier CSV.")
                    continue

                contact = fetch_latest_contact(cn, nom, prenom, ddn)
                email = ""
                nom_bdd = nom
                prenom_bdd = prenom
                if contact:
                    nom_bdd = contact.nom or nom
                    prenom_bdd = contact.prenom or prenom
                    email = contact.email.strip()

                if not email:
                    email = (row.get("Email") or "").strip()

                if not email:
                    failures.append(f"{nom} {prenom} : adresse e-mail introuvable dans la base.")
                    continue

                data = AttestationData(
                    nom=nom_bdd or nom,
                    prenom=prenom_bdd or prenom,
                    email=email,
                    montant=montant,
                    expire=expire,
                    date_de_naissance=ddn,
                )

                try:
                    pdf_path = generate_attestation_pdf(self.attestations_dir, data)
                    self._attestation_sender(settings, data, pdf_path)
                except Exception as exc:  # pragma: no cover - dépend des backends SMTP
                    failures.append(f"{data.prenom} {data.nom} : {exc}")
                    continue

                successes += 1
        finally:
            cn.close()
            self.config(cursor=old_cursor)
            self.update_idletasks()

        if successes:
            self.toast(f"{successes} attestation(s) envoyée(s)")

        if failures and successes:
            details = "\n".join(failures)
            messagebox.showwarning(
                "Envoi d'attestation",
                f"{successes} attestation(s) envoyée(s).\n\nÉchecs :\n{details}",
            )
        elif failures:
            messagebox.showerror("Envoi d'attestation", "\n".join(failures))
        elif successes:
            messagebox.showinfo("Succès", f"{successes} attestation(s) envoyée(s).")
    def on_reset_db(self):
        ok = messagebox.askyesno(
            "Confirmation",
            "Réinitialiser la base ?",
            icon=messagebox.WARNING,
            detail="Cette action SUPPRIME le fichier data/app.db puis recrée le schéma.",
        )
        if not ok:
            return
        try:
            db_file = self.db_path
            if db_file.exists():
                db_file.unlink()  # supprime la base
            init_db(self.db_path)     # recrée tables / index / vues
            self.refresh_from_db_stats()
            self.apply_filter()
            self.toast("Base réinitialisée.")
            messagebox.showinfo("OK", "Base réinitialisée.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Reset DB : {e}")


    # -------------- Helpers --------------
    def _validation_person_key(self, row: dict) -> tuple[str, str]:
        nom = normalize_name(str(row.get("Nom") or ""))
        prenom = normalize_name(str(row.get("Prénom") or ""))
        return (nom, prenom)

    def _apply_validation_indicators(self, conn):
        if not self.rows:
            return
        if not self._validation_lookup and not self._latest_validation_path:
            for row in self.rows:
                row["ErreurValide"] = ""
            return
        db_lookup = load_latest_expiration_by_person(conn)
        default_expire = self.expiration_default_value or ""
        validators = self._validators
        for row in self.rows:
            key = self._validation_person_key(row)
            status = compute_validation_status(
                key,
                db_lookup,
                self._validation_lookup,
                default_expire,
                validators,
            )
            row["ErreurValide"] = status

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

    def _normalize_erreur_valide(self, value) -> str:
        if value is None:
            return ""
        if isinstance(value, bool):
            return "green" if value else "red"
        v = str(value).strip().lower()
        if not v or v in {"none", "null"}:
            return ""
        mapping = {
            "question": "question",
            "unknown": "question",
            "?": "question",
            "red": "red",
            "error": "red",
            "ko": "red",
            "orange": "orange",
            "warning": "orange",
            "amber": "orange",
            "green": "green",
            "ok": "green",
            "valid": "green",
        }
        if v in {"yes", "true", "1", "y", "oui"}:
            return "green"
        if v in {"no", "false", "0", "n", "non"}:
            return "red"
        return mapping.get(v, "")

    def _fmt_erreur_valide(self, value) -> str:
        normalized = self._normalize_erreur_valide(value)
        mapping = {
            "red": "Invalide",
            "orange": "A verifier",
            "green": "Valide",
            "question": "A verifier",
        }
        return mapping.get(normalized, "")

    def _erreur_valide_sort_key(self, value) -> tuple[int, str]:
        normalized = self._normalize_erreur_valide(value)
        order = {"": 0, "question": 1, "red": 2, "orange": 3, "green": 4}
        return (order.get(normalized, 0), normalized)

    def _update_headers(self):
        def label(col: str, base: str) -> str:
            if self.sort_col != col:
                return base
            return base + (" ▲" if self.sort_asc else " ▼")

        self.tree.heading("Nom", text=label("Nom", "Nom"), command=lambda: self.sort_by("Nom"))
        self.tree.heading("Prénom", text=label("Prénom", "Prénom"), command=lambda: self.sort_by("Prénom"))
        self.tree.heading(
            "Derniere",
            text=label("Derniere", "Dernière impression"),
            command=lambda: self.sort_by("Derniere"),
        )
        self.tree.heading(
            "Compteur",
            text=label("Compteur", "# Impressions"),
            command=lambda: self.sort_by("Compteur"),
        )
        self.tree.heading(
            "ErreurValide",
            text=label("ErreurValide", "Erreur valide"),
            command=lambda: self.sort_by("ErreurValide"),
        )

    def _status_tag_for(self, status_key: str) -> str | None:
        if not status_key:
            return None
        tag_mapping = {
            "green": "status-green",
            "red": "status-red",
            "orange": "status-orange",
            "question": "status-orange",
        }
        return tag_mapping.get(status_key)

    def _persist_last_import(self, source: Path):
        try:
            metadata = persist_last_import(source)
        except Exception as exc:  # pragma: no cover - best-effort persistence
            self.toast(f"Sauvegarde dernier import impossible: {exc}")
            return

        label = metadata.get("source_name") or source.name
        self.toast(f"Fichier chargé: {label} ({len(self.rows)} lignes)")

    def _load_last_import_if_available(self):
        try:
            rows, metadata = load_last_import()
        except Exception as exc:  # pragma: no cover - best-effort loading
            self.toast(f"Dernier import illisible: {exc}")
            return

        if not rows:
            return

        self.rows = rows
        self.refresh_from_db_stats()
        self.apply_filter()
        name = metadata.get("source_name") or metadata.get("cached_name") or "Dernier import"
        self.toast(f"Dernier import rechargé: {name} ({len(self.rows)} lignes)")

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
            "Initialiser / mettre à jour la base SQLite ?",
            icon=messagebox.WARNING,
            detail=(
                "Cette opération crée / met à jour le schéma (tables, index, vues)\n"
                "sans supprimer les données existantes."
            ),
        )
        if not ok:
            self.toast("Init DB annulé")
            return
        try:
            init_db(self.db_path)

            # Auto-import si on trouve deja_imprimes.csv à la racine
            try:

                csv_init = ROOT / self.cfg.get("auto_import_file", "deja_imprimes.csv")
                exp = (self.exp_var.get() or "").strip()
                imp = 0
                skip = 0
                if csv_init.exists() and exp:
                    # construit un lookup DDN à partir des données déjà chargées dans la GUI
                    ddn_lookup = self._ddn_lookup_builder(self.rows)
                    imp, skip = self._csv_importer(
                        csv_init,
                        exp,
                        rows_ddn_lookup=ddn_lookup,
                        db_path=self.db_path,
                    )
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
            self._persist_last_import(Path(path))
        except Exception as e:
            messagebox.showerror("Erreur", f"Import: {e}")

    def on_import_validation_file(self):
        path = filedialog.askopenfilename(
            title="Choisir un fichier Validation",
            filetypes=[
                ("Fichiers CSV", "*.csv"),
                ("Fichiers Excel", "*.xlsx *.xls"),
                ("Tous", "*.*"),
            ],
        )
        if not path:
            return

        try:
            result = parse_validation_three_line_file(path)
        except Exception as exc:
            messagebox.showerror("Erreur", f"Import Validation : {exc}")
            return

        count = len(result.rows)
        if count:
            self.toast(f"{count} membre(s) Validation exporté(s)")
        else:
            self.toast("Aucune entrée trouvée dans le fichier Validation")

        export_path = result.export_path
        if export_path:
            try:
                loaded_rows = load_validation_export(export_path)
            except Exception as exc:
                self.toast(f"Validation : {exc}")
                self._set_validation_rows(result.rows, export_path=export_path)
            else:
                self._set_validation_rows(loaded_rows, export_path=export_path)
        else:
            self._set_validation_rows(result.rows, export_path=None)
        if export_path:
            if count:
                message = (
                    f"{count} membre(s) exporté(s) dans {export_path}."
                )
            else:
                message = (
                    f"Aucune entrée détectée. Fichier généré : {export_path}."
                )
            messagebox.showinfo("Import Validation", message)

    def refresh_from_db_stats(self):
        """Complète chaque row avec Derniere/Compteur (toutes expirations) + map per-expire."""
        try:
            cn = connect(self.db_path)
        except Exception:
            return

        try:
            # ---- Stats par personne (toutes expirations confondues) ----
            stats_person = {}
            self.per_expire_count = {}
            try:
                cur = cn.execute("SELECT nom, prenom, ddn, last_print, cnt FROM v_person_stats")
            except sqlite3.OperationalError:
                try:
                    cur = cn.execute(
                        """
                        SELECT
                            nom, prenom, ddn,
                            MAX(CASE WHEN status='printed' THEN printed_at END) AS last_print,
                            SUM(CASE WHEN status='printed' THEN 1 ELSE 0 END)   AS cnt
                        FROM prints
                        GROUP BY nom, prenom, ddn
                        """
                    )
                except sqlite3.OperationalError:
                    self.toast("Base SQLite non initialisée. Utilisez Init DB.")
                    return
            except Exception:
                cur = cn.execute(
                    """
                    SELECT
                        nom, prenom, ddn,
                        MAX(CASE WHEN status='printed' THEN printed_at END) AS last_print,
                        SUM(CASE WHEN status='printed' THEN 1 ELSE 0 END)   AS cnt
                    FROM prints
                    GROUP BY nom, prenom, ddn
                    """
                )
            for row in cur.fetchall():
                key = (
                    (row["nom"] or "").strip().lower(),
                    (row["prenom"] or "").strip().lower(),
                    (row["ddn"] or "").strip(),
                )
                stats_person[key] = (row["last_print"], int(row["cnt"] or 0))

            # ---- Compteur par personne+expiration (pour filtrer À imprimer / Déjà imprimées) ----
            per_expire = {}
            try:
                cur2 = cn.execute(
                    """
                    SELECT nom, prenom, ddn, expire,
                           SUM(CASE WHEN status='printed' THEN 1 ELSE 0 END) AS cnt_exp
                    FROM prints
                    GROUP BY nom, prenom, ddn, expire
                    """
                )
            except sqlite3.OperationalError:
                self.toast("Base SQLite non initialisée. Utilisez Init DB.")
                return
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

            self._apply_validation_indicators(cn)
        finally:
            cn.close()


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
            sel_txt = "☑" if idx in self.checked else "☐"
            values = (
                sel_txt,
                str(r.get("Nom", "") or ""),
                str(r.get("Prénom", "") or ""),
                self._fmt_dt(r.get("Derniere")) or "",
                str(r.get("Compteur", 0) or 0),
                self._fmt_erreur_valide(r.get("ErreurValide")),
            )
            status_key = self._normalize_erreur_valide(r.get("ErreurValide"))
            tag = self._status_tag_for(status_key)
            tags = (tag,) if tag else ()
            insert_kwargs = {"text": "", "values": values, "tags": tags}
            self.tree.insert("", tk.END, iid=str(idx), **insert_kwargs)
        self.status.set(f"Affichées: {len(self.view_rows)} (sélectionnées: {len(self.checked)})")

    def on_tree_click(self, event):
        # Clic sur la colonne "sel" (#1) pour cocher/décocher
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
            if col == "ErreurValide":
                return self._erreur_valide_sort_key(v)
            if isinstance(v, str):
                return (v or "").lower()
            return v

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
            ecrire_sorties(self.sorties_dir, fichiers)
        except Exception as e:
            messagebox.showerror("Erreur", f"Génération ZPL : {e}")
            return

        # Journalise en DB
        try:
            with connect(self.db_path) as cn:
                for r, (fname, contenu) in zip(selected_records, fichiers):
                    record_print(
                        cn,
                        r.get("Nom", ""),
                        r.get("Prénom", ""),
                        r.get("Date_de_naissance", ""),
                        r.get("Expire_le", ""),
                        r.get("Email", ""),
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
        messagebox.showinfo(
            "Succès",
            f"{len(selected_records)} étiquette(s) générée(s) dans {self.sorties_dir}.",
        )

    def toast(self, msg: str):
        self.status.set(msg)


if __name__ == "__main__":
    SORTIES_DIR.mkdir(parents=True, exist_ok=True)
    ATTESTATIONS_DIR.mkdir(parents=True, exist_ok=True)
    app = App()
    app.mainloop()
