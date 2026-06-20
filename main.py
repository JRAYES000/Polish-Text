"""
TextEnhancer AI
===============
Petit utilitaire Windows qui vit dans la barre des tâches.

Principe :
  1. Tu sélectionnes du texte dans n'importe quelle application.
  2. Tu presses ton raccourci global (Alt+Q par défaut).
  3. Le texte sélectionné est capturé, envoyé à OpenRouter (LLM au choix)
     avec l'instruction (« preset ») que tu choisis.
  4. Une fenêtre d'aperçu affiche le résultat : tu peux l'éditer, le
     regénérer, changer de preset/modèle, puis le COLLER en texte enrichi
     (vrai gras pour Word / Outlook) ou le copier.

Tout est configurable dans la fenêtre Paramètres : clé API OpenRouter,
modèle par défaut, presets (nom + instruction), raccourci clavier,
démarrage avec Windows.

Auteur : généré pour Julien. Licence : usage personnel.
"""

import os
import sys
import json
import time
import threading
import traceback
import tempfile
import subprocess
import webbrowser

# --- Dépendances tierces -----------------------------------------------------
import requests
import keyboard                       # raccourcis globaux + simulation de touches
import markdown as md_lib             # markdown -> HTML
import pystray
from PIL import Image, ImageDraw

import tkinter as tk
from tkinter import ttk, messagebox

# --- Spécifique Windows ------------------------------------------------------
import win32clipboard
import win32con
import win32gui
import winreg


APP_NAME = "TextEnhancer AI"
APP_DIR_NAME = "TextEnhancerAI"
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
OPENROUTER_MODELS_URL = "https://openrouter.ai/api/v1/models"

# --- Version & mise à jour automatique ---------------------------------------
# Dépôt GitHub utilisé pour les mises à jour (modifiable aussi dans
# Paramètres → Dépôt GitHub, sans recompiler).
APP_VERSION = "1.0.1"
GITHUB_REPO = "JRAYES000/Polish-Text"
GITHUB_API_LATEST = "https://api.github.com/repos/{repo}/releases/latest"


# =============================================================================
#  Configuration
# =============================================================================
def config_dir():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    path = os.path.join(base, APP_DIR_NAME)
    os.makedirs(path, exist_ok=True)
    return path


def config_path():
    return os.path.join(config_dir(), "config.json")


DEFAULT_CONFIG = {
    "api_key": "",
    "default_model": "qwen/qwen3.7-plus",
    "hotkey": "alt+q",
    "github_repo": GITHUB_REPO,
    "check_updates_on_start": True,
    "known_models": [
        "qwen/qwen3.7-plus",
        "anthropic/claude-opus-4.8",
        "anthropic/claude-sonnet-4.6",
        "google/gemini-2.5-flash",
        "openai/gpt-5",
    ],
    "presets": [
        {
            "name": "Reformuler (clair + gras)",
            "instruction": (
                "Tu es un assistant de rédaction. Reformule le texte fourni "
                "pour le rendre plus clair, fluide et professionnel. Structure "
                "le contenu si pertinent (paragraphes, listes), et mets en gras "
                "(syntaxe Markdown **gras**) les points clés. Réponds uniquement "
                "avec le texte reformulé, sans commentaire ni introduction. "
                "Conserve la langue d'origine."
            ),
        },
        {
            "name": "Résumer",
            "instruction": (
                "Résume le texte fourni en quelques points essentiels sous forme "
                "de liste à puces Markdown, en mettant en gras les idées clés. "
                "Réponds uniquement avec le résumé, dans la langue d'origine."
            ),
        },
        {
            "name": "Email professionnel",
            "instruction": (
                "Transforme le texte fourni en un email professionnel, poli et "
                "bien structuré (formule d'appel, corps clair, formule de "
                "politesse). Mets en gras les éléments importants. Réponds "
                "uniquement avec l'email, dans la langue d'origine."
            ),
        },
        {
            "name": "Corriger l'orthographe",
            "instruction": (
                "Corrige uniquement l'orthographe, la grammaire et la ponctuation "
                "du texte fourni, sans en changer le sens ni le style. Réponds "
                "uniquement avec le texte corrigé, dans la langue d'origine."
            ),
        },
    ],
}


def load_config():
    path = config_path()
    if not os.path.exists(path):
        save_config(DEFAULT_CONFIG)
        return json.loads(json.dumps(DEFAULT_CONFIG))
    try:
        with open(path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except Exception:
        cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    # Complète les clés manquantes (migrations douces)
    for k, v in DEFAULT_CONFIG.items():
        cfg.setdefault(k, v)
    if not cfg.get("presets"):
        cfg["presets"] = json.loads(json.dumps(DEFAULT_CONFIG["presets"]))
    return cfg


def save_config(cfg):
    with open(config_path(), "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# =============================================================================
#  Presse-papiers (lecture texte + écriture texte enrichi via CF_HTML)
# =============================================================================
def get_clipboard_text(retries=10, delay=0.05):
    """Lit le presse-papiers (texte) avec quelques tentatives (il peut être
    verrouillé brièvement par une autre application)."""
    for _ in range(retries):
        try:
            win32clipboard.OpenClipboard()
            try:
                if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                    data = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
                    return data or ""
                return ""
            finally:
                win32clipboard.CloseClipboard()
        except Exception:
            time.sleep(delay)
    return ""


def set_clipboard_text(text, retries=10, delay=0.05):
    for _ in range(retries):
        try:
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
                return True
            finally:
                win32clipboard.CloseClipboard()
        except Exception:
            time.sleep(delay)
    return False


def _build_cf_html(fragment_html):
    """Encapsule un fragment HTML dans l'en-tête « HTML Format » exigé par
    Windows (CF_HTML) avec les offsets en octets corrects."""
    header_tpl = (
        "Version:0.9\r\n"
        "StartHTML:{:09d}\r\n"
        "EndHTML:{:09d}\r\n"
        "StartFragment:{:09d}\r\n"
        "EndFragment:{:09d}\r\n"
    )
    html_pre = "<html><body>\r\n<!--StartFragment-->"
    html_post = "<!--EndFragment-->\r\n</body></html>"

    header_len = len(header_tpl.format(0, 0, 0, 0).encode("utf-8"))
    start_html = header_len
    start_fragment = start_html + len(html_pre.encode("utf-8"))
    end_fragment = start_fragment + len(fragment_html.encode("utf-8"))
    end_html = end_fragment + len(html_post.encode("utf-8"))

    header = header_tpl.format(start_html, end_html, start_fragment, end_fragment)
    return header + html_pre + fragment_html + html_post


def set_clipboard_rich(markdown_text):
    """Met dans le presse-papiers une version HTML (texte enrichi : gras,
    listes...) ET une version texte brut de secours."""
    html_fragment = md_lib.markdown(
        markdown_text, extensions=["extra", "sane_lists", "nl2br"]
    )
    cf_html_payload = _build_cf_html(html_fragment).encode("utf-8")
    plain = markdown_to_plain(markdown_text)

    for _ in range(10):
        try:
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, plain)
                cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
                win32clipboard.SetClipboardData(cf_html, cf_html_payload)
                return True
            finally:
                win32clipboard.CloseClipboard()
        except Exception:
            time.sleep(0.05)
    return False


def markdown_to_plain(text):
    """Version texte brut « propre » : enlève les marqueurs Markdown les plus
    courants pour les applications qui n'acceptent pas le texte enrichi."""
    import re
    t = text
    t = re.sub(r"\*\*(.+?)\*\*", r"\1", t)   # **gras**
    t = re.sub(r"__(.+?)__", r"\1", t)       # __gras__
    t = re.sub(r"\*(.+?)\*", r"\1", t)       # *italique*
    t = re.sub(r"`{1,3}(.+?)`{1,3}", r"\1", t)  # `code`
    t = re.sub(r"^#{1,6}\s*", "", t, flags=re.MULTILINE)  # titres
    t = re.sub(r"^\s*[-*+]\s+", "• ", t, flags=re.MULTILINE)  # puces
    return t


# =============================================================================
#  Capture de la sélection / collage dans l'application cible
# =============================================================================
def capture_selection():
    """Récupère le texte sélectionné. On tente une copie automatique (Ctrl+C) ;
    si elle ne renvoie rien, on se rabat sur le contenu actuel du presse-papiers
    (cas où l'utilisateur a déjà fait Ctrl+C lui-même). On ne vide JAMAIS le
    presse-papiers, pour ne pas perdre une copie manuelle."""
    # Relâche les touches du raccourci pour ne pas parasiter le Ctrl+C.
    for k in ("alt", "ctrl", "shift", "windows", "q", "r", "e", "w"):
        try:
            keyboard.release(k)
        except Exception:
            pass
    time.sleep(0.12)

    before = get_clipboard_text()
    try:
        keyboard.send("ctrl+c")
    except Exception:
        pass

    # Attend qu'une nouvelle sélection soit copiée (jusqu'à ~1,2 s).
    deadline = time.time() + 1.2
    while time.time() < deadline:
        time.sleep(0.08)
        current = get_clipboard_text()
        if current and current != before:
            return current

    # Rien de neuf : on garde ce qui était déjà dans le presse-papiers.
    return get_clipboard_text() or before


def paste_into(target_hwnd, rich=True, text=""):
    """Redonne le focus à l'application cible puis colle (Ctrl+V)."""
    if rich:
        set_clipboard_rich(text)
    else:
        set_clipboard_text(text)
    time.sleep(0.05)
    if target_hwnd:
        try:
            win32gui.SetForegroundWindow(target_hwnd)
        except Exception:
            pass
    time.sleep(0.15)
    keyboard.send("ctrl+v")


# =============================================================================
#  Client OpenRouter
# =============================================================================
def call_openrouter(api_key, model, instruction, user_text, timeout=90):
    if not api_key:
        raise RuntimeError("Aucune clé API OpenRouter n'est configurée "
                           "(Paramètres → Clé API).")
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://localhost/textenhancer",
        "X-Title": APP_NAME,
    }
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": instruction},
            {"role": "user", "content": user_text},
        ],
    }
    resp = requests.post(OPENROUTER_URL, headers=headers,
                         json=payload, timeout=timeout)
    if resp.status_code != 200:
        raise RuntimeError(f"Erreur OpenRouter {resp.status_code} : {resp.text[:500]}")
    data = resp.json()
    try:
        return data["choices"][0]["message"]["content"].strip()
    except (KeyError, IndexError):
        raise RuntimeError(f"Réponse inattendue d'OpenRouter : {data}")


def fetch_models(api_key, timeout=30):
    headers = {"Authorization": f"Bearer {api_key}"} if api_key else {}
    resp = requests.get(OPENROUTER_MODELS_URL, headers=headers, timeout=timeout)
    resp.raise_for_status()
    data = resp.json()
    return sorted(m["id"] for m in data.get("data", []) if "id" in m)


# =============================================================================
#  Mise à jour automatique (via GitHub Releases)
# =============================================================================
def _parse_version(v):
    """'v1.2.0' -> (1, 2, 0). Robuste aux suffixes non numériques."""
    v = (v or "").strip().lstrip("vV")
    parts = []
    for p in v.split("."):
        digits = "".join(ch for ch in p if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts) if parts else (0,)


def is_newer(latest_tag, current=APP_VERSION):
    try:
        a = _parse_version(latest_tag)
        b = _parse_version(current)
        # Égalise les longueurs pour comparer proprement.
        n = max(len(a), len(b))
        a = a + (0,) * (n - len(a))
        b = b + (0,) * (n - len(b))
        return a > b
    except Exception:
        return False


def get_latest_release(repo, timeout=20):
    """Renvoie (tag, url_exe, url_page) de la dernière release GitHub."""
    url = GITHUB_API_LATEST.format(repo=repo)
    resp = requests.get(url, timeout=timeout,
                        headers={"Accept": "application/vnd.github+json"})
    if resp.status_code == 404:
        raise RuntimeError("Aucune release publiée sur ce dépôt.")
    if resp.status_code != 200:
        raise RuntimeError(f"GitHub {resp.status_code} : {resp.text[:200]}")
    data = resp.json()
    tag = data.get("tag_name", "")
    asset_url = None
    for a in data.get("assets", []):
        if a.get("name", "").lower().endswith(".exe"):
            asset_url = a.get("browser_download_url")
            break
    return tag, asset_url, data.get("html_url")


def download_and_apply_update(asset_url):
    """Télécharge le nouvel .exe et programme le remplacement de l'exe en cours
    via un script batch (qui attend la fermeture du process, échange le fichier,
    puis relance l'application)."""
    if not getattr(sys, "frozen", False):
        raise RuntimeError("La mise à jour automatique n'est disponible que "
                           "pour l'exécutable (.exe).")
    current_exe = sys.executable
    new_exe = os.path.join(tempfile.gettempdir(), "TextEnhancerAI_new.exe")

    with requests.get(asset_url, stream=True, timeout=180) as r:
        r.raise_for_status()
        with open(new_exe, "wb") as f:
            for chunk in r.iter_content(chunk_size=65536):
                if chunk:
                    f.write(chunk)

    pid = os.getpid()
    bat_path = os.path.join(tempfile.gettempdir(), "TextEnhancerAI_update.bat")
    script = (
        "@echo off\r\n"
        ":waitloop\r\n"
        f'tasklist /FI "PID eq {pid}" 2>nul | find "{pid}" >nul\r\n'
        "if not errorlevel 1 (\r\n"
        "  timeout /t 1 /nobreak >nul\r\n"
        "  goto waitloop\r\n"
        ")\r\n"
        "timeout /t 1 /nobreak >nul\r\n"
        f'move /y "{new_exe}" "{current_exe}" >nul\r\n'
        f'start "" "{current_exe}"\r\n'
        'del "%~f0"\r\n'
    )
    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(script)

    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0) | \
        getattr(subprocess, "DETACHED_PROCESS", 0)
    subprocess.Popen(["cmd", "/c", bat_path], creationflags=creationflags,
                     close_fds=True)


# =============================================================================
#  Application (gère le root Tk caché, l'icône tray, le hotkey)
# =============================================================================
class TextEnhancerApp:
    def __init__(self):
        self.config = load_config()
        self.target_hwnd = None
        self.preview_win = None
        self.settings_win = None
        self.icon = None

        # Root Tk caché : toutes les fenêtres en sont des Toplevel.
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.title(APP_NAME)

        self._register_hotkey()

    # ---- Raccourci global ---------------------------------------------------
    def _register_hotkey(self):
        try:
            keyboard.clear_all_hotkeys()
        except Exception:
            pass
        try:
            keyboard.add_hotkey(self.config["hotkey"], self._on_hotkey)
        except Exception as e:
            print(f"Impossible d'enregistrer le raccourci "
                  f"'{self.config['hotkey']}': {e}")

    def _on_hotkey(self):
        # Capture immédiate (l'app cible est encore au premier plan),
        # puis on bascule sur le thread Tk pour ouvrir la fenêtre.
        self.target_hwnd = win32gui.GetForegroundWindow()
        selection = capture_selection()
        self.root.after(0, lambda: self._open_preview(selection))

    # ---- Fenêtre d'aperçu ---------------------------------------------------
    def _open_preview(self, selection):
        if self.preview_win and self.preview_win.win.winfo_exists():
            self.preview_win.win.lift()
            self.preview_win.set_source(selection)
            return
        self.preview_win = PreviewWindow(self, selection)

    # ---- Fenêtre paramètres -------------------------------------------------
    def open_settings(self):
        if self.settings_win and self.settings_win.win.winfo_exists():
            self.settings_win.win.lift()
            return
        self.settings_win = SettingsWindow(self)

    def manual_trigger(self):
        """Déclenchement depuis le menu tray : utilise le presse-papiers actuel."""
        self.target_hwnd = None
        self.root.after(0, lambda: self._open_preview(get_clipboard_text()))

    # ---- Tray ---------------------------------------------------------------
    def _make_tray_image(self):
        img = Image.new("RGB", (64, 64), (37, 99, 235))
        d = ImageDraw.Draw(img)
        d.rectangle([12, 16, 52, 24], fill="white")
        d.rectangle([28, 16, 36, 50], fill="white")  # un "T" stylisé
        return img

    def run_tray(self):
        menu = pystray.Menu(
            pystray.MenuItem("Améliorer (presse-papiers)",
                             lambda icon, item: self.manual_trigger()),
            pystray.MenuItem("Paramètres",
                             lambda icon, item: self.root.after(0, self.open_settings)),
            pystray.MenuItem("Vérifier les mises à jour",
                             lambda icon, item: self.root.after(
                                 0, lambda: self.check_updates(silent=False))),
            pystray.MenuItem("Quitter", lambda icon, item: self.quit()),
        )
        self.icon = pystray.Icon(APP_NAME, self._make_tray_image(),
                                 APP_NAME, menu)
        # pystray tourne dans un thread démon ; Tk garde le thread principal.
        threading.Thread(target=self.icon.run, daemon=True).start()

    def quit(self):
        try:
            if self.icon:
                self.icon.stop()
        except Exception:
            pass
        try:
            keyboard.clear_all_hotkeys()
        except Exception:
            pass
        self.root.after(0, self.root.destroy)

    def reload_hotkey(self):
        self._register_hotkey()

    # ---- Mise à jour --------------------------------------------------------
    def check_updates(self, silent=True):
        """silent=True : ne dit rien si on est déjà à jour (vérif au démarrage).
        silent=False : affiche aussi 'à jour' / les erreurs (vérif manuelle)."""
        repo = (self.config.get("github_repo") or GITHUB_REPO).strip()
        if not repo or "OWNER/REPO" in repo:
            if not silent:
                messagebox.showinfo(
                    APP_NAME,
                    "Dépôt GitHub non configuré.\nRenseigne-le dans "
                    "Paramètres → Dépôt GitHub (ex. ton-pseudo/text-enhancer-ai).")
            return

        def worker():
            try:
                tag, asset_url, html = get_latest_release(repo)
            except Exception as e:
                if not silent:
                    msg = str(e)
                    self.root.after(0, lambda: messagebox.showerror(
                        APP_NAME, f"Vérification impossible : {msg}"))
                return
            if tag and is_newer(tag):
                self.root.after(0, lambda: self._prompt_update(tag, asset_url, html))
            elif not silent:
                self.root.after(0, lambda: messagebox.showinfo(
                    APP_NAME, f"Tu es déjà à jour (version {APP_VERSION})."))

        threading.Thread(target=worker, daemon=True).start()

    def _prompt_update(self, tag, asset_url, html):
        repo = (self.config.get("github_repo") or GITHUB_REPO).strip()
        if not getattr(sys, "frozen", False):
            messagebox.showinfo(
                APP_NAME,
                f"Nouvelle version {tag} disponible (tu es en v{APP_VERSION}).\n"
                "La mise à jour automatique ne fonctionne que sur l'exécutable. "
                "En mode script, fais un 'git pull'.")
            return
        if not asset_url:
            if messagebox.askyesno(
                    APP_NAME,
                    f"Version {tag} disponible, mais aucun .exe n'y est attaché.\n"
                    "Ouvrir la page des releases ?"):
                webbrowser.open(html or f"https://github.com/{repo}/releases")
            return
        if messagebox.askyesno(
                APP_NAME,
                f"Nouvelle version {tag} disponible (actuelle : v{APP_VERSION}).\n\n"
                "Mettre à jour maintenant ? L'application va se fermer puis "
                "redémarrer automatiquement sur la nouvelle version.\n"
                "(Tes réglages et presets sont conservés.)"):
            try:
                download_and_apply_update(asset_url)
                self.quit()
            except Exception as e:
                messagebox.showerror(APP_NAME, f"Échec de la mise à jour : {e}")

    def start(self):
        self.run_tray()
        if not self.config.get("api_key"):
            # Première utilisation : on ouvre directement les paramètres.
            self.root.after(500, self.open_settings)
        # Vérification des mises à jour au démarrage (en arrière-plan).
        if self.config.get("check_updates_on_start", True):
            self.root.after(2500, lambda: self.check_updates(silent=True))
        self.root.mainloop()


# =============================================================================
#  Fenêtre d'aperçu
# =============================================================================
class PreviewWindow:
    def __init__(self, app, source_text):
        self.app = app
        self.cfg = app.config
        self.source_text = source_text or ""

        self.win = tk.Toplevel(app.root)
        self.win.title(f"{APP_NAME} — Aperçu")
        self.win.geometry("720x560")
        self.win.attributes("-topmost", True)

        preset_names = [p["name"] for p in self.cfg["presets"]]

        # --- Barre du haut : preset + modèle ---
        top = ttk.Frame(self.win, padding=10)
        top.pack(fill="x")

        ttk.Label(top, text="Style :").grid(row=0, column=0, sticky="w")
        self.preset_var = tk.StringVar(value=preset_names[0] if preset_names else "")
        self.preset_cb = ttk.Combobox(top, textvariable=self.preset_var,
                                      values=preset_names, state="readonly",
                                      width=30)
        self.preset_cb.grid(row=0, column=1, sticky="w", padx=(4, 16))

        ttk.Label(top, text="Modèle :").grid(row=0, column=2, sticky="w")
        self.model_var = tk.StringVar(value=self.cfg.get("default_model", ""))
        self.model_cb = ttk.Combobox(top, textvariable=self.model_var,
                                     values=self.cfg.get("known_models", []),
                                     width=28)
        self.model_cb.grid(row=0, column=3, sticky="w", padx=(4, 0))

        # --- Texte source (repliable, lecture) ---
        src_frame = ttk.LabelFrame(self.win, text="Texte sélectionné", padding=6)
        src_frame.pack(fill="x", padx=10, pady=(0, 6))
        self.src_text = tk.Text(src_frame, height=4, wrap="word")
        self.src_text.pack(fill="x")
        self.src_text.insert("1.0", self.source_text)
        self.src_text.configure(state="disabled")

        # --- Résultat (éditable) ---
        res_frame = ttk.LabelFrame(self.win, text="Résultat (éditable)", padding=6)
        res_frame.pack(fill="both", expand=True, padx=10, pady=(0, 6))
        self.result_text = tk.Text(res_frame, wrap="word")
        self.result_text.pack(fill="both", expand=True)

        # --- Statut ---
        self.status_var = tk.StringVar(value="")
        ttk.Label(self.win, textvariable=self.status_var,
                  foreground="#2563eb").pack(fill="x", padx=12)

        # --- Boutons ---
        btns = ttk.Frame(self.win, padding=10)
        btns.pack(fill="x")
        self.gen_btn = ttk.Button(btns, text="Générer / Regénérer",
                                  command=self.generate)
        self.gen_btn.pack(side="left")
        ttk.Button(btns, text="Coller (texte enrichi)",
                   command=lambda: self.paste(rich=True)).pack(side="left", padx=6)
        ttk.Button(btns, text="Coller (brut)",
                   command=lambda: self.paste(rich=False)).pack(side="left")
        ttk.Button(btns, text="Copier",
                   command=self.copy).pack(side="left", padx=6)
        ttk.Button(btns, text="Fermer",
                   command=self.win.destroy).pack(side="right")

        self.win.bind("<Escape>", lambda e: self.win.destroy())

        # Génération automatique au démarrage si on a du texte.
        if self.source_text.strip():
            self.win.after(150, self.generate)
        else:
            self.status_var.set("Aucun texte capturé. Sélectionne du texte puis "
                                "réessaie, ou colle dans le champ ci-dessus.")

    def set_source(self, text):
        self.source_text = text or ""
        self.src_text.configure(state="normal")
        self.src_text.delete("1.0", "end")
        self.src_text.insert("1.0", self.source_text)
        self.src_text.configure(state="disabled")
        if self.source_text.strip():
            self.generate()

    def _current_instruction(self):
        name = self.preset_var.get()
        for p in self.cfg["presets"]:
            if p["name"] == name:
                return p["instruction"]
        return self.cfg["presets"][0]["instruction"]

    def generate(self):
        if not self.source_text.strip():
            self.status_var.set("Aucun texte à traiter.")
            return
        self.gen_btn.configure(state="disabled")
        self.status_var.set("Génération en cours…")
        instruction = self._current_instruction()
        model = self.model_var.get().strip() or self.cfg.get("default_model")
        api_key = self.cfg.get("api_key", "")
        src = self.source_text

        def worker():
            try:
                result = call_openrouter(api_key, model, instruction, src)
                self.win.after(0, lambda: self._show_result(result))
            except Exception as e:
                msg = str(e)
                self.win.after(0, lambda: self._show_error(msg))

        threading.Thread(target=worker, daemon=True).start()

    def _show_result(self, result):
        self.result_text.delete("1.0", "end")
        self.result_text.insert("1.0", result)
        self.status_var.set("Prêt. Édite si besoin puis colle ou copie.")
        self.gen_btn.configure(state="normal")

    def _show_error(self, msg):
        self.status_var.set("Erreur : " + msg)
        self.gen_btn.configure(state="normal")

    def _get_result(self):
        return self.result_text.get("1.0", "end-1c")

    def copy(self):
        text = self._get_result()
        if not text.strip():
            return
        set_clipboard_rich(text)
        self.status_var.set("Copié dans le presse-papiers (texte enrichi).")

    def paste(self, rich=True):
        text = self._get_result()
        if not text.strip():
            return
        hwnd = self.app.target_hwnd
        self.win.destroy()
        # Petit délai pour laisser la fenêtre se fermer avant de coller.
        time.sleep(0.2)
        paste_into(hwnd, rich=rich, text=text)


# =============================================================================
#  Fenêtre paramètres
# =============================================================================
class SettingsWindow:
    def __init__(self, app):
        self.app = app
        self.cfg = app.config

        self.win = tk.Toplevel(app.root)
        self.win.title(f"{APP_NAME} — Paramètres")
        self.win.geometry("760x620")
        self.win.attributes("-topmost", True)

        nb = ttk.Notebook(self.win)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        self._build_general_tab(nb)
        self._build_presets_tab(nb)

        bottom = ttk.Frame(self.win, padding=8)
        bottom.pack(fill="x")
        ttk.Button(bottom, text="Enregistrer",
                   command=self.save).pack(side="right")
        ttk.Button(bottom, text="Fermer",
                   command=self.win.destroy).pack(side="right", padx=6)

    # ---- Onglet général -----------------------------------------------------
    def _build_general_tab(self, nb):
        f = ttk.Frame(nb, padding=12)
        nb.add(f, text="Général")

        ttk.Label(f, text="Clé API OpenRouter :").grid(row=0, column=0, sticky="w", pady=4)
        self.api_var = tk.StringVar(value=self.cfg.get("api_key", ""))
        self.api_entry = ttk.Entry(f, textvariable=self.api_var, width=58, show="•")
        self.api_entry.grid(row=0, column=1, sticky="w", pady=4)
        self.show_key = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="Afficher", variable=self.show_key,
                        command=self._toggle_key).grid(row=0, column=2, padx=6)

        ttk.Label(f, text="Modèle par défaut :").grid(row=1, column=0, sticky="w", pady=4)
        self.model_var = tk.StringVar(value=self.cfg.get("default_model", ""))
        self.model_cb = ttk.Combobox(f, textvariable=self.model_var,
                                     values=self.cfg.get("known_models", []),
                                     width=55)
        self.model_cb.grid(row=1, column=1, sticky="w", pady=4)
        ttk.Button(f, text="Charger la liste des modèles",
                   command=self._refresh_models).grid(row=1, column=2, padx=6)

        ttk.Label(f, text="Raccourci clavier :").grid(row=2, column=0, sticky="w", pady=4)
        self.hotkey_var = tk.StringVar(value=self.cfg.get("hotkey", "alt+q"))
        ttk.Entry(f, textvariable=self.hotkey_var, width=20).grid(
            row=2, column=1, sticky="w", pady=4)
        ttk.Label(f, text="(ex : alt+q, ctrl+alt+r, ctrl+shift+e)",
                  foreground="#666").grid(row=3, column=1, sticky="w")

        self.startup_var = tk.BooleanVar(value=is_startup_enabled())
        ttk.Checkbutton(f, text="Lancer au démarrage de Windows",
                        variable=self.startup_var).grid(
            row=4, column=1, sticky="w", pady=(10, 2))

        ttk.Label(f, text="Dépôt GitHub :").grid(row=5, column=0, sticky="w", pady=4)
        self.repo_var = tk.StringVar(value=self.cfg.get("github_repo", ""))
        ttk.Entry(f, textvariable=self.repo_var, width=40).grid(
            row=5, column=1, sticky="w", pady=4)
        ttk.Label(f, text="(ex : ton-pseudo/text-enhancer-ai — pour les mises à jour)",
                  foreground="#666").grid(row=6, column=1, sticky="w")

        self.updcheck_var = tk.BooleanVar(
            value=self.cfg.get("check_updates_on_start", True))
        ttk.Checkbutton(f, text="Vérifier les mises à jour au démarrage",
                        variable=self.updcheck_var).grid(
            row=7, column=1, sticky="w", pady=(8, 2))

        ttk.Label(f, text=f"Version installée : {APP_VERSION}",
                  foreground="#666").grid(row=8, column=1, sticky="w", pady=(8, 0))

    def _toggle_key(self):
        self.api_entry.configure(show="" if self.show_key.get() else "•")

    def _refresh_models(self):
        try:
            models = fetch_models(self.api_var.get().strip())
            if models:
                self.model_cb.configure(values=models)
                self.cfg["known_models"] = models
                messagebox.showinfo(APP_NAME,
                                    f"{len(models)} modèles chargés.",
                                    parent=self.win)
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Échec du chargement : {e}",
                                 parent=self.win)

    # ---- Onglet presets -----------------------------------------------------
    def _build_presets_tab(self, nb):
        f = ttk.Frame(nb, padding=12)
        nb.add(f, text="Presets (instructions)")

        left = ttk.Frame(f)
        left.pack(side="left", fill="y", padx=(0, 10))
        self.preset_list = tk.Listbox(left, width=28, height=18,
                                      exportselection=False)
        self.preset_list.pack(fill="y", expand=True)
        self.preset_list.bind("<<ListboxSelect>>", self._on_select_preset)
        btnf = ttk.Frame(left)
        btnf.pack(fill="x", pady=4)
        ttk.Button(btnf, text="+ Ajouter",
                   command=self._add_preset).pack(side="left")
        ttk.Button(btnf, text="Supprimer",
                   command=self._del_preset).pack(side="left", padx=4)

        right = ttk.Frame(f)
        right.pack(side="left", fill="both", expand=True)
        ttk.Label(right, text="Nom :").pack(anchor="w")
        self.name_var = tk.StringVar()
        ttk.Entry(right, textvariable=self.name_var, width=50).pack(
            anchor="w", fill="x")
        ttk.Label(right, text="Instruction (prompt système) :").pack(
            anchor="w", pady=(8, 0))
        self.instr_text = tk.Text(right, wrap="word", height=16)
        self.instr_text.pack(fill="both", expand=True)
        ttk.Button(right, text="Mettre à jour ce preset",
                   command=self._apply_preset_edit).pack(anchor="e", pady=6)

        # working copy
        self.presets = json.loads(json.dumps(self.cfg["presets"]))
        self._reload_preset_list()
        if self.presets:
            self.preset_list.selection_set(0)
            self._on_select_preset()

    def _reload_preset_list(self):
        self.preset_list.delete(0, "end")
        for p in self.presets:
            self.preset_list.insert("end", p["name"])

    def _current_index(self):
        sel = self.preset_list.curselection()
        return sel[0] if sel else None

    def _on_select_preset(self, *_):
        i = self._current_index()
        if i is None:
            return
        self.name_var.set(self.presets[i]["name"])
        self.instr_text.delete("1.0", "end")
        self.instr_text.insert("1.0", self.presets[i]["instruction"])

    def _apply_preset_edit(self):
        i = self._current_index()
        if i is None:
            return
        self.presets[i]["name"] = self.name_var.get().strip() or "Sans nom"
        self.presets[i]["instruction"] = self.instr_text.get("1.0", "end-1c")
        self._reload_preset_list()
        self.preset_list.selection_set(i)

    def _add_preset(self):
        self.presets.append({"name": "Nouveau preset",
                             "instruction": "Décris ici l'instruction…"})
        self._reload_preset_list()
        i = len(self.presets) - 1
        self.preset_list.selection_clear(0, "end")
        self.preset_list.selection_set(i)
        self._on_select_preset()

    def _del_preset(self):
        i = self._current_index()
        if i is None or len(self.presets) <= 1:
            messagebox.showinfo(APP_NAME, "Il faut garder au moins un preset.",
                                parent=self.win)
            return
        del self.presets[i]
        self._reload_preset_list()

    # ---- Sauvegarde ---------------------------------------------------------
    def save(self):
        # On capture la dernière édition non validée du preset courant.
        self._apply_preset_edit()
        self.cfg["api_key"] = self.api_var.get().strip()
        self.cfg["default_model"] = self.model_var.get().strip()
        old_hotkey = self.cfg.get("hotkey")
        self.cfg["hotkey"] = self.hotkey_var.get().strip() or "alt+q"
        self.cfg["github_repo"] = self.repo_var.get().strip()
        self.cfg["check_updates_on_start"] = bool(self.updcheck_var.get())
        self.cfg["presets"] = self.presets
        save_config(self.cfg)

        try:
            set_startup(self.startup_var.get())
        except Exception as e:
            messagebox.showwarning(APP_NAME,
                                   f"Démarrage Windows non configuré : {e}",
                                   parent=self.win)

        if old_hotkey != self.cfg["hotkey"]:
            self.app.reload_hotkey()

        messagebox.showinfo(APP_NAME, "Paramètres enregistrés.", parent=self.win)
        self.win.destroy()


# =============================================================================
#  Démarrage avec Windows (clé de registre Run)
# =============================================================================
RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"


def _startup_command():
    if getattr(sys, "frozen", False):           # exécutable PyInstaller
        return f'"{sys.executable}"'
    # mode script : python.exe + chemin du script
    return f'"{sys.executable}" "{os.path.abspath(__file__)}"'


def is_startup_enabled():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as k:
            winreg.QueryValueEx(k, APP_NAME)
            return True
    except FileNotFoundError:
        return False
    except Exception:
        return False


def set_startup(enable):
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0,
                        winreg.KEY_SET_VALUE) as k:
        if enable:
            winreg.SetValueEx(k, APP_NAME, 0, winreg.REG_SZ, _startup_command())
        else:
            try:
                winreg.DeleteValue(k, APP_NAME)
            except FileNotFoundError:
                pass


# =============================================================================
#  Point d'entrée
# =============================================================================
def main():
    try:
        app = TextEnhancerApp()
        app.start()
    except Exception:
        # En cas de crash, on log dans le dossier de config pour diagnostic.
        with open(os.path.join(config_dir(), "crash.log"), "a",
                  encoding="utf-8") as f:
            f.write(traceback.format_exc() + "\n")
        raise


if __name__ == "__main__":
    main()
