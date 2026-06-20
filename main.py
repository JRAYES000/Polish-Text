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
import re
import sys
import json
import time
import threading
import traceback
import tempfile
import subprocess
import webbrowser
import base64
import shutil

# --- Dépendances tierces -----------------------------------------------------
import requests
import keyboard                       # raccourcis globaux + simulation de touches
import pystray
from PIL import Image, ImageDraw

import tkinter as tk
from tkinter import ttk, messagebox, font as tkfont

# --- Spécifique Windows ------------------------------------------------------
import win32clipboard
import win32con
import win32gui
import win32event
import win32api
import win32crypt
import winerror
import winreg


APP_NAME = "TextEnhancer AI"
APP_DIR_NAME = "TextEnhancerAI"
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
OPENROUTER_MODELS_URL = "https://openrouter.ai/api/v1/models"

# --- Version & mise à jour automatique ---------------------------------------
# Dépôt GitHub utilisé pour les mises à jour (modifiable aussi dans
# Paramètres → Dépôt GitHub, sans recompiler).
APP_VERSION = "1.3.2"
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
    # Modèle par défaut : économique et NON rate-limité (~0,10 $/M tokens en
    # sortie), largement suffisant pour de la reformulation.
    "default_model": "qwen/qwen3-235b-a22b-2507",
    "hotkey": "alt+q",
    "github_repo": GITHUB_REPO,
    "check_updates_on_start": True,
    # Le raccourci global historique sert maintenant de raccourci pour le 1er
    # preset (voir migration dans load_config). Chaque preset a son raccourci.
    "window_second_left": True,
    "theme": "light",  # "light" ou "dark"
    "known_models": [
        "qwen/qwen3-235b-a22b-2507",
        "qwen/qwen3-next-80b-a3b-instruct",
        "qwen/qwen3.7-plus",
        "anthropic/claude-sonnet-4.6",
        "google/gemini-2.5-flash",
    ],
    "presets": [
        {
            "name": "Reformuler (clair + gras)",
            "hotkey": "alt+q",
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


def _encrypt_secret(text):
    """Chiffre une chaîne via la DPAPI Windows (déchiffrable uniquement par la
    session Windows courante). Renvoie du base64, ou '' en cas d'échec/vide."""
    if not text:
        return ""
    try:
        blob = win32crypt.CryptProtectData(text.encode("utf-8"),
                                           "TextEnhancerAI", None, None, None, 0)
        return base64.b64encode(blob).decode("ascii")
    except Exception:
        return ""


def _decrypt_secret(b64):
    if not b64:
        return ""
    try:
        blob = base64.b64decode(b64.encode("ascii"))
        _, data = win32crypt.CryptUnprotectData(blob, None, None, None, 0)
        return data.decode("utf-8")
    except Exception:
        return ""


def load_config():
    path = config_path()
    if not os.path.exists(path):
        cfg = json.loads(json.dumps(DEFAULT_CONFIG))
        save_config(cfg)
        return cfg
    cfg = None
    try:
        with open(path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except Exception:
        # Fichier corrompu : on tente la sauvegarde .bak.
        bak = path + ".bak"
        if os.path.exists(bak):
            try:
                with open(bak, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
            except Exception:
                cfg = None
    if cfg is None:
        cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    # Complète les clés manquantes (migrations douces)
    for k, v in DEFAULT_CONFIG.items():
        cfg.setdefault(k, v)
    if not cfg.get("presets"):
        cfg["presets"] = json.loads(json.dumps(DEFAULT_CONFIG["presets"]))
    # Migration : l'ancien raccourci global unique devient le raccourci du
    # 1er preset (si aucun preset n'en a déjà un).
    presets = cfg.get("presets") or []
    if presets and not any((p.get("hotkey") or "").strip() for p in presets):
        legacy = (cfg.get("hotkey") or "").strip()
        if legacy:
            presets[0]["hotkey"] = legacy
    # Déchiffre la clé API si elle est stockée chiffrée.
    enc = cfg.get("api_key_enc")
    if enc:
        dec = _decrypt_secret(enc)
        if dec:
            cfg["api_key"] = dec
    return cfg


def save_config(cfg):
    """Écriture atomique (fichier temporaire puis os.replace) avec sauvegarde
    .bak. La clé API est chiffrée (DPAPI) et n'est jamais écrite en clair."""
    path = config_path()
    to_save = dict(cfg)
    enc = _encrypt_secret(cfg.get("api_key", ""))
    to_save["api_key_enc"] = enc
    # Si le chiffrement échoue, on conserve la clé pour ne pas la perdre.
    to_save["api_key"] = "" if enc else cfg.get("api_key", "")

    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(to_save, f, ensure_ascii=False, indent=2)
        f.flush()
        try:
            os.fsync(f.fileno())
        except Exception:
            pass
    if os.path.exists(path):
        try:
            shutil.copy2(path, path + ".bak")
        except Exception:
            pass
    os.replace(tmp, path)


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
    html_pre = ('<html><head><meta charset="utf-8"></head><body>\r\n'
                "<!--StartFragment-->")
    html_post = "<!--EndFragment-->\r\n</body></html>"

    header_len = len(header_tpl.format(0, 0, 0, 0).encode("utf-8"))
    start_html = header_len
    start_fragment = start_html + len(html_pre.encode("utf-8"))
    end_fragment = start_fragment + len(fragment_html.encode("utf-8"))
    end_html = end_fragment + len(html_post.encode("utf-8"))

    header = header_tpl.format(start_html, end_html, start_fragment, end_fragment)
    return header + html_pre + fragment_html + html_post


def set_clipboard_html(html_fragment, plain_text):
    """Place dans le presse-papiers une version HTML (texte enrichi : gras,
    sauts de ligne...) ET une version texte brut de secours."""
    cf_html_payload = _build_cf_html(html_fragment).encode("utf-8")
    for _ in range(10):
        try:
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, plain_text)
                cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
                win32clipboard.SetClipboardData(cf_html, cf_html_payload)
                return True
            finally:
                win32clipboard.CloseClipboard()
        except Exception:
            time.sleep(0.05)
    return False


# --- Rendu enrichi (gras + liens cliquables) dans un widget Text -------------
LINK_MD_RE = re.compile(r"\[([^\]]*)\]\((https?://[^)\s]+)\)")
URL_RE = re.compile(r"(https?://[^\s)]+)")
LINK_COLOR = "#2563eb"


def _normalize_bold(line):
    """Convertit le gras HTML (<b>, <strong>) et __..__ en **..** pour un
    découpage uniforme."""
    s = re.sub(r"</?(?:b|strong)\s*>", "**", line, flags=re.IGNORECASE)
    s = re.sub(r"__(.+?)__", r"**\1**", s)
    return s


def _parse_bold_runs(line):
    """Découpe une ligne en segments (texte, gras?) selon les marqueurs **."""
    runs = []
    bold = False
    for part in _normalize_bold(line).split("**"):
        if part:
            runs.append((part, bold))
        bold = not bold
    return runs


def _segments_plain(s):
    """Segments (texte, gras, url|None) en détectant les URLs nues."""
    out = []
    pos = 0
    for m in URL_RE.finditer(s):
        for t, b in _parse_bold_runs(s[pos:m.start()]):
            out.append((t, b, None))
        out.append((m.group(1), False, m.group(1)))  # URL nue, cliquable
        pos = m.end()
    for t, b in _parse_bold_runs(s[pos:]):
        out.append((t, b, None))
    return out


def parse_segments(line):
    """Découpe une ligne en segments (texte, gras?, url|None) : gère les liens
    Markdown [texte](url) ET les URLs nues."""
    out = []
    pos = 0
    for m in LINK_MD_RE.finditer(line):
        out.extend(_segments_plain(line[pos:m.start()]))
        text, url = (m.group(1) or m.group(2)), m.group(2)
        out.append((text, False, url))  # lien Markdown : texte cliquable
        pos = m.end()
    out.extend(_segments_plain(line[pos:]))
    return out


def _insert_segment(widget, text, bold, url):
    tags = []
    if bold:
        tags.append("bold")
    if url:
        tag = f"link-{widget._link_counter}"
        widget._link_counter += 1
        widget._link_urls[tag] = url
        widget.tag_configure(tag, foreground=LINK_COLOR, underline=1)
        widget.tag_bind(tag, "<Button-1>", lambda e, u=url: webbrowser.open(u))
        widget.tag_bind(tag, "<Enter>",
                        lambda e: widget.configure(cursor="hand2"))
        widget.tag_bind(tag, "<Leave>", lambda e: widget.configure(cursor=""))
        tags.append(tag)
    widget.insert("end", text, tuple(tags))


def render_markup_into(text_widget, markup):
    """Affiche le texte du LLM avec un vrai gras et des liens cliquables, sans
    laisser apparaître les marqueurs ** ni la syntaxe [texte](url)."""
    text_widget.configure(state="normal")
    text_widget.delete("1.0", "end")
    text_widget._link_urls = {}
    text_widget._link_counter = 0
    lines = (markup or "").replace("\r\n", "\n").split("\n")
    for li, line in enumerate(lines):
        heading = False
        m = re.match(r"^\s*#{1,6}\s*(.*)$", line)
        if m:
            line, heading = m.group(1), True
        mb = re.match(r"^\s*[-*+]\s+(.*)$", line)
        if mb:
            text_widget.insert("end", "•  ")
            line = mb.group(1)
        for seg, is_bold, url in parse_segments(line):
            # Un lien n'est jamais en gras (évite tout chevauchement de balises).
            _insert_segment(text_widget, seg, (is_bold or heading) and not url, url)
        if li < len(lines) - 1:
            text_widget.insert("end", "\n")


def _esc_html(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def text_widget_to_html(text_widget):
    """Reconstruit un fragment HTML (gras + liens <a> + sauts de ligne) à partir
    du contenu et des plages du widget — fonctionne même après édition."""
    urls = getattr(text_widget, "_link_urls", {})
    parts = []
    for key, value, _ in text_widget.dump("1.0", "end-1c", text=True, tag=True):
        if key == "text":
            parts.append(_esc_html(value).replace("\n", "<br>\n"))
        elif key == "tagon":
            if value == "bold":
                parts.append("<b>")
            elif value.startswith("link-"):
                href = _esc_html(urls.get(value, "")).replace('"', "%22")
                parts.append(f'<a href="{href}">')
        elif key == "tagoff":
            if value == "bold":
                parts.append("</b>")
            elif value.startswith("link-"):
                parts.append("</a>")
    return "".join(parts)


def text_widget_to_plain(text_widget):
    """Texte brut : conserve les URLs sous la forme « texte (url) » (ou juste
    l'url quand le texte du lien EST l'url)."""
    urls = getattr(text_widget, "_link_urls", {})
    out = []
    active_url = None
    link_text = []
    for key, value, _ in text_widget.dump("1.0", "end-1c", text=True, tag=True):
        if key == "text":
            (link_text if active_url is not None else out).append(value)
        elif key == "tagon" and value.startswith("link-"):
            active_url = urls.get(value, "")
            link_text = []
        elif key == "tagoff" and value.startswith("link-"):
            t = "".join(link_text).strip()
            out.append(f"{t} ({active_url})" if (t and t != active_url)
                       else (active_url or t))
            active_url = None
            link_text = []
    return "".join(out)


def style_editor(widget):
    """Fond blanc + texte foncé pour les zones d'édition principales
    (lisibilité garantie, même en thème sombre)."""
    try:
        widget.configure(background="#ffffff", foreground="#1f2430",
                         insertbackground="#1f2430", selectbackground="#cfe0ff",
                         highlightthickness=1, highlightbackground="#c9ccd1",
                         highlightcolor="#2563eb", relief="flat", borderwidth=0)
    except Exception:
        pass


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


def paste_into(target_hwnd, html=None, plain=""):
    """Redonne le focus à l'application cible puis colle (Ctrl+V). Si `html`
    est fourni, colle en texte enrichi ; sinon en texte brut."""
    if html is not None:
        set_clipboard_html(html, plain)
    else:
        set_clipboard_text(plain)
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
    last_err = "Échec de l'appel à OpenRouter."
    for attempt in range(3):
        try:
            resp = requests.post(OPENROUTER_URL, headers=headers,
                                 json=payload, timeout=timeout)
        except requests.exceptions.RequestException as e:
            last_err = f"Problème réseau : {e}"
            time.sleep(1.5 * (attempt + 1))
            continue

        if resp.status_code == 200:
            try:
                data = resp.json()
                return data["choices"][0]["message"]["content"].strip()
            except (KeyError, IndexError, ValueError):
                raise RuntimeError("Réponse inattendue d'OpenRouter : "
                                   f"{resp.text[:300]}")
        # Erreurs transitoires : on réessaie.
        if resp.status_code in (429, 500, 502, 503, 504):
            last_err = _friendly_openrouter_error(resp.status_code, resp.text)
            time.sleep(1.5 * (attempt + 1))
            continue
        # Erreur non récupérable : message clair immédiat.
        raise RuntimeError(_friendly_openrouter_error(resp.status_code, resp.text))

    raise RuntimeError(last_err)


def _friendly_openrouter_error(status, text):
    if status == 401:
        return ("Clé API invalide ou expirée (401). Vérifie ta clé dans "
                "Paramètres → Clé API.")
    if status == 402:
        return ("Crédit OpenRouter insuffisant (402). Ajoute du crédit sur "
                "ton compte OpenRouter.")
    if status == 429:
        return ("Quota atteint / trop de requêtes (429). Réessaie dans un "
                "instant, ou utilise un modèle non gratuit.")
    if status == 404:
        return ("Modèle introuvable (404). Vérifie l'identifiant du modèle "
                "dans les Paramètres.")
    return f"Erreur OpenRouter {status} : {text[:300]}"


def stream_openrouter(api_key, model, instruction, user_text, timeout=120):
    """Générateur : émet le texte de la réponse au fur et à mesure (streaming
    SSE d'OpenRouter). Lève RuntimeError avec un message clair en cas d'erreur."""
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
        "stream": True,
        "messages": [
            {"role": "system", "content": instruction},
            {"role": "user", "content": user_text},
        ],
    }
    with requests.post(OPENROUTER_URL, headers=headers, json=payload,
                       stream=True, timeout=timeout) as resp:
        if resp.status_code != 200:
            resp.encoding = "utf-8"
            raise RuntimeError(
                _friendly_openrouter_error(resp.status_code, resp.text))
        # On lit les octets bruts et on décode explicitement en UTF-8 : sans
        # cela, requests décode en Latin-1 (charset absent) -> accents corrompus.
        for raw_b in resp.iter_lines():
            if not raw_b:
                continue
            raw = raw_b.decode("utf-8", errors="replace")
            if not raw.startswith("data:"):
                continue
            data = raw[5:].strip()
            if data == "[DONE]":
                break
            try:
                obj = json.loads(data)
                delta = obj["choices"][0]["delta"].get("content")
            except (KeyError, IndexError, ValueError):
                continue
            if delta:
                yield delta


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
    via un script batch. Le script réessaie le remplacement en boucle jusqu'à ce
    que l'ancien exe soit déverrouillé (donc dès que l'app est fermée), puis
    relance l'application. Conçu pour fonctionner avec une console cachée."""
    if not getattr(sys, "frozen", False):
        raise RuntimeError("La mise à jour automatique n'est disponible que "
                           "pour l'exécutable (.exe).")
    current_exe = sys.executable
    exe_name = os.path.basename(current_exe)
    new_exe = os.path.join(tempfile.gettempdir(), "TextEnhancerAI_new.exe")

    with requests.get(asset_url, stream=True, timeout=180) as r:
        r.raise_for_status()
        with open(new_exe, "wb") as f:
            for chunk in r.iter_content(chunk_size=65536):
                if chunk:
                    f.write(chunk)

    # Garde-fou : un exe valide pèse plusieurs Mo.
    if os.path.getsize(new_exe) < 1_000_000:
        raise RuntimeError("Téléchargement de la mise à jour incomplet.")

    log = os.path.join(tempfile.gettempdir(), "TextEnhancerAI_update.log")
    bat_path = os.path.join(tempfile.gettempdir(), "TextEnhancerAI_update.bat")
    # Remarques :
    #  - 'ping' sert de temporisation (contrairement à 'timeout', il ne dépend
    #    pas d'une console interactive).
    #  - On réessaie 'move' jusqu'à ce qu'il réussisse : il échoue tant que
    #    l'ancien exe est verrouillé (app encore ouverte), puis réussit.
    script = (
        "@echo off\r\n"
        f'echo [%DATE% %TIME%] debut mise a jour> "{log}"\r\n'
        "set /a TRIES=0\r\n"
        ":retry\r\n"
        "ping 127.0.0.1 -n 2 >nul\r\n"
        f'move /y "{new_exe}" "{current_exe}" >> "{log}" 2>&1\r\n'
        f'if not exist "{new_exe}" goto done\r\n'
        "set /a TRIES+=1\r\n"
        "if %TRIES% LSS 60 goto retry\r\n"
        f'echo [%DATE% %TIME%] echec apres 60 tentatives>> "{log}"\r\n'
        "exit\r\n"
        ":done\r\n"
        f'echo [%DATE% %TIME%] move OK, pause avant relance>> "{log}"\r\n'
        # Pause : laisse l'antivirus finir de scanner l'exe fraîchement déplacé
        # (sinon l'extraction du DLL Python au lancement peut échouer).
        "ping 127.0.0.1 -n 4 >nul\r\n"
        f'start "" "{current_exe}"\r\n'
        # Vérifie que l'app a bien démarré ; sinon, relance de secours.
        "ping 127.0.0.1 -n 6 >nul\r\n"
        f'tasklist /FI "IMAGENAME eq {exe_name}" /NH 2>nul | find /I "{exe_name}" >nul\r\n'
        "if not errorlevel 1 goto launched\r\n"
        f'echo [%DATE% %TIME%] 1er demarrage KO, relance de secours>> "{log}"\r\n'
        f'start "" "{current_exe}"\r\n'
        ":launched\r\n"
        f'echo [%DATE% %TIME%] termine>> "{log}"\r\n'
        'del "%~f0"\r\n'
    )
    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(script)

    # CREATE_NO_WINDOW : le script tourne avec une console cachée (donc 'ping',
    # 'move' & co fonctionnent) et survit à la fermeture de l'application.
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
    subprocess.Popen(["cmd", "/c", bat_path], creationflags=creationflags,
                     close_fds=True)


# =============================================================================
#  Thème (clair / sombre)
# =============================================================================
PALETTES = {
    "light": {
        "bg": "#f4f5f7", "surface": "#ffffff", "fg": "#1f2430",
        "muted": "#6b7280", "field_bg": "#ffffff", "field_fg": "#1f2430",
        "accent": "#2563eb", "accent_fg": "#ffffff",
        "select_bg": "#dbeafe", "border": "#d6d9de",
    },
    "dark": {
        "bg": "#1e2128", "surface": "#262a33", "fg": "#e6e8ec",
        "muted": "#9aa0ab", "field_bg": "#2f343d", "field_fg": "#e6e8ec",
        "accent": "#4f8cff", "accent_fg": "#0b0d10",
        "select_bg": "#34405a", "border": "#3a3f4a",
    },
}


def get_palette(cfg):
    return PALETTES.get(cfg.get("theme", "light"), PALETTES["light"])


def apply_theme(root, style, pal):
    """Applique la palette aux styles ttk (thème 'clam' recolorable) et aux
    listes déroulantes des Combobox."""
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure(".", background=pal["bg"], foreground=pal["fg"],
                    fieldbackground=pal["field_bg"], bordercolor=pal["border"])
    style.configure("TFrame", background=pal["bg"])
    style.configure("TLabelframe", background=pal["bg"], bordercolor=pal["border"])
    style.configure("TLabelframe.Label", background=pal["bg"],
                    foreground=pal["fg"])
    style.configure("TLabel", background=pal["bg"], foreground=pal["fg"])
    style.configure("Muted.TLabel", background=pal["bg"], foreground=pal["muted"])
    style.configure("Status.TLabel", background=pal["bg"], foreground=pal["accent"])
    style.configure("Toast.TLabel", background=pal["accent"],
                    foreground=pal["accent_fg"], padding=6)
    style.configure("TButton", background=pal["surface"], foreground=pal["fg"],
                    bordercolor=pal["border"], padding=6)
    style.map("TButton", background=[("active", pal["select_bg"]),
                                     ("pressed", pal["select_bg"])])
    style.configure("Accent.TButton", background=pal["accent"],
                    foreground=pal["accent_fg"])
    style.map("Accent.TButton", background=[("active", pal["accent"]),
                                            ("pressed", pal["accent"])])
    style.configure("TCheckbutton", background=pal["bg"], foreground=pal["fg"])
    style.map("TCheckbutton", background=[("active", pal["bg"])])
    style.configure("TNotebook", background=pal["bg"], bordercolor=pal["border"])
    style.configure("TNotebook.Tab", background=pal["surface"],
                    foreground=pal["fg"], padding=(12, 6))
    style.map("TNotebook.Tab", background=[("selected", pal["bg"])],
              foreground=[("selected", pal["accent"])])
    style.configure("TCombobox", fieldbackground=pal["field_bg"],
                    background=pal["surface"], foreground=pal["field_fg"],
                    bordercolor=pal["border"], arrowcolor=pal["fg"])
    style.map("TCombobox", fieldbackground=[("readonly", pal["field_bg"])],
              foreground=[("readonly", pal["field_fg"])])
    style.configure("TEntry", fieldbackground=pal["field_bg"],
                    foreground=pal["field_fg"], bordercolor=pal["border"],
                    insertcolor=pal["fg"])
    style.configure("Vertical.TScrollbar", background=pal["surface"],
                    troughcolor=pal["bg"], bordercolor=pal["border"],
                    arrowcolor=pal["fg"])
    style.configure("TPanedwindow", background=pal["bg"])
    for opt, key in (("background", "field_bg"), ("foreground", "field_fg"),
                     ("selectBackground", "select_bg"),
                     ("selectForeground", "fg")):
        try:
            root.option_add(f"*TCombobox*Listbox.{opt}", pal[key])
        except Exception:
            pass


def style_text_widget(widget, pal):
    """Colore un tk.Text ou tk.Listbox selon la palette."""
    try:
        widget.configure(background=pal["field_bg"], foreground=pal["field_fg"],
                         insertbackground=pal["fg"],
                         selectbackground=pal["select_bg"],
                         highlightthickness=1, highlightbackground=pal["border"],
                         highlightcolor=pal["accent"], relief="flat", borderwidth=0)
    except Exception:
        pass


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
        self.hotkey_ok = True
        self.hotkey_errors = []
        self.registered_count = 0
        self.history = []          # derniers résultats (en mémoire, session)
        self.history_max = 15

        # Root Tk caché : toutes les fenêtres en sont des Toplevel.
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.title(APP_NAME)

        self.style = ttk.Style(self.root)
        self.palette = get_palette(self.config)
        apply_theme(self.root, self.style, self.palette)

        self._register_hotkeys()

    def reapply_theme(self):
        self.palette = get_palette(self.config)
        apply_theme(self.root, self.style, self.palette)

    def add_history(self, entry):
        """entry = dict(preset, model, source, result)."""
        self.history.insert(0, entry)
        del self.history[self.history_max:]

    # ---- Raccourcis par preset ----------------------------------------------
    def _register_hotkeys(self):
        """Enregistre un raccourci global par preset qui en possède un."""
        try:
            keyboard.clear_all_hotkeys()
        except Exception:
            pass
        self.hotkey_errors = []
        used = {}
        registered = 0
        for p in self.config.get("presets", []):
            hk = (p.get("hotkey") or "").strip().lower()
            if not hk:
                continue
            if hk in used:
                self.hotkey_errors.append(
                    f"« {hk} » est en double ({used[hk]} et {p['name']})")
                continue
            try:
                keyboard.add_hotkey(
                    hk, lambda name=p["name"]: self._on_preset_hotkey(name))
                used[hk] = p["name"]
                registered += 1
            except Exception as e:
                self.hotkey_errors.append(f"{p['name']} ({hk}) : {e}")
        self.registered_count = registered
        self.hotkey_ok = (len(self.hotkey_errors) == 0)
        return self.hotkey_ok

    def _on_preset_hotkey(self, preset_name):
        # Capture immédiate (l'app cible est encore au premier plan),
        # puis on ouvre l'aperçu sur le thread Tk, présélectionné sur ce preset.
        self.target_hwnd = win32gui.GetForegroundWindow()
        selection = capture_selection()
        self.root.after(0, lambda: self._open_preview(selection, preset_name))

    # ---- Fenêtre d'aperçu ---------------------------------------------------
    def _open_preview(self, selection, preset_name=None):
        if self.preview_win and self.preview_win.win.winfo_exists():
            self.preview_win.win.lift()
            self.preview_win.set_source(selection, preset_name)
            return
        self.preview_win = PreviewWindow(self, selection, preset_name)

    # ---- Fenêtre paramètres -------------------------------------------------
    def open_settings(self):
        if self.settings_win and self.settings_win.win.winfo_exists():
            self.settings_win.win.lift()
            return
        self.settings_win = SettingsWindow(self)

    def manual_trigger(self):
        """Déclenchement depuis le menu tray : utilise le presse-papiers actuel."""
        self.target_hwnd = None
        self.root.after(0, lambda: self._open_preview(get_clipboard_text(), None))

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

    def reload_hotkeys(self):
        ok = self._register_hotkeys()
        if not ok:
            messagebox.showwarning(
                APP_NAME,
                "Certains raccourcis n'ont pas pu être enregistrés :\n\n- "
                + "\n- ".join(self.hotkey_errors)
                + "\n\nVérifie qu'ils sont valides (ex. alt+q, ctrl+alt+r) "
                "et non dupliqués.")
        return ok

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
        if not self.hotkey_ok:
            self.root.after(900, lambda: messagebox.showwarning(
                APP_NAME,
                "Certains raccourcis de presets n'ont pas pu être enregistrés :"
                "\n\n- " + "\n- ".join(self.hotkey_errors)
                + "\n\nModifie-les dans Paramètres → Presets. Tu peux aussi "
                "utiliser « Améliorer (presse-papiers) » via l'icône de la "
                "barre des tâches."))
        # Vérification des mises à jour au démarrage (en arrière-plan).
        if self.config.get("check_updates_on_start", True):
            self.root.after(2500, lambda: self.check_updates(silent=True))
        self.root.mainloop()


def compute_window_geometry(width, height, prefer_second_left=True):
    """Renvoie une chaîne Tk 'WxH+X+Y'. Si prefer_second_left et qu'un 2e écran
    existe, place la fenêtre sur la MOITIÉ GAUCHE de ce 2e écran ; sinon centre
    la fenêtre (taille demandée) sur l'écran principal."""
    mons = []
    try:
        for (hmon, _hdc, _rect) in win32api.EnumDisplayMonitors():
            mons.append(win32api.GetMonitorInfo(hmon))
    except Exception:
        mons = []
    if not mons:
        return f"{width}x{height}"

    primary = None
    for mi in mons:
        if mi.get("Flags", 0) & 1:  # MONITORINFOF_PRIMARY
            primary = mi
    second = None
    for mi in mons:
        if mi is not primary:
            second = mi
            break

    if prefer_second_left and second is not None:
        l, t, r, b = second.get("Work") or second.get("Monitor")
        w = max(1, (r - l) // 2)
        # On garde une marge en haut et en bas : la barre de titre s'ajoute à la
        # hauteur, et il faut que le bas (les boutons) ne passe pas sous la
        # barre des tâches.
        margin_top = 30
        h = max(300, (b - t) - 100)
        return f"{w}x{h}+{l}+{t + margin_top}"

    target = primary or mons[0]
    l, t, r, b = target.get("Work") or target.get("Monitor")
    x = l + max(0, ((r - l) - width) // 2)
    y = t + max(0, ((b - t) - height) // 2)
    return f"{width}x{height}+{x}+{y}"


# =============================================================================
#  Fenêtre d'aperçu
# =============================================================================
class PreviewWindow:
    def __init__(self, app, source_text, preset_name=None):
        self.app = app
        self.cfg = app.config
        self.pal = app.palette
        self.source_text = source_text or ""
        self.preset_items = self.cfg["presets"]
        self.prompt_override = None  # instruction modifiée pour la session

        self.win = tk.Toplevel(app.root)
        self.win.title(f"{APP_NAME} — Aperçu")
        self.win.configure(bg=self.pal["bg"])
        self.win.minsize(560, 480)
        try:
            self.win.geometry(compute_window_geometry(
                780, 640, self.cfg.get("window_second_left", True)))
        except Exception:
            self.win.geometry("780x640")
        # Apparaît au premier plan, puis redevient une fenêtre normale : elle
        # passe en arrière-plan dès qu'on clique ailleurs.
        self.win.lift()
        self.win.focus_force()
        self.win.attributes("-topmost", True)
        self.win.after(400, self._release_topmost)

        labels = [self._preset_label(p) for p in self.preset_items]
        init_idx = self._index_of_name(preset_name)

        # --- Barre du haut : preset + modèle + récents (toujours visible) ---
        top = ttk.Frame(self.win, padding=(10, 8))
        top.pack(side="top", fill="x")
        ttk.Label(top, text="Style :").grid(row=0, column=0, sticky="w")
        self.preset_var = tk.StringVar(value=labels[init_idx] if labels else "")
        self.preset_cb = ttk.Combobox(top, textvariable=self.preset_var,
                                      values=labels, state="readonly", width=30)
        self.preset_cb.grid(row=0, column=1, sticky="w", padx=(4, 16))
        if labels:
            self.preset_cb.current(init_idx)
        self.preset_cb.bind("<<ComboboxSelected>>", self._on_preset_change)

        ttk.Label(top, text="Modèle :").grid(row=0, column=2, sticky="w")
        self.model_var = tk.StringVar()
        self.model_cb = ttk.Combobox(top, textvariable=self.model_var,
                                     values=self.cfg.get("known_models", []),
                                     width=22)
        self.model_cb.grid(row=0, column=3, sticky="w", padx=(4, 12))
        top.columnconfigure(4, weight=1)  # espaceur
        ttk.Button(top, text="Prompt…", command=self._open_prompt_editor).grid(
            row=0, column=5, sticky="e", padx=(0, 6))
        ttk.Button(top, text="Récents ▾", command=self._open_recents).grid(
            row=0, column=6, sticky="e")

        # --- Barre de boutons : packée AVANT le centre -> TOUJOURS visible ---
        btns = ttk.Frame(self.win, padding=(10, 8))
        btns.pack(side="bottom", fill="x")
        self.gen_btn = ttk.Button(btns, text="Générer / Regénérer",
                                  style="Accent.TButton", command=self.generate)
        self.gen_btn.pack(side="left")
        ttk.Button(btns, text="Coller (texte enrichi)",
                   command=lambda: self.paste(rich=True)).pack(side="left", padx=6)
        ttk.Button(btns, text="Coller (brut)",
                   command=lambda: self.paste(rich=False)).pack(side="left")
        ttk.Button(btns, text="Copier",
                   command=self.copy).pack(side="left", padx=6)
        ttk.Button(btns, text="Fermer",
                   command=self.win.destroy).pack(side="right")

        self.status_var = tk.StringVar(value="")
        ttk.Label(self.win, textvariable=self.status_var, style="Status.TLabel",
                  padding=(12, 0)).pack(side="bottom", fill="x")

        # --- Zone centrale redimensionnable : on tire le séparateur (sash) ---
        paned = ttk.PanedWindow(self.win, orient="vertical")
        paned.pack(side="top", fill="both", expand=True, padx=10, pady=(0, 4))

        src_frame = ttk.LabelFrame(paned, text="Texte sélectionné", padding=6)
        res_frame = ttk.LabelFrame(
            paned, text="Résultat (éditable — mise en forme réelle)", padding=6)
        paned.add(src_frame, weight=1)
        paned.add(res_frame, weight=1)
        self._paned = paned

        swrap = ttk.Frame(src_frame)
        swrap.pack(fill="both", expand=True)
        self.src_text = tk.Text(swrap, height=4, wrap="word")
        sscroll = ttk.Scrollbar(swrap, orient="vertical",
                                command=self.src_text.yview)
        self.src_text.configure(yscrollcommand=sscroll.set)
        sscroll.pack(side="right", fill="y")
        self.src_text.pack(side="left", fill="both", expand=True)
        self.src_text.insert("1.0", self.source_text)
        self.src_text.configure(state="disabled")
        style_editor(self.src_text)

        rwrap = ttk.Frame(res_frame)
        rwrap.pack(fill="both", expand=True)
        self.result_text = tk.Text(rwrap, wrap="word", undo=True)
        rscroll = ttk.Scrollbar(rwrap, orient="vertical",
                                command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=rscroll.set)
        rscroll.pack(side="right", fill="y")
        self.result_text.pack(side="left", fill="both", expand=True)
        style_editor(self.result_text)

        # Tag « gras » pour l'affichage en mise en forme réelle.
        try:
            base = tkfont.nametofont(self.result_text.cget("font"))
        except Exception:
            base = tkfont.nametofont("TkTextFont")
        bold_font = base.copy()
        bold_font.configure(weight="bold")
        self.result_text.tag_configure("bold", font=bold_font)

        self._refresh_model_for_preset()
        self.win.bind("<Escape>", lambda e: self.win.destroy())
        # Hauteurs égales par défaut : on place le séparateur au milieu une fois
        # la fenêtre dimensionnée.
        self.win.after(80, self._equalize_panes)

        if self.source_text.strip():
            self.win.after(150, self.generate)
        else:
            self.status_var.set("Aucun texte capturé. Sélectionne du texte "
                                "(Ctrl+C) puis le raccourci du preset.")

    def _equalize_panes(self):
        try:
            self._paned.update_idletasks()
            h = self._paned.winfo_height()
            if h > 1:
                self._paned.sashpos(0, h // 2)
        except Exception:
            pass

    # ---- Helpers presets ----------------------------------------------------
    @staticmethod
    def _preset_label(p):
        hk = (p.get("hotkey") or "").strip()
        return f"{p['name']}  —  {hk}" if hk else p["name"]

    def _index_of_name(self, name):
        if name:
            for i, p in enumerate(self.preset_items):
                if p["name"] == name:
                    return i
        return 0

    def _current_preset(self):
        idx = self.preset_cb.current()
        if 0 <= idx < len(self.preset_items):
            return self.preset_items[idx]
        return self.preset_items[0] if self.preset_items else {
            "name": "", "instruction": "", "model": ""}

    def _refresh_model_for_preset(self):
        """Met le modèle de l'aperçu sur celui du preset (sinon défaut)."""
        p = self._current_preset()
        model = (p.get("model") or "").strip() or self.cfg.get("default_model", "")
        self.model_var.set(model)

    def _on_preset_change(self, *_):
        self.prompt_override = None  # le prompt modifié était lié au preset
        self._refresh_model_for_preset()

    def _select_preset(self, name):
        idx = self._index_of_name(name)
        labels = [self._preset_label(p) for p in self.preset_items]
        if labels:
            self.preset_cb.configure(values=labels)
            self.preset_cb.current(idx)
            self.prompt_override = None
            self._refresh_model_for_preset()

    def _release_topmost(self):
        try:
            if self.win.winfo_exists():
                self.win.attributes("-topmost", False)
        except Exception:
            pass

    # ---- Édition du prompt --------------------------------------------------
    def _open_prompt_editor(self):
        preset = self._current_preset()
        idx = self.preset_cb.current()
        ed = tk.Toplevel(self.win)
        ed.title("Prompt du style")
        ed.configure(bg=self.pal["bg"])
        ed.geometry("580x440")
        ed.transient(self.win)
        ed.lift()
        ed.focus_force()
        ttk.Label(ed, text=f"Instruction (prompt système) pour « "
                  f"{preset.get('name', '')} » :", padding=(10, 8)).pack(anchor="w")
        wrap = ttk.Frame(ed, padding=(10, 0))
        wrap.pack(fill="both", expand=True)
        txt = tk.Text(wrap, wrap="word")
        sc = ttk.Scrollbar(wrap, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=sc.set)
        sc.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)
        style_text_widget(txt, self.pal)
        txt.insert("1.0", self.prompt_override or preset.get("instruction", ""))

        info = ttk.Label(ed, style="Muted.TLabel", padding=(10, 4),
                         text="« Appliquer » : pour cette session. "
                              "« Enregistrer » : modifie le preset durablement.")
        info.pack(anchor="w")
        bar = ttk.Frame(ed, padding=10)
        bar.pack(fill="x")

        def apply_session():
            self.prompt_override = txt.get("1.0", "end-1c")
            ed.destroy()
            self.status_var.set("Prompt modifié (session). Régénération…")
            self.generate()

        def save_to_preset():
            new_instr = txt.get("1.0", "end-1c")
            if 0 <= idx < len(self.preset_items):
                self.preset_items[idx]["instruction"] = new_instr
                self.cfg["presets"] = self.preset_items
                try:
                    save_config(self.cfg)
                except Exception:
                    pass
            self.prompt_override = None
            ed.destroy()
            self.status_var.set("Prompt enregistré dans le preset. Régénération…")
            self.generate()

        ttk.Button(bar, text="Appliquer (session)", style="Accent.TButton",
                   command=apply_session).pack(side="left")
        ttk.Button(bar, text="Enregistrer dans le preset",
                   command=save_to_preset).pack(side="left", padx=6)
        ttk.Button(bar, text="Annuler", command=ed.destroy).pack(side="right")

    # ---- Récents ------------------------------------------------------------
    def _open_recents(self):
        menu = tk.Menu(self.win, tearoff=0)
        if not self.app.history:
            menu.add_command(label="(aucun résultat récent)", state="disabled")
        else:
            for h in self.app.history:
                src = " ".join((h.get("source") or "").split())
                label = f"{h.get('preset', '?')} · {src[:45]}…"
                menu.add_command(label=label,
                                 command=lambda e=h: self._load_history(e))
        try:
            menu.tk_popup(self.win.winfo_pointerx(), self.win.winfo_pointery())
        finally:
            menu.grab_release()

    def _load_history(self, entry):
        self._select_preset(entry.get("preset"))
        self.source_text = entry.get("source", "")
        self.src_text.configure(state="normal")
        self.src_text.delete("1.0", "end")
        self.src_text.insert("1.0", self.source_text)
        self.src_text.configure(state="disabled")
        render_markup_into(self.result_text, entry.get("result", ""))
        self.status_var.set("Résultat récent chargé.")

    # ---- Source / génération ------------------------------------------------
    def set_source(self, text, preset_name=None):
        if preset_name:
            self._select_preset(preset_name)
        self.source_text = text or ""
        self.src_text.configure(state="normal")
        self.src_text.delete("1.0", "end")
        self.src_text.insert("1.0", self.source_text)
        self.src_text.configure(state="disabled")
        if self.source_text.strip():
            self.generate()

    def generate(self):
        if not self.source_text.strip():
            self.status_var.set("Aucun texte à traiter.")
            return
        preset = self._current_preset()
        instruction = self.prompt_override or preset.get("instruction", "")
        preset_name = preset.get("name", "")
        model = self.model_var.get().strip() or self.cfg.get("default_model")
        api_key = self.cfg.get("api_key", "")
        src = self.source_text

        self.gen_btn.configure(state="disabled")
        self.status_var.set("Génération en cours…")
        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", "end")

        def worker():
            try:
                acc = []
                for chunk in stream_openrouter(api_key, model, instruction, src):
                    acc.append(chunk)
                    piece = chunk
                    self._safe_after(lambda c=piece: self._stream_append(c))
                full = "".join(acc).strip()
                self._safe_after(lambda: self._stream_done(
                    full, preset_name, model, src))
            except Exception as e:
                msg = str(e)
                self._safe_after(lambda: self._show_error(msg))

        threading.Thread(target=worker, daemon=True).start()

    def _safe_after(self, fn):
        try:
            self.win.after(0, fn)
        except Exception:
            pass

    def _stream_append(self, chunk):
        self.result_text.insert("end", chunk)
        self.result_text.see("end")

    def _stream_done(self, full, preset_name, model, src):
        render_markup_into(self.result_text, full)
        self.status_var.set("Prêt. Le gras est réel : « Coller (texte enrichi) » "
                            "pour Word/Outlook.")
        self.gen_btn.configure(state="normal")
        if full:
            self.app.add_history({"preset": preset_name, "model": model,
                                  "source": src, "result": full})

    def _show_error(self, msg):
        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", "end")
        self.status_var.set("Erreur : " + msg)
        self.gen_btn.configure(state="normal")

    def _result_plain(self):
        return text_widget_to_plain(self.result_text)

    def _result_html(self):
        return text_widget_to_html(self.result_text)

    # ---- Toast --------------------------------------------------------------
    def _toast(self, msg):
        try:
            t = ttk.Label(self.win, text=msg, style="Toast.TLabel")
            t.place(relx=0.5, rely=1.0, y=-48, anchor="s")
            self.win.after(1500, t.destroy)
        except Exception:
            pass

    def copy(self):
        if not self._result_plain().strip():
            return
        set_clipboard_html(self._result_html(), self._result_plain())
        self.status_var.set("Copié (texte enrichi).")
        self._toast("Copié ✓")

    def paste(self, rich=True):
        if not self._result_plain().strip():
            return
        hwnd = self.app.target_hwnd
        plain = self._result_plain()
        html = self._result_html() if rich else None
        self.win.destroy()
        # Petit délai pour laisser la fenêtre se fermer avant de coller.
        time.sleep(0.2)
        paste_into(hwnd, html=html, plain=plain)


# =============================================================================
#  Fenêtre paramètres
# =============================================================================
class SettingsWindow:
    def __init__(self, app):
        self.app = app
        self.cfg = app.config

        self.win = tk.Toplevel(app.root)
        self.win.title(f"{APP_NAME} — Paramètres")
        self.win.configure(bg=app.palette["bg"])
        self.win.geometry("760x660")
        self.win.lift()
        self.win.focus_force()
        self.win.attributes("-topmost", True)
        self.win.after(400, self._release_topmost)

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

    def _release_topmost(self):
        try:
            if self.win.winfo_exists():
                self.win.attributes("-topmost", False)
        except Exception:
            pass

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

        ttk.Label(f, text="Les raccourcis clavier se règlent par preset "
                  "(onglet « Presets »).", style="Muted.TLabel").grid(
            row=2, column=1, sticky="w", pady=(6, 2))

        self.dark_var = tk.BooleanVar(
            value=self.cfg.get("theme", "light") == "dark")
        ttk.Checkbutton(f, text="Mode sombre",
                        variable=self.dark_var).grid(
            row=3, column=1, sticky="w", pady=(4, 2))

        self.second_left_var = tk.BooleanVar(
            value=self.cfg.get("window_second_left", True))
        ttk.Checkbutton(
            f, text="Afficher la fenêtre sur le 2e écran (moitié gauche)",
            variable=self.second_left_var).grid(
            row=4, column=1, sticky="w", pady=(4, 2))

        self.startup_var = tk.BooleanVar(value=is_startup_enabled())
        ttk.Checkbutton(f, text="Lancer au démarrage de Windows",
                        variable=self.startup_var).grid(
            row=5, column=1, sticky="w", pady=(4, 2))

        ttk.Label(f, text="Dépôt GitHub :").grid(row=6, column=0, sticky="w", pady=4)
        self.repo_var = tk.StringVar(value=self.cfg.get("github_repo", ""))
        ttk.Entry(f, textvariable=self.repo_var, width=40).grid(
            row=6, column=1, sticky="w", pady=4)
        ttk.Label(f, text="(ex : ton-pseudo/text-enhancer-ai — pour les mises à jour)",
                  style="Muted.TLabel").grid(row=7, column=1, sticky="w")

        self.updcheck_var = tk.BooleanVar(
            value=self.cfg.get("check_updates_on_start", True))
        ttk.Checkbutton(f, text="Vérifier les mises à jour au démarrage",
                        variable=self.updcheck_var).grid(
            row=8, column=1, sticky="w", pady=(8, 2))

        ttk.Label(f, text=f"Version installée : {APP_VERSION}",
                  style="Muted.TLabel").grid(row=9, column=1, sticky="w",
                                             pady=(8, 0))

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

        pal = self.app.palette
        left = ttk.Frame(f)
        left.pack(side="left", fill="y", padx=(0, 10))
        self.preset_list = tk.Listbox(left, width=26, height=18,
                                      exportselection=False)
        self.preset_list.pack(fill="y", expand=True)
        try:
            self.preset_list.configure(
                background=pal["field_bg"], foreground=pal["field_fg"],
                selectbackground=pal["select_bg"], selectforeground=pal["fg"],
                highlightthickness=1, highlightbackground=pal["border"],
                relief="flat", borderwidth=0)
        except Exception:
            pass
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

        ttk.Label(right, text="Raccourci clavier (vide = aucun) :").pack(
            anchor="w", pady=(8, 0))
        hkrow = ttk.Frame(right)
        hkrow.pack(anchor="w", fill="x")
        self.hotkey_var = tk.StringVar()
        ttk.Entry(hkrow, textvariable=self.hotkey_var, width=20).pack(side="left")
        self.capture_btn = ttk.Button(hkrow, text="Définir…",
                                      command=self._capture_hotkey)
        self.capture_btn.pack(side="left", padx=6)

        ttk.Label(right, text="Modèle (vide = modèle par défaut) :").pack(
            anchor="w", pady=(8, 0))
        self.preset_model_var = tk.StringVar()
        ttk.Combobox(right, textvariable=self.preset_model_var,
                     values=self.cfg.get("known_models", []), width=40).pack(
            anchor="w")

        ttk.Label(right, text="Instruction (prompt système) :").pack(
            anchor="w", pady=(8, 0))
        self.instr_text = tk.Text(right, wrap="word", height=13)
        self.instr_text.pack(fill="both", expand=True)
        style_text_widget(self.instr_text, pal)
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
        self.hotkey_var.set(self.presets[i].get("hotkey", ""))
        self.preset_model_var.set(self.presets[i].get("model", ""))
        self.instr_text.delete("1.0", "end")
        self.instr_text.insert("1.0", self.presets[i]["instruction"])

    def _apply_preset_edit(self):
        i = self._current_index()
        if i is None:
            return
        self.presets[i]["name"] = self.name_var.get().strip() or "Sans nom"
        self.presets[i]["hotkey"] = self.hotkey_var.get().strip().lower()
        self.presets[i]["model"] = self.preset_model_var.get().strip()
        self.presets[i]["instruction"] = self.instr_text.get("1.0", "end-1c")
        self._reload_preset_list()
        self.preset_list.selection_set(i)

    def _capture_hotkey(self):
        """Capture la prochaine combinaison de touches pressée par l'utilisateur."""
        self.capture_btn.configure(state="disabled", text="Pressez la combinaison…")

        def worker():
            # On désactive temporairement les raccourcis de l'app pour ne pas
            # déclencher un preset pendant la capture.
            try:
                keyboard.clear_all_hotkeys()
            except Exception:
                pass
            try:
                hk = keyboard.read_hotkey(suppress=False)
            except Exception:
                hk = ""

            def done():
                if hk and hk.lower() != "esc":
                    self.hotkey_var.set(hk.lower())
                self.capture_btn.configure(state="normal", text="Définir…")
                self.app.reload_hotkeys()  # restaure les raccourcis enregistrés

            try:
                self.win.after(0, done)
            except Exception:
                pass

        threading.Thread(target=worker, daemon=True).start()

    def _add_preset(self):
        self.presets.append({"name": "Nouveau preset", "hotkey": "", "model": "",
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
        self.cfg["github_repo"] = self.repo_var.get().strip()
        self.cfg["check_updates_on_start"] = bool(self.updcheck_var.get())
        self.cfg["window_second_left"] = bool(self.second_left_var.get())
        self.cfg["theme"] = "dark" if self.dark_var.get() else "light"
        self.cfg["presets"] = self.presets
        save_config(self.cfg)

        try:
            set_startup(self.startup_var.get())
        except Exception as e:
            messagebox.showwarning(APP_NAME,
                                   f"Démarrage Windows non configuré : {e}",
                                   parent=self.win)

        # Thème + raccourcis : on réapplique toujours.
        self.win.destroy()
        self.app.reapply_theme()
        self.app.reload_hotkeys()
        messagebox.showinfo(APP_NAME,
                            "Paramètres enregistrés. Le thème s'applique aux "
                            "prochaines fenêtres ouvertes.")


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
def ensure_single_instance():
    """Empêche deux instances simultanées (sinon conflit sur le raccourci
    global). Renvoie le handle du mutex à conserver vivant, ou None si une
    autre instance tourne déjà."""
    try:
        mutex = win32event.CreateMutex(None, False, "TextEnhancerAI_singleton")
        if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
            return None
        return mutex
    except Exception:
        # En cas de souci avec le mutex, on n'empêche pas le lancement.
        return True


def main():
    mutex = ensure_single_instance()
    if mutex is None:
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo(APP_NAME,
                                "TextEnhancer AI est déjà en cours d'exécution "
                                "(icône dans la barre des tâches).")
            root.destroy()
        except Exception:
            pass
        return
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
