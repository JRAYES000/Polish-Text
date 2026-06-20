"""
Tests unitaires des fonctions « pures » de TextEnhancer AI.
Exécutés par la CI (GitHub Actions) AVANT la compilation, pour éviter qu'une
régression parte en release. N'ouvre aucune fenêtre (pas de Tk).
Lancer : python test_core.py
"""
import re
import sys

import main  # importe le module (ne lance pas l'application)

failures = []


def check(label, condition):
    if condition:
        print(f"  OK   {label}")
    else:
        print(f"  FAIL {label}")
        failures.append(label)


print("== Comparaison de versions ==")
check("1.0.1 > 1.0.0", main.is_newer("v1.0.1", "1.0.0") is True)
check("1.0.0 == 1.0.0 (pas plus récent)", main.is_newer("1.0.0", "1.0.0") is False)
check("1.10.0 > 1.9.0 (numérique)", main.is_newer("1.10.0", "1.9.0") is True)
check("0.9.0 < 1.0.0", main.is_newer("v0.9.0", "1.0.0") is False)
check("_parse_version v2.3 -> (2,3)", main._parse_version("v2.3") == (2, 3))

print("== Normalisation / parsing du gras ==")
check("<b> -> **", main._normalize_bold("<b>x</b>") == "**x**")
check("<strong> -> **", main._normalize_bold("<strong>y</strong>") == "**y**")
check("__z__ -> **z**", main._normalize_bold("__z__") == "**z**")
runs = main._parse_bold_runs("a **b** c")
check("segment gras détecté", ("b", True) in runs)
check("segment normal détecté", ("a ", False) in runs)

print("== Offsets CF_HTML (texte enrichi presse-papiers) ==")
frag = "Bonjour <b>Axel</b> éàù"
full = main._build_cf_html(frag)
b = full.encode("utf-8")
m = dict(re.findall(r"(StartFragment|EndFragment):(\d+)", full))
sf, ef = int(m["StartFragment"]), int(m["EndFragment"])
check("StartFragment/EndFragment encadrent exactement le fragment",
      b[sf:ef].decode("utf-8") == frag)
check("meta charset utf-8 présent (accents corrects au collage)",
      'meta charset="utf-8"' in full)

print("== Messages d'erreur OpenRouter ==")
check("401 -> message clé", "401" in main._friendly_openrouter_error(401, ""))
check("402 -> message crédit", "402" in main._friendly_openrouter_error(402, ""))
check("429 -> message quota", "429" in main._friendly_openrouter_error(429, ""))

print("== Géométrie fenêtre (placement multi-écrans) ==")
geo = main.compute_window_geometry(780, 640, True)
check("renvoie une géométrie Tk valide",
      bool(re.match(r"^\d+x\d+(\+-?\d+\+-?\d+)?$", geo)))

print("== Libellé preset (raccourci affiché) ==")
check("affiche le raccourci",
      main.PreviewWindow._preset_label({"name": "Reformuler", "hotkey": "alt+q"})
      .endswith("alt+q"))
check("sans raccourci -> nom seul",
      main.PreviewWindow._preset_label({"name": "X", "hotkey": ""}) == "X")

print("== Streaming ==")
import inspect
check("stream_openrouter est un générateur",
      inspect.isgeneratorfunction(main.stream_openrouter))

print("== Thème ==")
check("palette claire et sombre définies",
      "light" in main.PALETTES and "dark" in main.PALETTES)
check("get_palette défaut = clair",
      main.get_palette({}) is main.PALETTES["light"])

print("== Liens (parsing) ==")
segs = main.parse_segments(
    "Voir [Notion](https://app.notion.com/p/x?source=copy_link) ici")
check("lien Markdown -> texte + url",
      any(u == "https://app.notion.com/p/x?source=copy_link" and t == "Notion"
          for (t, b, u) in segs))
check("aucun marqueur []() dans les segments",
      not any("](" in t for (t, b, u) in segs))
segs2 = main.parse_segments("Lien nu https://example.com/a fin.")
check("URL nue détectée comme lien",
      any(u == "https://example.com/a" for (t, b, u) in segs2))

print()
if failures:
    print(f"ÉCHEC : {len(failures)} test(s) en échec -> {failures}")
    sys.exit(1)
print("TOUS LES TESTS PASSENT.")
