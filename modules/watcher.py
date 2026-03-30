import os
from pathlib import Path

ERLAUBTE_ENDUNGEN = [".jpg", ".jpeg", ".png"]

def suche_neue_bilder(ordner):
    bilder = []
    if not os.path.exists(ordner):
        return bilder
    for datei in os.listdir(ordner):
        if datei.startswith("."):
            continue
        if Path(datei).suffix.lower() in ERLAUBTE_ENDUNGEN:
            bilder.append(os.path.join(ordner, datei))
    return sorted(bilder)
