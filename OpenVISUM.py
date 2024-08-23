# ==============================
# Erstellt von: Ermal Sylejmani
# Datum: 08/08/2024
# ==============================# ==============================
# PTV Visum Skript
# ==============================
# Skript zum Öffnen einer PTV Visum-Projektdatei und zum Offenhalten der Anwendung,
# bis der Benutzer sie manuell schließt. Dies ermöglicht es, ohne Unterbrechungen 
# in der Visum-Umgebung zu arbeiten.
# 

# ==============================
#
# ==============================
# ANWENDUNGSANLEITUNG:
# 
# 1. Stelle sicher, dass du eine Sicherungskopie deiner .ver-Datei hast, bevor du das Skript ausführst.
# 2. Aktualisiere die Variable 'path_to_version' unten, um auf deine .ver-Datei zu verweisen.
# 3. Führe das Skript aus. Visum wird geöffnet und lädt die angegebene Projektdatei.
# 4. Das Visum-Fenster bleibt geöffnet, bis du im Konsolenfenster die Eingabetaste drückst.
# ==============================

import win32com.client

def Init(path=None):        
    Visum = win32com.client.Dispatch('Visum.Visum') 
    if path is not None: 
        Visum.LoadVersion(path) 
    return Visum

# Pfad zur Visum-Datei
# HINWEIS: Wenn du dieses Skript für ein anderes Projekt verwenden möchtest, musst du den Pfad unten
# an die Datei anpassen, die du in Visum öffnen möchtest.
path_to_version = r"C:\Users\syleer\OneDrive - INROS LACKNER SE\Desktop\Reaktivierungen Landkreis Lüneburg 2022\Inros_Lueneburg_testver.ver"

try:
    Visum = Init(path_to_version)
    if Visum:
        print("Visum erfolgreich geladen.")
        # Hält Visum geöffnet, bis du die Eingabetaste drückst
        input("Drücke die Eingabetaste, um Visum zu schließen...")
    else:
        print("Fehler beim Laden von Visum.")
except Exception as e:
    print(f"Fehler beim Initialisieren von Visum: {e}")

# Was du ändern musst, wenn du dieses Skript für ein anderes Projekt verwenden möchtest:
# 1. Ändere den Pfad bei 'path_to_version' zu dem Speicherort deiner .ver-Datei.
# 2. Wenn du das Skript für mehrere Projekte verwenden willst, 
#    kannst du den Pfad dynamisch über eine Eingabe abfragen oder in einer Schleife durch verschiedene Pfade iterieren.
