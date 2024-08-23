"""
Erstellt von: Ermal Sylejmani
Datum: 01/08/2024

Dieses Python-Skript verwendet die COM-Schnittstelle, um mit der PTV Visum-Software zu interagieren.
Es dient dem Zweck, ein Netzwerkmodell aus einer Visum-Netzwerkdatei zu laden, eine konfigurierbare Bilddatei des Netzwerks zu exportieren und das Visum-Programm anschließend zu schließen.

Verwendungszweck:
- Laden eines Netzwerkmodells aus einer angegebenen Visum-Datei.
- Konfigurieren von Druckparametern, um einen bestimmten Bereich des Netzwerks als Bilddatei zu exportieren.
- Speichern des Bildes in einem definierten Dateipfad.
- Schließen der Visum-Anwendung nach Abschluss des Prozesses.

Projektabhängige Anpassungen:
- Pfad zur Visum-Netzwerkdatei (`NETWORK_PATH`)
- Pfad für das gespeicherte Bild (`IMAGE_PATH`)
"""

import win32com.client
import os

# Pfade
# Pfad zur Visum-Netzwerkdatei (dies muss auf die spezifische Datei Ihres Projekts aktualisiert werden)
NETWORK_PATH = r"C:\Users\syleer\OneDrive - INROS LACKNER SE\Desktop\Reaktivierungen Landkreis Lüneburg 2022\Inros_Lueneburg_testver.ver"

# Pfad für das gespeicherte Bild (dies muss auf den gewünschten Speicherort und Dateinamen aktualisiert werden)
IMAGE_PATH = r"C:\Users\syleer\OneDrive - INROS LACKNER SE\Desktop\Reaktivierungen Landkreis Lüneburg 2022\network.jpg"

def main():
    visum = None
    try:
        # Erstellen eines COM-Objekts für Visum
        visum = win32com.client.Dispatch("Visum.Visum")
        print("Visum COM-Objekt erfolgreich erstellt.")
        
        # Laden des Netzwerkmodells in Visum
        if not os.path.exists(NETWORK_PATH):
            raise FileNotFoundError(f"Die Netzdatei existiert nicht: {NETWORK_PATH}")
        
        visum.LoadVersion(NETWORK_PATH)
        print(f"Netzwerk erfolgreich geladen: {NETWORK_PATH}")
        
        # Konfigurieren des Druckrahmens
        print_frame = visum.Net.PrintParameters.PrintFrame
        print_frame.BorderStyle.SetAttValue("COLOR", "ff0000ff")
        
        # Konfigurieren des Druckbereichs (dies muss möglicherweise angepasst werden, wenn andere Druckbereichseinstellungen erforderlich sind)
        print_area = visum.Net.PrintParameters.PrintArea
        left_margin = print_area.AttValue("LEFTMARGIN")
        bottom_margin = print_area.AttValue("BOTTOMMARGIN")
        right_margin = print_area.AttValue("RIGHTMARGIN")
        top_margin = print_area.AttValue("TOPMARGIN")
        
        # Exportieren des Netzwerks als Bilddatei (dies muss aktualisiert werden, wenn der Pfad oder die Bildqualität geändert werden soll)
        visum.Graphic.ExportNetworkImageFile(IMAGE_PATH, left_margin, bottom_margin, right_margin, top_margin, 1000)
        print(f"Netzwerk erfolgreich exportiert: {IMAGE_PATH}")
    
    except Exception as e:
        # Fehlerbehandlung im Hauptprozess
        print(f"Fehler im Hauptprozess: {e}")
    
    finally:
        # Schließen der Visum-Anwendung
        if visum:
            try:
                visum.Quit()  # Schließt die Visum-Anwendung
                print("Visum-Anwendung erfolgreich geschlossen.")
            except Exception as e:
                print(f"Fehler beim Schließen der Visum-Anwendung: {e}")

if __name__ == "__main__":
    main()
