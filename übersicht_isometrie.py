
import tkinter as tk
from tkinter import messagebox
import subprocess


# Funktion zum Starten einer .exe-Anwendung
def start_application(app_path, app_name):
    try:
        # Nachricht ins Status-Textfeld schreiben
        status_text.insert(tk.END, f"Anwendung '{app_name}' wird gestartet...\n")
        status_text.see(tk.END)  # Scrollt automatisch nach unten
        subprocess.Popen(app_path, shell=True)
    except FileNotFoundError:
        messagebox.showerror("Fehler", f"Die Anwendung {app_path} wurde nicht gefunden.")
        status_text.insert(tk.END, f"Fehler: Anwendung '{app_name}' konnte nicht gestartet werden.\n")
        status_text.see(tk.END)  # Scrollt automatisch nach unten


# Haupt-GUI-Fenster erstellen
root = tk.Tk()
root.title("Isokinetik-Anwendungen")
root.geometry("800x800")

# Anwendungen mit Button und Beschreibung
apps = [
    {"name": "Datenvorbereitung", "path": r"K:\Team\Böhmer_Michael\Anwendungen\Datenvorbereitung_Isokinet\datenvorbereitung_isokinet.exe", "description": "Daten für die Auswertung vorbereiten."},
    {"name": "Datenauswertung", "path": r"K:\Team\Böhmer_Michael\Anwendungen\Isokinet_isometrisch\isometrie_schritt_1.exe", "description": "Auswertung durchführen. Daten mit unerwartetem Verlauf werden gekennzeichnet."},
    {"name": "Nachbearbeitung", "path": r"K:\Team\Böhmer_Michael\Anwendungen\Isokinet_isometrisch\isometrie_markieren.exe", "description": "Daten zum Nachbearbeiten werden markiert."}
]

# Anwendungen hinzufügen
for app in apps:
    frame = tk.Frame(root)
    frame.pack(pady=10)  # Abstand zwischen den Abschnitten

    button = tk.Button(frame, text=app["name"], command=lambda p=app["path"], n=app["name"]: start_application(p, n), width=25)
    button.pack()  # Button oben

    label = tk.Label(frame, text=app["description"], anchor="center", wraplength=500)
    label.pack()  # Beschreibung unten

# Textfeld für Notizen hinzufügen
text_frame = tk.Frame(root)
text_frame.pack(fill="both", expand=True, padx=10, pady=10)

text_label = tk.Label(text_frame, text="Platz für Notizen:")
text_label.pack(anchor="w")

text_entry = tk.Text(text_frame, height=10)
text_entry.pack(fill="both", expand=True)

# Status-Textfeld hinzufügen
status_frame = tk.Frame(root)
status_frame.pack(fill="x", padx=10, pady=10)

status_label = tk.Label(status_frame, text="Status:")
status_label.pack(anchor="w")

status_text = tk.Text(status_frame, height=5, state="normal", bg="#f0f0f0")
status_text.pack(fill="x", expand=True)

# GUI starten
root.mainloop()
