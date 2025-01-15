import pandas as pd
from caldav import DAVClient
from icalendar import Calendar, Event
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import configparser
import os

# Funktion zum Erstellen eines iCalendar-Events
def create_ical_event(summary, start, end, description=""):
    event = Event()
    event.add('summary', summary)
    event.add('dtstart', start)
    event.add('dtend', end)
    event.add('description', description)
    return event

# Funktion zum Einfügen eines Events in den Kalender
def add_event_to_calendar(client, calendar_url, event):
    calendar = client.calendar(url=calendar_url)
    calendar.save_event(event.to_ical())

# Funktion zum Bereinigen der CSV-Daten
def clean_dataframe(df):
    # Entferne führende '=' und '"' von den Spaltennamen
    df.columns = df.columns.str.strip('=\"')
    
    # Entferne führende '=' und '"' von den Zellenwerten
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip('=\"')
    return df

# Funktion zum Verarbeiten der CSV-Datei
def process_csv(csv_file, username, password):
    # Verbindung zum Kerio Connect Server herstellen mit Fehlerbehandlung
    try:
        client = DAVClient('https://kerio1.kampmail.de/caldav/')
        client.login(username, password)
    except Exception as e:
        messagebox.showerror("Verbindungsfehler", f"Fehler bei der Verbindung zum Kalender: {e}")
        return

    # Kalender-URL (dies kann variieren, abhängig von deiner Konfiguration)
    calendar_url = f'https://kerio1.kampmail.de/caldav/{username}/kalender/'

    # CSV-Datei laden und bereinigen
    try:
        df = pd.read_csv(csv_file, delimiter=';', decimal=',', encoding='utf-8-sig')
        df = clean_dataframe(df)
    except Exception as e:
        messagebox.showerror("CSV-Fehler", f"Fehler beim Laden der CSV-Datei: {e}")
        return

    # Durch die CSV-Datei iterieren und Events erstellen
    for index, row in df.iterrows():
        objekt = row['Objekt']
        mitarbeiter = row['Mitarbeiter']
        datum = row['Datum']
        von = row['Von']
        bis = row['Bis']

        # Termintitel basierend auf den Regeln setzen
        if objekt == 'PRO NSL':
            summary = 'ProSi: NSL'
        elif objekt == 'PRO Mitarbeiter':
            summary = 'ProSi: Backoffice'
        elif objekt == 'Kinderklinik SEP':
            summary = 'ProSi: VKJK'
        else:
            summary = objekt

        # Datum und Uhrzeit parsen mit Fehlerbehandlung
        try:
            start = datetime.strptime(f"{datum} {von}", '%d.%m.%Y %H:%M:%S')
            end = datetime.strptime(f"{datum} {bis}", '%d.%m.%Y %H:%M:%S')
        except ValueError as ve:
            messagebox.showerror("Datumsfehler", f"Fehler beim Parsen des Datums oder der Uhrzeit: {ve}")
            continue

        event = create_ical_event(summary, start, end, description=f"Mitarbeiter: {mitarbeiter.strip()}")
        try:
            add_event_to_calendar(client, calendar_url, event)
        except Exception as e:
            messagebox.showerror("Kalenderfehler", f"Fehler beim Hinzufügen des Events zum Kalender: {e}")
            continue

    # Verbindung schließen
    try:
        client.logout()
    except Exception as e:
        messagebox.showwarning("Abmeldefehler", f"Fehler beim Abmelden vom Kalender: {e}")

    messagebox.showinfo("Erfolg", "Termine erfolgreich importiert!")

# Funktion zum Speichern der Anmeldedaten
def save_credentials(username, password):
    config = configparser.ConfigParser()
    config['Kerio'] = {'Username': username, 'Password': password}
    with open('config.ini', 'w') as configfile:
        config.write(configfile)

# Funktion zum Laden der Anmeldedaten
def load_credentials():
    config = configparser.ConfigParser()
    if os.path.exists('config.ini'):
        config.read('config.ini')
        return config['Kerio']['Username'], config['Kerio']['Password']
    return '', ''

# GUI-Funktion
def create_gui():
    root = tk.Tk()
    root.title("Kerio Connect CSV Import")

    tk.Label(root, text="Kerio Benutzername:").grid(row=0, column=0, padx=10, pady=10)
    tk.Label(root, text="Kerio Passwort:").grid(row=1, column=0, padx=10, pady=10)

    username_entry = tk.Entry(root)
    password_entry = tk.Entry(root, show="*")

    username_entry.grid(row=0, column=1, padx=10, pady=10)
    password_entry.grid(row=1, column=1, padx=10, pady=10)

    def select_file():
        file_path = filedialog.askopenfilename(filetypes=[("CSV Dateien", "*.csv")])
        if file_path:
            username = username_entry.get()
            password = password_entry.get()
            save_credentials(username, password)
            process_csv(file_path, username, password)

    tk.Button(root, text="CSV-Datei auswählen", command=select_file).grid(row=2, column=0, columnspan=2, pady=10)

    # Anmeldedaten laden
    username, password = load_credentials()
    username_entry.insert(0, username)
    password_entry.insert(0, password)

    root.mainloop()

# GUI starten
create_gui()
