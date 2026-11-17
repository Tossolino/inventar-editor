# Inventar-Editor

Eigenständiges Windows-Programm zum Bearbeiten der geöffneten Excel-Prüfliste
(Blatt „akt. Med.-Produkte“) über die Inventarnummer. Die Excel-Datei enthält
dafür keine Makros; das Programm steuert das laufende Excel von außen.

## Installation

1. Unter **Releases** die neueste `InventarEditor.exe` herunterladen.
2. In einen eigenen Ordner legen, z.B. `Dokumente\InventarEditor`
   (nicht unter „Programme“ – sonst funktionieren Updates nicht).
3. Excel-Liste öffnen, dann `InventarEditor.exe` starten.

Beim ersten Start kann Windows SmartScreen warnen („Unbekannter Herausgeber“):
*Weitere Informationen → Trotzdem ausführen*.

## Bedienung

| Aktion | So geht's |
|---|---|
| Inventarnummer laden | Nummer eintippen, **Enter** |
| Feld wechseln | **Tab** |
| Übernehmen (schreibt zurück + „x“ in Spalte L) | **Enter** in einem Feld oder Knopf *Übernehmen* |
| Zurück ins Nummernfeld | **Esc** |
| Mappe speichern | **Strg+S** oder Knopf *Speichern* |

Statuspunkt: **grün** = Q, R, S leer · **gelb** = Mangel (Q) ohne Klasse (R) ·
**rot** = Mangel mit Klasse (R ausgefüllt). Bei gelb/rot werden Q, R und S
automatisch eingeblendet und beim nächsten Gerät wieder ausgeblendet.

Neues Bett: Ist Feld E (Typ/Modell) leer, öffnet sich eine Auswahl der Typen aus
„PRO Pflegebett“ (P7 abwärts, Hersteller aus O). Nach „OK“ werden C = Pflegebett,
D = Hersteller, E = Typ, H = CE und O = Heim eingetragen – gespeichert wird erst mit *Übernehmen*.

Formelzellen werden grau angezeigt und können nicht bearbeitet werden.
Die Inventarnummern in Spalte A werden nie verändert.

## Neue Version veröffentlichen

1. In `version.py` die Versionsnummer erhöhen (`V1.000.26` → `V1.001.26`).
2. Committen und auf `main` pushen – die Commit-Nachricht wird zur
   Änderungsbeschreibung, die die App beim Update anzeigt.
3. GitHub baut die EXE automatisch (Reiter *Actions*, ca. 3 Minuten) und
   veröffentlicht sie als Release. Die App zeigt dann „Update verfügbar“.
