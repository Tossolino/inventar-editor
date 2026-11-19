# Stichwort „*fernbedienung“ – Bedienung per Tablet (vertagt)

Stand: 17.11.2026 – noch nicht umgesetzt, nur Planung.

## Entscheidungen
- **Weg 2:** Das Tool bekommt einen eingebauten kleinen Webserver; das Tablet
  öffnet eine eigene, für Touch gebaute Seite im Browser (keine App-Installation,
  als Verknüpfung auf den Startbildschirm legbar). Keine Bildschirmspiegelung.
- **Gerät:** Android-Tablet.
- **Verbindung:** Mobiler Hotspot des Notebooks (Windows: Einstellungen →
  Netzwerk und Internet → Mobiler Hotspot). Kein WLAN vor Ort, Einsatzort wechselt.
  Kein Bluetooth (Browser kann es nicht, langsamer, Koppeln hakelig).

## Geplanter Umfang (erste Ausbaustufe)
- Inventarnummer eingeben, Felder anzeigen/ändern, Übernehmen
- Haken für PRO Pflegebett (Aufrichter, Bettleuchte, Seitengitter, B-Maß,
  Bedienteil, Netzstecker, Furnierschaden)
- Drucken, PRO Pflegebett umschalten, Typauswahl bei leerem E
- Notebook-Tool und Tablet zeigen denselben Stand

## Technische Punkte
- Adresse im Hotspot fest; Tool zeigt Adresse als QR-Code zum Scannen
- PIN-Schutz für die Seite
- Windows-Firewall fragt beim ersten Start einmal nach
- Alle Excel-Zugriffe laufen weiter im Hauptprogramm (Warteschlange),
  da Excel/COM nur aus einem Thread sicher bedient werden kann
- Zu prüfen: Lässt sich der Hotspot auf dem Notebook ohne bestehende
  Netzverbindung einschalten?

## E-Ink-Geräte (Nachtrag 19.11.2026)
- **Kindle Paperwhite:** nicht empfohlen. Nur „experimenteller Browser“ (langsam,
  eingeschränktes JavaScript, träge Bildschirmtastatur).
- **Android-E-Ink-Tablets** (z.B. Onyx Boox Note Air / Tab / Go, Bigme): geeignet,
  laufen mit Chrome wie ein normales Android-Tablet.
- Für E-Ink: optionaler **Schwarz-Weiß-Modus** der Seite – nichts nur über Farbe
  unterscheiden (aktive Kachel = schwarz invertiert/dicker Rahmen statt grün,
  rote Haken = durchgestrichen/„NEIN“ fett), große Schaltflächen, keine Animationen,
  Seite nur bei echten Änderungen neu zeichnen (weniger Geisterbilder).
