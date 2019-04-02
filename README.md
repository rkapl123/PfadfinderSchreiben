# Schreiben Automatisierung Pfadfinder

Die Datei Schreiben.xlsm kann zum Automatisieren diverser (v.a. Mitgliedsbeitragsbriefe) Schreiben an Mitglieder in Pfadfindervereinen genutzt werden.

Folgende drei Blätter sind darin wesentlich:

## Brief

Hier wird der zu erstellende Serienbrief gestaltet, der vorliegende Text (inkl. Grafik in der Kopfzeile) ist ein Vorschlag, wie er für einen Mitgliedsbeitragsbrief
verwendet werden kann. Die Ermittlung der Beiträge erfolgt parametergesteuert, wobei ein eher einfaches degressives Beitragsmodell verwendet wird:
Das erste Mitglied (Kind) zahlt den vollen Gruppen bzw. Verbandsbeitrag (Registrierung), ab dem zweiten Mitglied einer Familie wird ein ermäßigter Beitrag verrechnet (diese Ermäßigungen können mit den Parametern konfiguriert werden).
Erwachsene Mitglieder zahlen nur den Verbandsbeitrag.

Mit den Knöpfen "vor" und "zurück" kann man durch die Schreiben an die einzelnen Familien durchklicken,
mit "drucken & vor" druckt man das gerade sichtbare Schreiben aus und geht eine Familie nach vor.
mit "PDF & vor" speichert man das gerade sichtbare Schreiben als PDF (Nach- und Vorname(n) als Dateinamen) und geht eine Familie nach vor.

Als speziellen Zusatz gibt es noch am Ende des Briefs einen "Zahlen mit Code" Block, der für handelsübliche Mobile-Banking Apps verwendet werden kann.
Damit können die Zahlungsdaten aus dem Brief sofort ins E-Banking übernommen (und somit sehr einfach bezahlt) werden.
Hier ist die Anwendung zint.exe (von http://zint.org.uk) nötig, diese wird beim Betätigen der obigen Knöpfen mit dem Datenträger (Bereich "datenträger") aufgerufen

## Parameter

Im Blatt "Parameter" können die Parameter für die Serienbrieferstellung eingestellt werden, also:

- Mitgliedsbeitrag: Der Gruppenmitgliedsbeitrag für das erste Mitglied (Kind)
- Mitgliedsbeitrag weitere K.: Der Gruppenmitgliedsbeitrag für alle weiteren Mitglieder einer Familie (Kinder)
- max. Anzahl Mitglbeiträge: Ab dieser Anzahl werden keine Gruppenmitgliedsbeiträge für Kinder eingehoben.
- Registrierung: Der Verbandsbeitrag (Registrierung) für das erste Mitglied (wird vom Landesverband bestimmt)
- Registrierung weitere K.: Der Verbandsbeitrag (Registrierung) für alle weiteren Mitglieder einer Familie (wird vom Landesverband bestimmt)
- Einschreibgebühr: Einmalige Gebühr bei der ersten Registrierung.

- Kopfname: Name der Pfadfindergruppe im Briefkopf
- Pfadfinderkonto: Kontonummer der Pfadfindergruppe (IBAN)
- lautend auf: Name für das Konto
- Bank: Name der Bank, bei der das Konto geführt wird
- Adresse: Adresse der Pfadfindergruppe
- ZVR Zahl: Zentralen Vereinsregisterzahl der Pfadfindergruppe
- www: Website der Pfadfindergruppe


## Liste

Im Blatt "Liste" werden die Familien/Personen geführt, die zur Erstellung der Schreiben herangzeogen werden.
Diese können entweder manuell gewartet werden, das Layout ist wie folgt:

| Anrede  | Name   | Adresse                          | Vorname1 | Einschreib1 | Ermäßigt1 | Vorname2 | Einschreib2 | Ermäßigt2 | ... | Summe |
|---------|--------|----------------------------------|----------|-------------|-----------|----------|-------------|-----------|-----|-------|
| Familie | Muster | Musterstrasse 8/15, Musterhausen | Thomas   | J           |           | Maria    |             |           | ... |Formel |
| Familie | ...    | ...                              |          |             |           |          |             |           | ... |wird   |
| ...     | ...    | ...                              | ...      | ...         | ...       | ...      | ...         | ...       | ... |eingetr|


Die Spalten mit "Ermäßigt" dienen zur Unterscheidung von erwachsenen Mitgliedern, wenn hier ein nichtleerer Eintrag steht, zahlt das Mitglied nur den Verbandsbeitrag.
Die Spalten mit "Einschreib" dienen zur Berücksichtigung des einmaligen Einschreibbetrags, wenn hier ein nichtleerer Eintrag steht, wird der unter Einschreibgebühr im Blatt "Parameter" stehende Betrag für dieses Mitglied berechnet.

Für Nutzer von iGRINS (Gruppen Registrierungs und Informations System, http://www.noe.pfadfinder.at/igrins) gibt es einen Knopf zum Importieren aus einer angeführten Exportdatei, dazu muss diese vorher aus iGRINS exportiert werden:

In iGrins ins Menü person/xls export navigieren:
![Image1](https://raw.githubusercontent.com/rkapl123/PfadfinderSchreiben/master/export.PNG)

Dort dann den Filter entsprechend einstellen (wobei ohnehin nur aktive übernommen werden):
![Image2](https://raw.githubusercontent.com/rkapl123/PfadfinderSchreiben/master/exportAuswahl.PNG)

Dann auf exportieren klicken und den Download der Datei bestätigen:
![Image3](https://raw.githubusercontent.com/rkapl123/PfadfinderSchreiben/master/excel.PNG)

Dieser Import wird mit sechs Einstellungen gesteuert:
1. Leiter aus iGRINS übernehmen (zahlen selbst): Ja=es werden auch die Leiter übernommen und bekommen einen Mitgliedsbeitragsbrief
2. Pfadijahr beginnt: Jahr, in dem das aktuelle Pfadijahr begint, wird für die Beschriftungen im Brief gebraucht. Üblicherweise wird hier das Vorjahr des aktuellen Jahres genommen (berechnet).
3. Nicht importieren wenn "Ermäßigt" auf: Mitglieder mit diesen Einträgen in Spalte "Ermaessigt" werden ausgenommen.
4. Neuregistrierungen ab: Wenn Datum in Spalte "AngelegtAm" grösser als dieser Wert, dann ist die einmalige Einschreibgebühr zu entrichten; Hier kann der Stichtag (Monat und Tag) des Pfadijahres (siehe oben) eingestellt werden.
5. Funktionen Erwachsene Mitglieder: Diese Funktionen werden zur Unterscheidung zu den Kindern ("WI", "WÖ", "GU", "SP", "CA", "EX", "RA", "RO") herangezogen, Mitglieder mit dieser Funktion bekommen unter Ermäßigt ein "E" und zahlen nur den Verbandsbeitrag.
6. Gruppierung beim Export nach: Sortiert die iGrins Liste nach diesen beiden Spalten und gruppiert dann auch die Familien danach. Wenn die Adressen gut gepflegt sind, empfiehlt sich Adresse/Plz, ansonsten ist Nachname/Plz besser (mit der Gefahr, dass ein Name doppelt in der Plz vorkommt bzw. Personen im selben Haushalt mit verschiedenen Nachnamen nicht als eine Familie erkannt werden)

Wenn "Leiter übernehmen" auf Nein steht, dann werden alle Mitglieder mit Funktionen, die weder Kinder noch erwachsene Mitglieder (Elternrat, Unterstützer, etc.) kennzeichnen (also alle mit Leiterfunktionen) NICHT übernommen.
Die Übernahme aus iGRINS ist aktuell auf 6 Mitglieder pro Familie beschränkt, wenn mehr benötigt werden, dann muss die Breite der Liste entsprechend erweitert werden und die Formeln im Blatt Brief, die auf die Liste bezug nehmen angepasst werden (im Wesentlichen analog zur bestehenden Formel fehlende Teile hinzugefügt werden).

## technisches Blatt: Beispiel

Das Blatt "Beispiel" dient nur zur einfachen Befüllung der Blätter "Liste" und "Parameter" mit Beispieldaten, wenn auf den Knopf "Schreiben in Github Ordner" geklickt wird (Entfernt alle personen- und gruppenbezogenen Daten).
Wird nur zu Entwicklungszwecken benötigt und kann (inklusive Knopf) entfernt werden.
