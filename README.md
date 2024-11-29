[![License](https://img.shields.io/github/license/rkapl123/PfadfinderSchreiben.svg)](https://github.com/rkapl123/PfadfinderSchreiben/blob/master/LICENSE)

- Schreiben.xlsm kann zum Automatisieren diverser (v.a. Mitgliedsbeitragsbriefe) Schreiben an Mitglieder in Pfadfindervereinen genutzt werden.
- Erinnerungsmail.xlsm ist analog dazu zum Automatisieren reiner Erinnerungsmails an Mitglieder in Pfadfindervereinen gedacht (ohne Brief-PDF, dafür sind die Stammdaten zum Prüfen verfügbar).
- BeitrittserklärungImport\iGRINS_Import.xls kann in Kombination mit einem aus BeitrittserklärungImport\BeitrittserklärungFormular.odt erstellten Beitrittserklärungformular (befüllbares PDF) von einem Emailkonto die befüllten Formulare importieren. Diese ins Excel importierten Datensätze können dann gleich weiter in iGRINS (Gruppen Registrierungs und Informations System, [http://www.noe-pfadfinder.at/igrins](http://www.noe-pfadfinder.at/igrins)) importiert werden.

Die Installation erfolgt einfach mit Entpacken eines der beiden assets (source code .zip oder tar.gz) im letzten Release in ein beliebiges Verzeichnis, und anschliessend Öffnen einer der drei Dateien.

## Schreiben.xlsm und Erinnerungsmail.xlsm

Folgende drei Blätter sind in Schreiben.xlsm und Erinnerungsmail.xlsm wesentlich:

### Brief (Mitgliedsbeitragsbrief) bzw. Mailing (Erinnerungsmail)

Hier wird der zu erstellende Serienbrief gestaltet, der vorliegende Text (inkl. Grafik in der Kopfzeile) ist ein Vorschlag, wie er für einen Mitgliedsbeitragsbrief
verwendet werden kann. Die Ermittlung der Beiträge erfolgt parametergesteuert, wobei ein eher einfaches degressives Beitragsmodell verwendet wird:
Das erste Mitglied (Kind) zahlt den vollen Gruppen bzw. Verbandsbeitrag (Registrierung), ab dem zweiten Mitglied einer Familie wird ein ermäßigter Beitrag verrechnet (diese Ermäßigungen können mit den Parametern konfiguriert werden).
Erwachsene Mitglieder zahlen nur den Verbandsbeitrag.

Mit den Knöpfen "vor" und "zurück" kann man durch die Schreiben an die einzelnen Familien durchklicken,
mit "drucken & vor" druckt man das gerade sichtbare Schreiben aus und geht eine Familie nach vor (nur Mitgliedsbeitragsbrief).
mit "PDF & vor" speichert man das gerade sichtbare Schreiben als PDF (Nach- und Vorname(n) als Dateinamen) und geht eine Familie nach vor (nur Mitgliedsbeitragsbrief).
mit "Mail & vor" schickt man das gerade sichtbare Schreiben per E-Mail als PDF ("Mitgliedsbeitragsbrief" + zeile als Dateinamen) und geht eine Familie nach vor.

Der Mailversand wird entweder mit interaktiv mit Outlook oder mit dem Programm cmail.exe ([https://www.inveigle.net/cmail](https://www.inveigle.net/cmail)) bewerkstelligt.
Für cmail.exe ist der versendende Account im Range "mailuser" sowie das zugehörige Passwort im Bereich "mailpwd" einzustellen.
Im Bereich "testmail" kann mit "ja" ein Testlauf (Versand an eine beliebig einstellbare Adresse) getätigt werden.
Für Gmail Konten ist (für die Zeit des Versendens) der Zugriff durch weniger sichere (durch google genehmigte) Apps wie folgt zu aktivieren:
![Image1](https://raw.githubusercontent.com/rkapl123/PfadfinderSchreiben/master/googleMailEnable.PNG)  
![Image2](https://raw.githubusercontent.com/rkapl123/PfadfinderSchreiben/master/googleMailEnable2.PNG)  
Nach Abschluss des Versendens empfiehlt es sich, diesen Zugriff wieder zu deaktivieren.  

Wenn der Range "mailuser" nicht befüllt ist, wird stattdessen ein installiertes Outlook vorausgesetzt, in dem das E-Mail erzeugt und dargestellt wird.
Der Versand muss hier manuell betätigt werden.

Ein Schaltfeld "Nachregistrierung" ist für die Halbierung (bzw. sonstwie gestaltete Teilung) des vollen Gruppenmitgliedsbeitrags vorgesehen, hier werden die Gruppenmitgliedsbeiträge für ALLE
aktuell in der Liste befindlichen entsprechend gekürzt (nur Mitgliedsbeitragsbrief).

Als speziellen Zusatz gibt es noch am Ende des Briefs einen "Zahlen mit Code" Block, der für handelsübliche Mobile-Banking Apps verwendet werden kann.
Damit können die Zahlungsdaten aus dem Brief sofort ins E-Banking übernommen (und somit sehr einfach bezahlt) werden.
Hier ist die Anwendung zint.exe (von [http://zint.org.uk](http://zint.org.uk)) nötig, diese wird beim Betätigen der obigen Knöpfen mit dem Datenträger (Bereich "datenträger") aufgerufen

### Parameter

Im Blatt "Parameter" können die Parameter für die Serienbrieferstellung eingestellt werden, also:

nur relevant für Mitgliedsbeitragsbrief:
- Mitgliedsbeitrag: Der Gruppenmitgliedsbeitrag für das erste Mitglied (Kind)
- Mitgliedsbeitrag weitere K.: Der Gruppenmitgliedsbeitrag für alle weiteren Mitglieder einer Familie (Kinder)
- max. Anzahl Mitglbeiträge: Ab dieser Anzahl werden keine Gruppenmitgliedsbeiträge für Kinder eingehoben.
- Registrierung: Der Verbandsbeitrag (Registrierung) für das erste Mitglied (wird vom Landesverband bestimmt)
- Registrierung weitere K.: Der Verbandsbeitrag (Registrierung) für alle weiteren Mitglieder einer Familie (wird vom Landesverband bestimmt)
- Einschreibgebühr: Einmalige Gebühr bei der ersten Registrierung.

Allgemeine Parameter (auch in Erinnerungsmail):
- Kopfname: Name der Pfadfindergruppe im Briefkopf
- Pfadfinderkonto: Kontonummer der Pfadfindergruppe (IBAN)
- lautend auf: Name für das Konto
- Bank: Name der Bank, bei der das Konto geführt wird
- Adresse: Adresse der Pfadfindergruppe
- ZVR Zahl: Zentrale Vereinsregisterzahl der Pfadfindergruppe
- www: Website der Pfadfindergruppe
- mailto: Mailadresse, an die Antworten betreffend Abmeldung bzw. Korrektur der Stammdaten geschickt werden können (nur Erinnerungsmail)

### Liste

Im Blatt "Liste" werden die Familien/Personen geführt, die zur Erstellung der Schreiben herangzeogen werden.
Diese können entweder manuell gewartet werden, das Layout ist wie folgt:

| Anrede  | Name   | Adresse                          | Vorname1 | Einschreib1 | Ermäßigt1 | Vorname2 | Einschreib2 | Ermäßigt2 | ... | Summe |AnzKinder |AnzMitg |AnzEinschr |pfadis |
|---------|--------|----------------------------------|----------|-------------|-----------|----------|-------------|-----------|-----|-------|----------|--------|-----------|-------|
| Familie | Muster | Musterstrasse 8/15, Musterhausen | Thomas   | J           |           | Maria    |             |           | ... |Formeln werden eingetragen                     |
| Familie | ...    | ...                              |          |             |           |          |             |           | ... |                                               |
| ...     | ...    | ...                              | ...      | ...         | ...       | ...      | ...         | ...       | ... |                                               |


Die Spalten mit "Ermäßigt" dienen zur Unterscheidung von erwachsenen Mitgliedern, wenn hier ein "M" steht, zahlt das Mitglied nur den Registrierungsbeitrag. Wenn hier ein "E" für ehrenamtliche erwachsene Mitglieder steht, dann zahlt das Mitglied weder Registrierungsbeitrag (wird durch Gruppe übernommen) noch Gruppenmitgliedsbeitrag. Bereits registrierte Mitglieder werden auch mit "E" versehen, also auch nicht im Versand berücksichtigt.

Die Spalten mit "Einschreib" dienen zur Berücksichtigung des einmaligen Einschreibbetrags, wenn hier ein nichtleerer Eintrag steht, wird der unter Einschreibgebühr im Blatt "Parameter" stehende Betrag für dieses Mitglied berechnet. Erwachsene Mitglieder zahlen keinen Einschreibbetrag.

Die letzten fünf Spalten sind 
- die im Brief ausgewiesene Summe (kann zur Gegenprüfung der Eingänge verwendet werden)
- die Anzahl der Kinder in dieser Zeile
- die Anzahl der Mitglieder in dieser Zeile
- die Anzahl der Ersteinschreiber in dieser Zeile
- die gesammelten Namen der zahlenden Mitglieder in dieser Zeile

Die Formeln dazu werden beim Importieren aus iGrins immer automatisch gesetzt.

Wichtig für die Briefgestaltung ist, dass bei Familien mit Kindern und ehrenamtlichen Mitgliedern, die ehrenamtlichen Mitglieder immer nach den Kindern (rechts davon) eingetragen werden.

Für Nutzer von iGRINS (Gruppen Registrierungs und Informations System, [http://www.noe-pfadfinder.at/igrins](http://www.noe-pfadfinder.at/igrins)) gibt es einen Knopf zum Importieren aus einer angeführten Exportdatei, dazu muss diese vorher aus iGRINS exportiert werden:

In iGrins ins Menü person/xls export navigieren:
![Image3](https://raw.githubusercontent.com/rkapl123/PfadfinderSchreiben/master/export.PNG)

Dort dann den Filter entsprechend einstellen (wobei ohnehin nur aktive übernommen werden):
![Image4](https://raw.githubusercontent.com/rkapl123/PfadfinderSchreiben/master/exportAuswahl.PNG)

Dann auf exportieren klicken und den Download der Datei bestätigen:
![Image5](https://raw.githubusercontent.com/rkapl123/PfadfinderSchreiben/master/excel.PNG)

Dieser Import wird mit den folgenden Einstellungen gesteuert:
1. Exportfile: Namen der iGrins Exportdatei, beim Öffnen wird diese Zelle mit der aktuellsten Exportdatei im Verzeichnis befüllt.
2. Nicht importieren wenn "Ermäßigt" auf: Mitglieder mit diesen Einträgen in Spalte "Ermaessigt" werden ausgenommen (z.b. "P" oder "PWA").
3. Neuregistrierungen ab: Wenn das Datum in Spalte "AngelegtAm" grösser ist als dieser Wert, dann ist die einmalige Einschreibgebühr zu entrichten; Hier kann der Stichtag (Monat und Tag) des Pfadijahres (siehe oben) eingestellt werden.
4. Funktionen ehrenamtlicher Mitglieder (Elternrat, Leiter): Diese Funktionen werden zur Unterscheidung zu den Kindern ("WI", "WÖ", "GU", "SP", "CA", "EX", "RA", "RO" -> hartcodiert) und "normalen" erwachsenen Mitgliedern ("MIT") herangezogen, Mitglieder mit dieser Funktion bekommen unter Ermäßigt ein "E" und zahlen weder Verbandsbeitrag noch Gruppenbeitrag. Leiter werden hardcodiert (like "F*" und like "AS*") ermittelt.
5. Gruppierung beim Export nach: Sortiert die iGrins Liste nach diesen beiden Spalten und gruppiert dann auch die Familien danach. Wenn die Adressen gut gepflegt (= exakt gleich für Familienmitglieder) sind, empfiehlt sich Adresse/Plz, ansonsten ist Nachname/Plz besser (mit der Gefahr, dass ein Name doppelt in der Plz vorkommt bzw. Personen im selben Haushalt mit verschiedenen Nachnamen nicht als eine Familie erkannt werden)

Die Übernahme aus iGRINS ist aktuell auf 6 Mitglieder pro Familie beschränkt, wenn mehr benötigt werden, dann muss die Breite der Liste entsprechend erweitert werden und die Formeln, die auf die Liste bezug nehmen angepasst werden (im Wesentlichen müssen analog zu den bestehenden Formeln die erweiterten Teile hinzugefügt werden).

Einstellung für "Pfadijahr beginnt": das Jahr, in dem das aktuelle Pfadijahr begint, wird für die Beschriftungen im Brief gebraucht. Üblicherweise wird hier das aktuelle Jahr genommen (berechnet aus der eingestellten iGrins Datei).

### technisches Blatt: Beispiel

Das Blatt "Beispiel" dient nur zur einfachen Befüllung der Blätter "Liste" und "Parameter" mit Beispieldaten, wenn auf den Knopf "Schreiben in Github Ordner" geklickt wird (Entfernt alle personen- und gruppenbezogenen Daten).
Wird nur zu Entwicklungszwecken benötigt und kann (inklusive Knopf) entfernt werden.

## BeitrittserklärungImport\iGRINS_Import.xls

Die Datei iGRINS_Import.xls im Unterverzeichnis BeitrittserklärungImport kann zum importieren von Beitrittserklärungen, die als befüllte PDF Formulare in Mails auf einem definierten Mailaccount liegen.
Dazu werden zwei Tools verwendet, zum Herunterladen der Mail(attachments) das Programm curl (seit Windows 10 Bestandteil von MS-Windows, für neuere Versionen siehe [https://curl.se/windows/](https://curl.se/windows/)), zum extrahieren der Daten aus den PDF-Attachments (befüllte Formulare) das Programm pdftk ([https://www.pdflabs.com/tools/pdftk-server/](https://www.pdflabs.com/tools/pdftk-server/)). Wenn das pdftk Verzeichnis nicht in den Environment-Pfad aufgenommen wurde, dann müssen libiconv2.dll und pdftk.exe ins Verzeichnis dieser Datei kopiert werden.

Der definierte Mailaccount kann gleich beim Öffnen der Datei abgefragt werden:
![Image6](https://raw.githubusercontent.com/rkapl123/PfadfinderSchreiben/master/BeitrittserklärungImport1.png)

Zum Befüllen der Datei wird auch noch ein Startindex (StartID) benötigt, der die nächste ID nach der aktuell höchsten in iGRINS darstellt (links oben einzutragen).
Weiters muss ein Präfix für den Bundesverbandschlüssel (BVKey, z.b. "3-GAB-") und ein Gruppenschlüssel (Grp-key, z.b. "GAB") definiert werden (ja, man hätte das eine aus dem anderen ableiten können, trotzdem ist der BVkey noch speziell...)
Falls man ein anderes Datum als heute zur Anlage der Datensätze in iGRINS verwenden will, kann das unter "AngelegtAm:" eingetragen werden.  

Unter "server", "username", "passwort" und "attachmentverz" werden die wesentlichen parameter 
- server: URL imap server, 
- username: mailadresse/username auf dem imap server, 
- passwort: dazugehöriges passwort, 
- attachmentverz: Verzeichnis unter dem die PDF Dateien abgelegt werden sollen (standard "Formulare")

eingetragen. 
Die Mails werden über IMAP von curl heruntergeladen und auf dem Server nicht gelöscht:
![Image7](https://raw.githubusercontent.com/rkapl123/PfadfinderSchreiben/master/BeitrittserklärungImport2.png)

Beim Importvorgang werden die PDF-attachments vom Server geholt (ins Attachment Verzeichnis) und anschliessend mit pdftk ausgelesen. Dieser Vorgang ist hardcodiert und bennötigt daher ein genau passendes Formular, das mit dem LibreOffice odt Dokument an die Gruppenbedürfnisse (Logo, diverse Angaben zu Mitgliedsbeiträgen, DSGVO, etc.) angepasst werden kann.
Die Vorlage sollte die notwendigsten Daten (die auch zukünftigen Mitgliedern zum Ausfüllen zugemutet werden können) umfassen, wenn mehr Felder notwendig sind, dann ist sowohl das odt anzupassen, als auch der Importcode (das Ergebnis von pdftk wird vom StdOut gelesen und in ein assoziatives Array (Feldname/Feldwert paare) gespeichert, das dann in die zeilen eingetragen wird).

Es werden dann auch zusätzlich Fixwerte und Formeln für abgeleitete Werte unter den Feldern ID,Art,Geschlecht,BVKey,Aktiv,Funktion,Grp-Key und AngelegtAm eingetragen, wenn diese nicht passen sollten, dann ganz unten in der (Haupt)Prozedur fillValues ändern.
Das Feld "Geschlecht" wird nicht als einzutragendes Feld erwartet sondern mit einer hinreichend umfangreichen Datenbank von Vornamen in der Datei "Geschlecht.xls" herausgefunden. Daher ist es wichtig bei offenkundigen Fehlern bzw. nicht gefundenen Namen (dann steht "unspezifisch" als Geschlecht) die Formel einfach mit M oder W auszubessern.
Das Feld Funktion (WI,WÖ,GU,SP,CA,EX,RA,RO) wird ebenfalls automatisch ausgefüllt und zwar mithilfe des Geschlechts und des Geburtsdatums. 
Die Altersgrenzen der Stufen sind im Blatt "FunktionAlter" zu finden, sie entsprechen bis auf WI/WÖ den Altersstufen der PPÖ [https://ppoe.at/ueber-uns/](https://ppoe.at/ueber-uns/))
Der Bundesverbandsschlüssel wird relativ simpel mit Präfix+Jahr+ID gebildet, das gewährleistet zum einen eine Erkennung des Ersteintrittsjahres, zum anderen eine Eindeutigkeit (über Präfix und Jahr).

Beim Speichern wird noch die Möglichkeit geboten, die StartID auf die nächsthöhere ID der gerade importierten zu stellen:
![Image8](https://raw.githubusercontent.com/rkapl123/PfadfinderSchreiben/master/BeitrittserklärungImport3.png)

Anschliessend kann mit der Importfunktion von iGRINS die Datei importiert werden:
![Image9](https://raw.githubusercontent.com/rkapl123/PfadfinderSchreiben/master/BeitrittserklärungImport4.png)
