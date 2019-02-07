# office_open_xml_examples##

## Ziel dieses Projekts
Mit Pentaho Report Designer ist es nicht möglich, docx-Berichte zu erstellen. Dieses Projekt war ein Versuch, Berichte im prpt-Format einzulesen und mithilfe der Apache-POI-Bibliothek programmatisch ein docx-Dokument zu erstellen.

## Funktionsweise dieses Java-Projekts
Die Klasse mit der main-Methode ist `de.fu_berlin.softwareprojekt1819.openxmlcreation.WordDocument2`.
Diese Klasse liest aus dem Ordner `template` und der Datei `template.prpt` die im Archiv enthaltenen Dateien und erstellt dann die docx-Datei.

## Hinweise zur Verbesserung
Mithilfe des Pentaho-Report-Designer-SDKs kann man die Berichtsdefintion von prpt-Dateien einlesen und darauf zugreifen. Aus diesen Informationen wird die docx-Datei erstellt. Das ist sehr wahrscheinlich eine bessere Vorgehensweise, als auf die Dateien des prpt-Archivs zuzugreifen und sich mit dem XML-Schema auseinanderzusetzen.

## Bekannte Probleme
prpt-Dateien sind normale Archive, die entpackt werden können. Momentan muss sowohl die Datei als auch das entpackte Archiv vorhanden sein, damit dieses Programm funktioniert. Es klappt noch nicht, das XML aus der ZIP-Datei zu parsen. (Es werden Fehler geworfen.) Und ich habs noch nicht angepasst, dass man nur den Ordner braucht und nicht die Datei.

## Weiterführende Links
Pentaho-Doku: https://help.pentaho.com/Documentation/8.2/Developer_Center/Embed_Reporting/In_Java_Application
