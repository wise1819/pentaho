# pentaho
Dieses Projekt beinhaltet alle wichtigen Dateien, die bzgl. der Pentaho Nachforschungen erstellt wurden.
# office_open_xml_examples

### Ziel dieses Projekts
Mit Pentaho Report Designer ist es nicht möglich, docx-Berichte zu erstellen. Dieses Projekt war ein Versuch, Berichte im prpt-Format einzulesen und mithilfe der Apache-POI-Bibliothek programmatisch ein docx-Dokument zu erstellen.

### Funktionsweise dieses Java-Projekts
Die Klasse mit der main-Methode ist `de.fu_berlin.softwareprojekt1819.openxmlcreation.WordDocument2`.
Diese Klasse liest aus dem Ordner `template` und der Datei `template.prpt` die im Archiv enthaltenen Dateien und erstellt dann die docx-Datei.

### Hinweise zur Verbesserung
Mithilfe des Pentaho-Report-Designer-SDKs kann man die Berichtsdefintion von prpt-Dateien einlesen und darauf zugreifen. Aus diesen Informationen wird die docx-Datei erstellt. Das ist sehr wahrscheinlich eine bessere Vorgehensweise, als auf die Dateien des prpt-Archivs zuzugreifen und sich mit dem XML-Schema auseinanderzusetzen.

### Bekannte Probleme
prpt-Dateien sind normale Archive, die entpackt werden können. Momentan muss sowohl die Datei als auch das entpackte Archiv vorhanden sein, damit dieses Programm funktioniert. Es klappt noch nicht, das XML aus der ZIP-Datei zu parsen. (Es werden Fehler geworfen.) Und ich habs noch nicht angepasst, dass man nur den Ordner braucht und nicht die Datei.

### Weiterführende Links
Pentaho-Doku: https://help.pentaho.com/Documentation/8.2/Developer_Center/Embed_Reporting/In_Java_Application

# prsdk_examples
Some Pentaho Reporting Engine SDK Examples

Using the Pentaho Reporting Engine SDK to enable and create extensions and functions.

### Dependencies

To be able to use the Pentaho Report Engine SDK:

```
		<dependency>
			<groupId>org.pentaho.reporting.engine</groupId>
			<artifactId>classic-core</artifactId>
			<version>8.2.0.0-SNAPSHOT</version>
			<scope>compile</scope>
		</dependency>
```
To be able to use the Pentaho Report Engine SDK Samples:

```
		<dependency>
			<groupId>org.pentaho.reporting.engine</groupId>
			<artifactId>classic-samples</artifactId>
			<version>8.2.0.0-SNAPSHOT</version>
		</dependency>
```

### First Example: Cronjob

For cronjob we're using Spring:

```
		<dependency>
			<groupId>org.springframework.boot</groupId>
			<artifactId>spring-boot-starter</artifactId>
		</dependency>
```
To enable the cronjob we have to annotate the class which contains the main method

```
  @SpringBootApplication
  @EnableScheduling
```
and we have to annotate the method which sould be scheduled

```
@Scheduled(fixedRate = 10000)
```
the arguement "fixedRate" defines the intervals on which the method will called.

At the method "generate" in the "SampleReportGenerator" class, we define the file name and the output type.
At the method "getReportDefinition", we define the masterreport file (prpt).

I got errors caused by springs logging factory. To fix this:

```
		<dependency>
			<groupId>org.springframework.boot</groupId>
			<artifactId>spring-boot-starter</artifactId>
			<exclusions>
				<exclusion>
					<groupId>org.springframework.boot</groupId>
					<artifactId>spring-boot-starter-logging</artifactId>
				</exclusion>
			</exclusions>
		</dependency>
```

```
		<dependency>
			<groupId>org.springframework.boot</groupId>
			<artifactId>spring-boot-starter-log4j</artifactId>
			<version>1.3.8.RELEASE</version>
		</dependency>
```

# prd_useful_files
Im Ordner Beispielberichte sind Beispielberichte mit Charts und Crosstabs zu finden sowie eine Version des von Projektron zur Verfügung gestellten Beispielberichtes, der mithilfe des PRD erstellt wurde.

Im Ordner Buchungsdaten ist die kettle-Datei zu finden, mit welcher die Verbindung zum BCS hergestellt wird. Im mylyn-requests-Ordner sind die Kettle-Dateien für die anderen Requests zu finden (Scrum und Mylyn-Tickets).

Im html-to-pptx ist der Quellcode zu finden, mit dem wir versucht hatten, einen Converter von HTML-Files zu PPTX-Files zu erstellen.
