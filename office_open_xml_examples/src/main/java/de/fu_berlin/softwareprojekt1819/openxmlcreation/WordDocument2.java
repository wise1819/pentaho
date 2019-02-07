package de.fu_berlin.softwareprojekt1819.openxmlcreation;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.StringWriter;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

public class WordDocument2 {

	private static final Map<String, String> FILES = new HashMap<>();
	public static final Map<String, Document> XML = new HashMap<>();
	private static String PRPT_DIRECTORY;

	private XWPFDocument document;

	private static Document createXmlDocument(InputStream s)
			throws ParserConfigurationException, SAXException, IOException {

		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
		Document doc = dBuilder.parse(s);
		doc.getDocumentElement().normalize();
		return doc;
	}

	private static Document createXmlDocument2(String s)
			throws ParserConfigurationException, SAXException, IOException {

		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
		Document doc = dBuilder.parse(s);
		doc.getDocumentElement().normalize();
		return doc;
	}

	private static void readReportFile(String path) throws IOException, ParserConfigurationException, SAXException {

		ZipFile zipFile = new ZipFile(path);
		PRPT_DIRECTORY = path.replace(".prpt", "");

		Enumeration<? extends ZipEntry> entries = zipFile.entries();

		while (entries.hasMoreElements()) {
			ZipEntry entry = entries.nextElement();
			InputStream stream = zipFile.getInputStream(entry);

			System.out.println("Name: " + entry.getName());

//			XML.put(entry.getName(), createXmlDocument(stream));

			try (BufferedReader br = new BufferedReader(new InputStreamReader(stream, "UTF-8"))) {
				FILES.put(entry.getName(), br.lines().collect(Collectors.joining(System.lineSeparator())));
			}
			if (entry.getName().endsWith(".xml")) {
				// TODO createXmlDocument2 erwartet den Pfad der Datei. Das heißt, man muss die prpt-Datei entpacken.
				// Man kann auch ausm ZIP-Archiv die Dateien einlesen, aber dann gibts Probleme beim Parsen des XMLs.
				XML.put(entry.getName(), createXmlDocument2(PRPT_DIRECTORY + File.separator + entry.getName()));
			}
			
			stream.close();
		}
	}

	private void generateOfficeOpnXmlFile(String filename, XWPFDocument document) throws IOException {
		FileOutputStream out = new FileOutputStream(filename);
		document.write(out);
		out.close();
		document.close();
	}

	private void createWordDocument() throws XmlException {
		String output = "test554_0.docx";

		document = new XWPFDocument();
		Styles styles = new Styles(document);

		styles.createPageHeader();

		try {
			generateOfficeOpnXmlFile(output, document);
		} catch (IOException e) {
			System.err.print("Fehler beim Erstellen der Datei");
			e.printStackTrace();
		}

	}



	public static void main(String[] args)
			throws IOException, ParserConfigurationException, SAXException, TransformerException, XmlException {
		
		readReportFile("template.prpt");

		WordDocument2 a = new WordDocument2();
		a.createWordDocument();

		// System.out.println(FILES.get(PrptContent.LAYOUT).toString());

		Document doc = XML.get(PrptContent.LAYOUT);
//		System.out.println(doc.hasAttributes());

//		System.out.println(doc.toString());

//		// Hole Root-Element
//		Element root = null;
//
//	    NodeList list = doc.getChildNodes();
//	    for (int i = 0; i < list.getLength(); i++) {
//	    	System.out.println("i: " + i);
//	      if (list.item(i) instanceof Element) {
//	        root = (Element) list.item(i);
//	        break;
//	      }
//	    }
//	    root = doc.getDocumentElement();
//		System.out.println(root.toString());

//		// Ausgabe XML
//		TransformerFactory tf = TransformerFactory.newInstance();
//		Transformer transformer = tf.newTransformer();
//		transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes");
//		StringWriter writer = new StringWriter();
//		transformer.transform(new DOMSource(doc), new StreamResult(writer));
//		String output = writer.getBuffer().toString().replaceAll("\n|\r", "");
//	
//		System.out.println("Ausgabe: " + output);
	}
}
