package de.fu_berlin.softwareprojekt1819.openxmlcreation;

import java.math.BigInteger;

import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPicture;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTxbxContent;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STLineSpacingRule;
import org.w3c.dom.Document;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import com.microsoft.schemas.vml.CTGroup;
import com.microsoft.schemas.vml.CTShape;

public class Styles {

	private XWPFDocument document;

	public Styles(XWPFDocument doc) {
		document = doc;
	}

	private void readElementStyle(Node node, CTShape ctShape, XWPFParagraph paragraph, XWPFRun run) {
		NodeList childNodes = node.getChildNodes();

		// absolute Position muss sein, damit Abstand der Textbox nach links und oben gesetzt werden kann
		String style = "position:absolute";

		for (int i = 0; i < childNodes.getLength(); i++) {
			Node childNode = childNodes.item(i);
			NamedNodeMap attributes = childNode.getAttributes();
			Node n;
			switch (childNode.getNodeName()) {
			case "border-styles":
				break;
			case "common-styles":
				// vertikale Ausrichtung
				n = attributes.getNamedItem("vertical-alignment");
				if (n != null) {
					if (style.length() > 0)
						style += ";";
					style += "v-text-anchor:" + n.getTextContent();
				}
				// horizontale Ausrichtung
				n = attributes.getNamedItem("alignment");
				if (n != null) {
					String alignment = n.getTextContent();
					switch (alignment) {
					case "center":
						paragraph.setAlignment(ParagraphAlignment.CENTER);
						break;
					// TODO links- und rechtsbündig noch hinzufügen
					default:
						break;
					}

				}
				break;
			case "spatial-styles":
				// Breite
				n = attributes.getNamedItem("min-width");
				if (n != null) {
					if (style.length() > 0)
						style += ";";
					style += "width:" + n.getTextContent() + "pt";
				}
				// Höhe
				n = attributes.getNamedItem("min-height");
				if (n != null) {
					if (style.length() > 0)
						style += ";";
					style += "height:" + n.getTextContent() + "pt";
				}
				// linker Abstand
				n = attributes.getNamedItem("x");
				if (n != null) {
					if (style.length() > 0)
						style += ";";
					style += "margin-left:" + n.getTextContent() + "pt";
				}
				// oberer Abstand
				n = attributes.getNamedItem("y");
				if (n != null) {
					if (style.length() > 0)
						style += ";";
					style += "margin-top:" + n.getTextContent() + "pt";
				}


				break;
			case "text-styles":

				run.setFontFamily(attributes.getNamedItem("font-face").getTextContent());
				run.setFontSize(Integer.parseInt(attributes.getNamedItem("font-size").getTextContent()));
				run.setBold(Boolean.parseBoolean(attributes.getNamedItem("bold").getTextContent()));
				break;
			}
		}
		ctShape.setStyle(style);
	}

	private void readLayoutResourceLabel(Node node, CTShape ctShape, XWPFParagraph paragraph, XWPFRun run) {
		NamedNodeMap attributes = node.getAttributes();
		Node coreValue = attributes.getNamedItem("core:value");

		run.setText(coreValue.getTextContent());

		NodeList nl = node.getChildNodes();
		for (int i = 0; i < nl.getLength(); i++) {
			Node childNode = nl.item(i);
			switch (childNode.getNodeName()) {
			case "element-style":
				readElementStyle(childNode, ctShape, paragraph, run);

				break;
			}
		}

	}

	public void createPageHeader() throws XmlException {
		XWPFHeader header = document.createHeader(HeaderFooterType.DEFAULT);

		// header's first paragraph
		XWPFParagraph paragraph = header.createParagraph();
		XWPFRun run = paragraph.createRun();

		// TODO Dieser Teil mit dem Erstellen der tGroup und der textboxrun muss in
		// den switch-case-Teil verschoben werden, da ja mehrere Textboxen in
		// der Kopfzeile sein können.
		// create inline text box in run
		CTGroup ctGroup = CTGroup.Factory.newInstance();

		CTShape ctShape = ctGroup.addNewShape();
		CTTxbxContent ctTxbxContent = ctShape.addNewTextbox().addNewTxbxContent();

		XWPFParagraph textboxparagraph = new XWPFParagraph(ctTxbxContent.addNewP(), (IBody) header);
		setSingleLineSpacing(textboxparagraph);
		XWPFRun textboxrun = textboxparagraph.createRun();

//		paragraph = header.createParagraph();

		Document xmlDoc = WordDocument2.XML.get(PrptContent.STYLES);

		Node style = xmlDoc.getChildNodes().item(0);
		NodeList nodeList = style.getChildNodes();
		for (int i = 0; i < nodeList.getLength(); i++) {
			Node node = nodeList.item(i);

			switch (node.getNodeName()) {
			case "layout:page-header":

				NodeList nl2 = node.getChildNodes();
				for (int j = 0; j < nl2.getLength(); j++) {
					Node n2 = nl2.item(j);

					switch (n2.getNodeName()) {
					case "layout:resource-label":
						readLayoutResourceLabel(n2, ctShape, textboxparagraph, textboxrun);
						break;
					}
				}

				break;
			}
		}

		// TODO das muss auch in den Switch-Case-Teil
		// ich bin mir aber nicht sicher, ob manb für jede Textbox wieder ein neues
		// CTGroup braucht
		// oder man in einer immer alle Textboxen hinzufügen kann.
		Node ctGroupNode = ctGroup.getDomNode();
		CTPicture ctPicture = CTPicture.Factory.parse(ctGroupNode);
		CTR cTR = run.getCTR();
		cTR.addNewPict();
		cTR.setPictArray(0, ctPicture);

		System.out.println("Länge: " + nodeList.getLength());
		for (int i = 0; i < nodeList.getLength(); i++) {
			System.out.println(nodeList.item(i).getNodeName());
		}
	}

	/**
	 * Setzt Zeilenabstand auf "einfach"<br>
	 * <br>
	 * Danke {@link https://stackoverflow.com/a/27752260}
	 * 
	 * @param para
	 */
	private void setSingleLineSpacing(XWPFParagraph para) {
		CTPPr ppr = para.getCTP().getPPr();
		if (ppr == null)
			ppr = para.getCTP().addNewPPr();
		CTSpacing spacing = ppr.isSetSpacing() ? ppr.getSpacing() : ppr.addNewSpacing();
		spacing.setAfter(BigInteger.valueOf(0));
		spacing.setBefore(BigInteger.valueOf(0));
		spacing.setLineRule(STLineSpacingRule.AUTO);
		spacing.setLine(BigInteger.valueOf(240));
	}

}
