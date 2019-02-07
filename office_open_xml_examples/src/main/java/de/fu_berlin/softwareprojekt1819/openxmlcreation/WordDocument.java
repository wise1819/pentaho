package de.fu_berlin.softwareprojekt1819.openxmlcreation;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.invoke.MethodHandles;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import javax.imageio.ImageIO;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

public class WordDocument {

	/**
	 * aus {@link https://stackoverflow.com/questions/45559120/how-can-i-set-custom-background-color-on-a-xwpftablecell}
	 * @param table
	 * @param row
	 * @param fromCol
	 * @param toCol
	 */
	private static void mergeCellsHorizontally(XWPFTable table, int row, int fromCol, int toCol) {
		for (int cellIndex = fromCol; cellIndex <= toCol; cellIndex++) {
			XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
			if (cellIndex == fromCol) {
				cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
			} else {
				cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
			}
		}
	}

	/**
	 * aus {@link https://stackoverflow.com/questions/45559120/how-can-i-set-custom-background-color-on-a-xwpftablecell}
	 * @param table
	 * @param col
	 * @param fromRow
	 * @param toRow
	 */
	private static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
		for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
			XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
			if (rowIndex == fromRow) {
				cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
			} else {
				cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
			}
		}
	}

	public void generateOfficeOpnXmlFile(String filename, XWPFDocument document) throws IOException {
		FileOutputStream out = new FileOutputStream(filename);
		document.write(out);
		out.close();
		document.close();
	}

	public void createWordDocument() {
		String output = "testofficeopenxml3.docx";

		XWPFDocument document = new XWPFDocument();

		// Logo einfügen
		XWPFParagraph logo = document.createParagraph();
		logo.setAlignment(ParagraphAlignment.RIGHT);
		XWPFRun logoRun = logo.createRun();
		try {
			URL url = ClassLoader.getSystemResource("Logo_Standardgroesse_RGB2.png");
			BufferedImage bi = ImageIO.read(url);
			int width = bi.getWidth();
			int height = bi.getHeight();
			String filename = url.getFile();
			logoRun.addPicture(new FileInputStream(filename), XWPFDocument.PICTURE_TYPE_PNG, filename,
					Units.toEMU(width), Units.toEMU(height));
		} catch (IOException | InvalidFormatException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		}

		// Tabelle einfügen
		int cols = 4;
		int rows = 10;
		XWPFTable table = document.createTable();
		table.setWidth("100%");
		XWPFTableRow firstRow = table.getRow(0);
		XWPFTableCell cell00 = firstRow.getCell(0);
		XWPFParagraph cell00paragraph = cell00.getParagraphs().get(0);
		XWPFRun cell00run = cell00paragraph.createRun();
		cell00paragraph.setAlignment(ParagraphAlignment.CENTER);
		cell00run.setBold(true);
		cell00run.setText("Nachweis (gebuchter Aufwand)");
		cell00.getCTTc().addNewTcPr().addNewShd().setFill("b4c6e7");
		XWPFTableRow secondRow = table.createRow();
		XWPFTableCell cell10 = secondRow.getCell(0);
		XWPFTableCell cell11 = secondRow.createCell();
		XWPFTableCell cell12 = secondRow.createCell();
		XWPFTableCell cell13 = secondRow.createCell();
		table.createRow();
		table.createRow().createCell().getParagraphs().get(0).addRun(cell00run);

//		XWPFTableCell cell00 = table.getRow(0).getCell(0);

		String current;
		try {
			current = new java.io.File(".").getCanonicalPath();
			System.out.println("Current dir: " + current);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

		try {
			generateOfficeOpnXmlFile(output, document);
		} catch (IOException e) {
			System.err.print("Fehler beim Erstellen der Datei");
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		System.out.println("Hello World!");
		WordDocument a = new WordDocument();
		a.createWordDocument();
	}
}
