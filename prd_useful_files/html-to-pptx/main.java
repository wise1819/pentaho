import org.apache.poi.sl.usermodel.TableCell;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.*;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import java.util.List;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {


    public static void main(String args[]) throws IOException {

        //Lese HTML Output von Pentaho
        File html = new File("HTML Report/report.html");
        Document doc = Jsoup.parse(html, "UTF-8", "Report");

        final Elements rows = doc.body().getAllElements();

        for (Element row : rows) {

            if (row.text().length() != 0) System.out.println(row.className() + row.text());

        }

        //creating a new empty slide show
        XMLSlideShow ppt = new XMLSlideShow();

        //creating an FileOutputStream object
        File file = new File("report.pptx");
        FileOutputStream out = new FileOutputStream(file);

        // Hole SlideLayouts
        XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);
        // title slide
        XSLFSlideLayout titleLayout = defaultMaster.getLayout(SlideLayout.TITLE);


        //Slide 1: Titel
        XSLFSlide slide1 = ppt.createSlide(titleLayout);
        XSLFTextShape title1 = slide1.getPlaceholder(0);
        title1.setText(doc.select("html > body > table.style-0 > tbody > tr:nth-child(2) > td.style-3").text());
        System.out.println(title1.getText());

        XSLFTextShape body1 = slide1.getPlaceholder(1);
        body1.clearText(); // unset any existing text
        body1.setText(doc.select("html > body > table.style-0 > tbody > tr:nth-child(3)").text());


        //Slide 2-N:

        List<String> produktlinie = doc.body().getElementsByClass("style-7").eachText(); // Bezeichnung der Produktlinie
        List<String> produktlinie_value = doc.body().getElementsByClass("style-8").eachText(); // Produktlinie XY Wert
        List<String> tabellenspalte1 = doc.body().getElementsByClass("style-11").eachText();
        List<String> tabellenspalte2 = doc.body().getElementsByClass("style-12").eachText();
        List<String> tabellenspalte3 = doc.body().getElementsByClass("style-13").eachText();
        List<String> tabellenspalte4 = doc.body().getElementsByClass("style-14").eachText();
        List<String> tabellenspalte1_inhalt = doc.body().getElementsByClass("style-16").eachText();
        List<String> tabellenspalte2_inhalt = doc.body().getElementsByClass("style-1").eachText();
        List<String> tabellenspalte3_inhalt = doc.body().getElementsByClass("style-17").eachText();
        List<String> tabellenspalte4_inhalt = doc.body().getElementsByClass("style-18").eachText();


        // tabellen slides
        //  XSLFSlideLayout tableLayout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);

        //produktlinie.text();

        while(tabellenspalte1.size()>0) {
            // Folie für Titel und Bilder anlegen
            XSLFSlide slide0 = ppt.createSlide(defaultMaster.getLayout(SlideLayout.TITLE_ONLY));

            //Titel
            XSLFTextShape title = slide0.getPlaceholder(0);
            title.setText(produktlinie.get(0).concat("\n").concat(produktlinie_value.get(0)));
            produktlinie.remove(0);
            produktlinie_value.remove(0);

            //reading Statistik
            File image = new File("HTML Report/picture371389636.png");

            //converting it into a byte array
            byte[] picture = IOUtils.toByteArray(new FileInputStream(image));

            //adding the image to the presentation
            XSLFPictureData idx = ppt.addPicture(picture, XSLFPictureData.PictureType.PNG);

            //creating a slide with given picture on it
            XSLFPictureShape pic = slide0.createPicture(idx);
            pic.setAnchor(new Rectangle(210, 200, 309, 261));

            //Folie für Tabelle anlegen
            XSLFSlide slide = ppt.createSlide();

            XSLFTable tbl = slide.createTable();
            tbl.setAnchor(new Rectangle(50, 50, 450, 300));
            XSLFTableRow headerRow = tbl.addRow();
            headerRow.setHeight(50);
            XSLFTableCell th = headerRow.addCell();
            XSLFTextParagraph p = th.addNewTextParagraph();
            p.setTextAlign(TextParagraph.TextAlign.CENTER);
            XSLFTextRun r = p.addNewTextRun();
            r.setText(tabellenspalte1.get(0));
            tabellenspalte1.remove(0);
            r.setBold(true);
            r.setFontColor(Color.white);
            th.setFillColor(new Color(79, 129, 189));
            th.setBorderWidth(TableCell.BorderEdge.bottom, 2.0);


            XSLFTableCell th2 = headerRow.addCell();
            XSLFTextParagraph p2 = th2.addNewTextParagraph();
            p2.setTextAlign(TextParagraph.TextAlign.CENTER);
            XSLFTextRun r2 = p2.addNewTextRun();
            r2.setText(tabellenspalte2.get(0));
            tabellenspalte2.remove(0);
            r2.setBold(true);
            r2.setFontColor(Color.white);
            th2.setFillColor(new Color(79, 129, 189));
            th2.setBorderWidth(TableCell.BorderEdge.bottom, 2.0);

            XSLFTableCell th3 = headerRow.addCell();
            XSLFTextParagraph p3 = th3.addNewTextParagraph();
            p3.setTextAlign(TextParagraph.TextAlign.CENTER);
            XSLFTextRun r3 = p3.addNewTextRun();
            r3.setText(tabellenspalte3.get(0));
            tabellenspalte3.remove(0);
            r3.setBold(true);
            r3.setFontColor(Color.white);
            th3.setFillColor(new Color(79, 129, 189));
            th3.setBorderWidth(TableCell.BorderEdge.bottom, 2.0);

            XSLFTableCell th4 = headerRow.addCell();
            XSLFTextParagraph p4 = th4.addNewTextParagraph();
            p4.setTextAlign(TextParagraph.TextAlign.CENTER);
            XSLFTextRun r4 = p4.addNewTextRun();
            r4.setText(tabellenspalte4.get(0));
            tabellenspalte4.remove(0);
            r4.setBold(true);
            r4.setFontColor(Color.white);
            th4.setFillColor(new Color(79, 129, 189));
            th4.setBorderWidth(TableCell.BorderEdge.bottom, 2.0);


            tbl.setColumnWidth(0, 150);
            tbl.setColumnWidth(1,150);
            tbl.setColumnWidth(2,150);
            tbl.setColumnWidth(3,150);



            // Festgelegte Zeilenanzahl kann hinzugefügt werden, in diesem Fall 8 pro Tabelle

            for (int i = 0; i < 8; i++){
            XSLFTableRow tr = tbl.addRow();
            tr.setHeight(50);
            // eine Zeile komplett:

            //Zeile 1
            XSLFTableCell cell = tr.addCell();
            XSLFTextParagraph p5 = cell.addNewTextParagraph();
            XSLFTextRun r5 = p5.addNewTextRun();

            r5.setText(tabellenspalte1_inhalt.get(0));
            tabellenspalte1_inhalt.remove(0);

            //Zeile 2
            XSLFTableCell cell6 = tr.addCell();
            XSLFTextParagraph p6 = cell6.addNewTextParagraph();
            XSLFTextRun r6 = p6.addNewTextRun();

            r6.setText(tabellenspalte2_inhalt.get(0));
            tabellenspalte2_inhalt.remove(0);

            XSLFTableCell cell7 = tr.addCell();
            XSLFTextParagraph p7 = cell7.addNewTextParagraph();
            XSLFTextRun r7 = p7.addNewTextRun();

            r7.setText(tabellenspalte3_inhalt.get(0));
            tabellenspalte3_inhalt.remove(0);


            XSLFTableCell cell8 = tr.addCell();
            XSLFTextParagraph p8 = cell8.addNewTextParagraph();
            XSLFTextRun r8 = p8.addNewTextRun();

            r8.setText(tabellenspalte4_inhalt.get(0));
            tabellenspalte4_inhalt.remove(0);

            }

        }

            // Ausgabe speichern/schreiben
            ppt.write(out);
            System.out.println("Presentation edited successfully");
            out.close();




            }



            }





