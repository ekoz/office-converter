package com.eastrobot.poi;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.FileInputStream;
import java.io.StringWriter;

/**
 * POIUtil
 *
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-04-25 16:49
 */
public class POIUtil {

    public static void wordToHtml(String filePath) {
        String extension = FilenameUtils.getExtension(filePath);
        if ("doc".equals(extension)) {
            docToHtml(filePath);
        } else if ("docx".equals(extension)) {
            docxToHtml(filePath);
        }
    }

    public static void docToHtml(String filePath) {}/*{
        File file = new File(filePath);
        String str = "";
        try {
            FileInputStream fis = new FileInputStream(file);
            HWPFDocument doc = new HWPFDocument(fis);
            String doc1 = doc.getDocumentText();
            System.out.println(doc1);

            StringBuilder doc2 = doc.getText();
            System.out.println(doc2);

            Range rang = doc.getRange();
            String doc3 = rang.text();
            System.out.println(doc3);

            int pages = doc.getSummaryInformation().getPageCount();//总页数
            int wordCount = doc.getSummaryInformation().getWordCount();//总字符数
            System.out.println("pages=" + pages + " wordCount=" + wordCount);

            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }*/

    public static void docxToHtml(String filePath) {
        File file = new File(filePath);
        String str = "";
        try {
            FileInputStream fis = new FileInputStream(file);
            XWPFDocument docx = new XWPFDocument(fis);
            XWPFWordExtractor extractor = new XWPFWordExtractor(docx);
            // String doc1 = extractor.getText();
            // System.out.println(doc1);

            int pages = docx.getProperties().getExtendedProperties().getUnderlyingProperties().getPages();//总页数
            // 忽略空格的总字符数 另外还有getCharactersWithSpaces()方法获取带空格的总字数。
            int wordCount = docx.getProperties().getExtendedProperties().getUnderlyingProperties().getCharacters();
            System.out.println("pages=" + pages + " wordCount=" + wordCount);

            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * doc
     */
    public static void wordConverter(HWPFDocumentCore hwpfDocumentCore) throws Exception {
        Document newDocument = DocumentBuilderFactory.newInstance()
                .newDocumentBuilder().newDocument();
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                newDocument);

        wordToHtmlConverter.processDocument(hwpfDocumentCore);

        StringWriter stringWriter = new StringWriter();
        Transformer transformer = TransformerFactory.newInstance()
                .newTransformer();
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        transformer.setOutputProperty(OutputKeys.METHOD, "html");
        transformer.transform(
                new DOMSource(wordToHtmlConverter.getDocument()),
                new StreamResult(stringWriter));

        String html = stringWriter.toString();
    }


    public static void main(String[] args) throws Exception {
        // StopWatch stopWatch = new StopWatch();
        // stopWatch.start();
        String docxStr = "C:\\Users\\User\\Desktop\\kbase-media-2016.docx";
        String docStr = "C:\\Users\\User\\Desktop\\系统操作模拟文档2003.doc";
        // String docStr = "C:\\Users\\User\\Desktop\\kbase-media-2003.doc";
        // Word2Html.wordToHtml(docStr);
        FileInputStream in = new FileInputStream(new File(docStr));
        HWPFDocument doc = new HWPFDocument(in);
        // Range range = doc.getRange();// 得到文档的读取范围
        // int pages = doc.getSummaryInformation().getPageCount();//总页数
        //
        // // 遍历段落
        // int numParagraphs = range.numParagraphs();
        // for (int i = 0; i < numParagraphs; i++) {
        //     System.out.println(i + "    >>>>>   " + range.getParagraph(i).text());
        // }
        //
        // TableIterator tableIterator = new TableIterator(range);
        // if (tableIterator.hasNext()) {
        //     Table tb = tableIterator.next();
        //     for (int i = 0; i < tb.numRows(); i++) {
        //         TableRow row = tb.getRow(i);
        //         for (int j = 0; j < row.numCells(); j++) {
        //             TableCell cell = row.getCell(j);
        //             for (int k = 0; k < cell.numParagraphs(); k++) {
        //                 Paragraph paragraph = cell.getParagraph(k);
        //                 System.out.println("i=" + i + ",j=" + j + ",k=" + k + " >> " + paragraph.text());
        //             }
        //         }
        //     }
        // }

        File outputFile = new File("C:\\Users\\User\\Desktop\\out.doc");


        XWPFDocument document=new XWPFDocument();
        XWPFParagraph paragraph = document.createParagraph();


        // System.out.println(hwpf.getDocProperties());


        // docToHtml(doc);
        // docxToHtml(docx);
        // FileInputStream fis = new FileInputStream(file);
        // XWPFDocument doc = new XWPFDocument(fis);
        // stopWatch.stop();
        // System.out.println(stopWatch.prettyPrint());
    }


}
