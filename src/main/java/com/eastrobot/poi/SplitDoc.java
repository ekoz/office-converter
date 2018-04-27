package com.eastrobot.poi;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-04-26 16:11
 */
public class SplitDoc {
    public static void main(String[] args) {

        FileInputStream in = null;
        HWPFDocument doc = null;

        XWPFDocument us = null;
        XWPFDocument japan = null;
        FileOutputStream outUs = null;
        FileOutputStream outJapan = null;

        try {
            in = new FileInputStream("C:\\Users\\User\\Desktop\\系统操作模拟文档2003.doc");
            doc = new HWPFDocument(in);

            us = new XWPFDocument();
            japan = new XWPFDocument();

            Range range = doc.getRange();

            for (int parIndex = 0; parIndex < range.numParagraphs(); parIndex++) {
                Paragraph paragraph = range.getParagraph(parIndex);

                String text = paragraph.text();
                System.out.println("***Paragraph" + parIndex + ": " + text);

                if ((parIndex >= 11) && (parIndex <= 20)) {
                    createParagraphInAnotherDocument(us, text);
                } else if ((parIndex >= 21) && (parIndex <= 30)) {
                    createParagraphInAnotherDocument(japan, text);
                }
            }

            outUs = new FileOutputStream("C:\\Users\\User\\Desktop\\us.docx");
            outJapan = new FileOutputStream("C:\\Users\\User\\Desktop\\japan.docx");
            us.write(outUs);
            japan.write(outJapan);

            in.close();
            outUs.close();
            outJapan.close();

        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private static void createParagraphInAnotherDocument(XWPFDocument document, String text) {
        XWPFParagraph newPar = document.createParagraph();
        newPar.createRun().setText(text, 0);
    }
}
