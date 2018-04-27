package com.eastrobot.poi;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
import java.util.List;

/**
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-04-26 16:17
 */
public class POITest3 {
    public static void main(String[] args) {
        try {
            InputStream is = new FileInputStream("C:\\Users\\User\\Desktop\\kbase-media-2016.docx");
            XWPFDocument doc = new XWPFDocument(is);

            List<XWPFParagraph> paras = doc.getParagraphs();

            XWPFDocument newdoc = new XWPFDocument();
            for (XWPFParagraph para : paras) {
                if (!para.getParagraphText().isEmpty()) {
                    XWPFParagraph newpara = newdoc.createParagraph();
                    copyAllRunsToAnotherParagraph(para, newpara);
                }
            }

            FileOutputStream fos = new FileOutputStream(new File("C:\\Users\\User\\Desktop\\test.docx"));
            newdoc.write(fos);
            fos.flush();
            fos.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Copy all runs from one paragraph to another, keeping the style unchanged
    private static void copyAllRunsToAnotherParagraph(XWPFParagraph oldPar, XWPFParagraph newPar) {
        final int DEFAULT_FONT_SIZE = 10;

        for (XWPFRun run : oldPar.getRuns()) {
            String textInRun = run.getText(0);
            if (textInRun == null || textInRun.isEmpty()) {
                continue;
            }

            int fontSize = run.getFontSize();
            System.out.println("run text = '" + textInRun + "' , fontSize = " + fontSize);

            XWPFRun newRun = newPar.createRun();

            // Copy text
            newRun.setText(textInRun);

            // Apply the same style
            newRun.setFontSize((fontSize == -1) ? DEFAULT_FONT_SIZE : run.getFontSize());
            newRun.setFontFamily(run.getFontFamily());
            newRun.setBold(run.isBold());
            newRun.setItalic(run.isItalic());
            newRun.setStrike(run.isStrike());
            newRun.setColor(run.getColor());
        }
    }
}
