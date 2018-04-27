package com.eastrobot.poi;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.util.Iterator;
import java.util.List;

/**
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-04-26 15:58
 */
public class POItest2 {
    public static void main(String[] args) throws Exception {
        XWPFDocument doc = new XWPFDocument(new FileInputStream("C:\\Users\\User\\Desktop\\kbase-media-2016.docx"));
        // 获取bodyElements
        List<IBodyElement> bodyElements = doc.getBodyElements();
        Iterator<XWPFParagraph> paragraphsIterator = doc.getParagraphsIterator();


        for (int i = 0; i < bodyElements.size(); i++) {
            IBodyElement e = bodyElements.get(i);
            if (e.getElementType().equals(BodyElementType.PARAGRAPH)) {
                System.out.println(((XWPFParagraph) e).getText());
            } else if (e.getElementType().equals(BodyElementType.TABLE)) {
                XWPFTable table = (XWPFTable) e;
                System.out.println(table.getText());
            }
            // XWPFParagraph paragraph = newDoc.createParagraph();
            // copyParagraph(paragraph, paras.get(i));


            // System.out.println("remove>> " + ((XWPFParagraph) bodyElements.get(i)).getText());
            // doc.removeBodyElement(i);
        }

    }
}
