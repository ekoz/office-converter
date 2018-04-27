/*
package com.eastrobot.poi;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;

import java.io.*;
import java.math.BigInteger;
import java.util.List;

*/
/**
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-04-26 14:20
 *//*

public class WordFinal {
    public static void main(String[] args) throws Exception {
        InputStream is = new FileInputStream("C:\\Users\\User\\Desktop\\kbase-media-2016.docx");
        XWPFDocument srcDoc = new XWPFDocument(is);
        CustomXWPFDocument destDoc = new CustomXWPFDocument();
        // Copy document layout.
        copyLayout(srcDoc, destDoc);
        OutputStream out = new FileOutputStream("C:\\Users\\User\\Desktop\\1.docx");

        for (IBodyElement bodyElement : srcDoc.getBodyElements()) {
            BodyElementType elementType = bodyElement.getElementType();
            if (elementType == BodyElementType.PARAGRAPH) {

                XWPFParagraph srcPr = (XWPFParagraph) bodyElement;
                copyStyle(srcDoc, destDoc, srcDoc.getStyles().getStyle(srcPr.getStyleID()));
                boolean hasImage = false;
                XWPFParagraph dstPr = destDoc.createParagraph();
                // Extract image from source docx file and insert into destination docx file.
                for (XWPFRun srcRun : srcPr.getRuns()) {

                    // You need next code when you want to call XWPFParagraph.removeRun().
                    dstPr.createRun();

                    if (srcRun.getEmbeddedPictures().size() > 0)
                        hasImage = true;

                    for (XWPFPicture pic : srcRun.getEmbeddedPictures()) {

                        byte[] img = pic.getPictureData().getData();

                        long cx = pic.getCTPicture().getSpPr().getXfrm().getExt().getCx();
                        long cy = pic.getCTPicture().getSpPr().getXfrm().getExt().getCy();

                        try {
                            // Working addPicture Code below...
                            String blipId = dstPr.getDocument().addPictureData(new ByteArrayInputStream(img),
                                    Document.PICTURE_TYPE_PNG);
                            destDoc.createPictureCxCy(blipId, destDoc.getNextPicNameNumber(Document.PICTURE_TYPE_PNG),
                                    cx, cy);

                        } catch (InvalidFormatException e1) {
                            e1.printStackTrace();
                        }
                    }
                }

                if (hasImage == false) {
                    int pos = destDoc.getParagraphs().size() - 1;
                    destDoc.setParagraph(srcPr, pos);
                }

            } else if (elementType == BodyElementType.TABLE) {

                XWPFTable table = (XWPFTable) bodyElement;

                copyStyle(srcDoc, destDoc, srcDoc.getStyles().getStyle(table.getStyleID()));

                destDoc.createTable();

                int pos = destDoc.getTables().size() - 1;

                destDoc.setTable(pos, table);
            }
        }

        destDoc.write(out);
        out.close();
    }

    // Copy Styles of Table and Paragraph.
    private static void copyStyle(XWPFDocument srcDoc, XWPFDocument destDoc, XWPFStyle style) {
        if (destDoc == null || style == null)
            return;

        if (destDoc.getStyles() == null) {
            destDoc.createStyles();
        }

        List<XWPFStyle> usedStyleList = srcDoc.getStyles().getUsedStyleList(style);
        for (XWPFStyle xwpfStyle : usedStyleList) {
            destDoc.getStyles().addStyle(xwpfStyle);
        }
    }

    private static void copyLayout(XWPFDocument srcDoc, XWPFDocument destDoc) {
        CTPageMar pgMar = srcDoc.getDocument().getBody().getSectPr().getPgMar();

        BigInteger bottom = pgMar.getBottom();
        BigInteger footer = pgMar.getFooter();
        BigInteger gutter = pgMar.getGutter();
        BigInteger header = pgMar.getHeader();
        BigInteger left = pgMar.getLeft();
        BigInteger right = pgMar.getRight();
        BigInteger top = pgMar.getTop();

        CTPageMar addNewPgMar = destDoc.getDocument().getBody().addNewSectPr().addNewPgMar();

        addNewPgMar.setBottom(bottom);
        addNewPgMar.setFooter(footer);
        addNewPgMar.setGutter(gutter);
        addNewPgMar.setHeader(header);
        addNewPgMar.setLeft(left);
        addNewPgMar.setRight(right);
        addNewPgMar.setTop(top);

        CTPageSz pgSzSrc = srcDoc.getDocument().getBody().getSectPr().getPgSz();

        BigInteger code = pgSzSrc.getCode();
        BigInteger h = pgSzSrc.getH();
        STPageOrientation.Enum orient = pgSzSrc.getOrient();
        BigInteger w = pgSzSrc.getW();

        CTPageSz addNewPgSz = destDoc.getDocument().getBody().addNewSectPr().addNewPgSz();

        addNewPgSz.setCode(code);
        addNewPgSz.setH(h);
        addNewPgSz.setOrient(orient);
        addNewPgSz.setW(w);
    }
}
*/
