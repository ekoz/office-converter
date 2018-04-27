/*
package com.eastrobot.poi;

import com.sun.corba.se.impl.interceptors.PICurrent;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

*/
/**
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-04-26 12:41
 *//*

public class POItest1 {

    public static void main(String[] args) throws Exception {
        split2();
    }

    public static void CopyRun(XWPFRun target, XWPFRun source) {
        target.getCTR().setRPr(source.getCTR().getRPr());
        // 设置文本
        target.setText(source.text());
    }

    public static void copyParagraph(XWPFParagraph target, XWPFParagraph source) {
        // 设置段落样式
        target.getCTP().setPPr(source.getCTP().getPPr());
        // 添加Run标签
        for (int pos = 0; pos < target.getRuns().size(); pos++) {
            target.removeRun(pos);
        }
        for (XWPFRun s : source.getRuns()) {
            XWPFRun targetrun = target.createRun();
            CopyRun(targetrun, s);
        }
    }

    public static void copyTableCell(XWPFTableCell target, XWPFTableCell source) {
        // 列属性
        target.getCTTc().setTcPr(source.getCTTc().getTcPr());
        // 删除目标 targetCell 所有单元格
        for (int pos = 0; pos < target.getParagraphs().size(); pos++) {
            target.removeParagraph(pos);
        }
        // 添加段落
        for (XWPFParagraph sp : source.getParagraphs()) {
            XWPFParagraph targetP = target.addParagraph();
            copyParagraph(targetP, sp);
        }
    }

    public static void CopytTableRow(XWPFTableRow target, XWPFTableRow source) {
        // 复制样式
        target.getCtRow().setTrPr(source.getCtRow().getTrPr());
        // 复制单元格
        for (int i = 0; i < target.getTableCells().size(); i++) {
            copyTableCell(target.getCell(i), source.getCell(i));
        }
    }

    */
/**
     * 读写文档中的图片
     *//*

    private static void readPicture(PicturesTable pTable, CharacterRun cr) throws Exception {
        // 提取图片
        Picture pic = pTable.extractPicture(cr, false);
        BufferedImage image = null;// 图片对象
        // 获取图片样式
        int picHeight = pic.getHeight() * pic.getVerticalScalingFactor() / 100;
        int picWidth = pic.getHorizontalScalingFactor() * pic.getWidth() / 100;
        if (picWidth > 500) {
            picHeight = 500 * picHeight / picWidth;
            picWidth = 500;
        }
        String style = " style='height:" + picHeight + "px;width:" + picWidth + "px'";

        // 返回POI建议的图片文件名
        String afileName = pic.suggestFullFileName();
        //单元测试路径
        // String directory = "E:\\converter-html\\images\\" + wordName + "\\";
        //项目路径
        //String directory = tempPath + "images/" + wordName + "/";
        // makeDir(directory);// 创建文件夹

        int picSize = cr.getFontSize();
        int myHeight = 0;

        if (afileName.indexOf(".wmf") > 0) {
            // OutputStream out = new FileOutputStream(new File(directory + afileName));
            out.write(pic.getContent());
            out.close();
            // afileName = Wmf2Png.convertdirectory + afileName);

            // File file = new File(directory + afileName);

            try {
                image = ImageIO.read(file);
            } catch (Exception e) {
                e.printStackTrace();
            }

            int pheight = image.getHeight();
            int pwidth = image.getWidth();
            // if (pwidth > 500) {
            //     htmlText += "<img style='width:" + pwidth + "px;height:" + myHeight + "px'" + " src=\"" + directory
            //             + afileName + "\"/>";
            // } else {
            //     myHeight = (int) (pheight / (pwidth / (picSize * 1.0)) * 1.5);
            //     htmlText += "<img style='vertical-align:middle;width:" + picSize * 1.5 + "px;height:" + myHeight
            //             + "px'" + " src=\"" + directory + afileName + "\"/>";
            // }

        } else

        {
            OutputStream out = new FileOutputStream(new File(directory + afileName));
            // pic.writeImageContent(out);
            out.write(pic.getContent());
            out.close();
            // 处理jpg或其他（即除png外）
            if (afileName.indexOf(".png") == -1) {
                try {
                    File file = new File(directory + afileName);
                    image = ImageIO.read(file);
                    picHeight = image.getHeight();
                    picWidth = image.getWidth();
                    if (picWidth > 500) {
                        picHeight = 500 * picHeight / picWidth;
                        picWidth = 500;
                    }
                    style = " style='height:" + picHeight + "px;width:" + picWidth + "px'";
                } catch (Exception e) {
                    // e.printStackTrace();
                }
            }
            htmlText += "<img " + style + " src=\"" + directory + afileName + "\"/>";
        }
        if (pic.getWidth() > 450) {
            htmlText += "<br/>";
        }
    }


    */
/**
     *   Range：它表示一个范围，这个范围可以是整个文档，也可以是里面的某一小节（Section），
     *   也可以是某一个段落（Paragraph），还可以是拥有共同属性的一段文本（CharacterRun）。
     *   Section：word文档的一个小节，一个word文档可以由多个小节构成。
     *   Paragraph：word文档的一个段落，一个小节可以由多个段落构成。
     *   CharacterRun：具有相同属性的一段文本，一个段落可以由多个CharacterRun组成。
     *   Table：一个表格。
     *   TableRow：表格对应的行。
     *   TableCell：表格对应的单元格。
*        Section、Paragraph、CharacterRun和Table都继承自Range。
     * @throws Exception
     *//*

    public static void split2() throws Exception {

        String path = "C:\\Users\\User\\Desktop\\kbase-media-2003.doc";
        InputStream is = new FileInputStream(path);
        HWPFDocument doc = new HWPFDocument(is);
        Range range = doc.getRange();
        PicturesTable picturesTable = doc.getPicturesTable();
        TableIterator tableIterator = new TableIterator(range);
        int length = doc.characterLength();

        // out

        String templatePath = "C:\\Users\\User\\Desktop\\test.doc";
        InputStream ois = new FileInputStream(templatePath);
        POIFSFileSystem pfs = new POIFSFileSystem(ois);
        HWPFDocument odoc = new HWPFDocument(is);
        Range orange = odoc.getRange();

        List<Table> tableList = new ArrayList<>();
        int[] tableStartOffset = new int[100];
        int[] tableEndOffset = new int[100];
        boolean hasTable = false;

        int tablePos = 0;
        while (tableIterator.hasNext()) {
            Table table = tableIterator.next();
            int startOffset = table.getStartOffset();
            int endOffset = table.getEndOffset();

            tableStartOffset[tablePos] = startOffset;
            tableEndOffset[tablePos] = endOffset;
            tableList.add(table);
            hasTable = true;
            tablePos++;
        }

        tablePos = 0;
        Iterator<Table> tableListIterator = tableList.iterator();
        for (int i = 0; i < length - 1; i++) {
            Range tmpRange = new Range(i, i + 1, doc);
            CharacterRun c1 = range.getCharacterRun(0);

            if (hasTable) {
                if (i == tableStartOffset[tablePos]) {
                    // TODO by Yogurt_lei : writetable to new doc & check limit
                    System.out.println(tablePos + ">>>表格");
                    i = tableEndOffset[tablePos] - 1;
                    tablePos++;
                }
            }else if (picturesTable.hasPicture(c1)) {
                // TODO by Yogurt_lei : writeimage to new doc & check limit
                System.out.println(i + ">>>图片");
                // readimg
            } else {
                //正常文本
                Range range2 = new Range(i + 1, i + 2, doc);
                // 第二个字符
                CharacterRun c2 = range2.getCharacterRun(0);
                // TODO by Yogurt_lei : write c2 to new doc & check limit
            }
        }


        //word 2003： 图片不会被读取
        // InputStream is = new FileInputStream(new File("c://files//2003.doc"));
        // WordExtractor ex = new WordExtractor(is);
        // String text2003 = ex.getText();
        // System.out.println(text2003);

        //word 2007 图片不会被读取， 表格中的数据会被放在字符串的最后
        // OPCPackage opcPackage = POIXMLDocument.openPackage("c://files//2007.docx");
        // POIXMLTextExtractor extractor = new XWPFWordExtractor(opcPackage);
        // String text2007 = extractor.getText();
        // System.out.println(text2007);

        // 获取bodyElements
        // List<IBodyElement> bodyElements = doc.getBodyElements();
        // int mid = bodyElements.size() / 2 - 1;

        // XWPFDocument newDoc = new XWPFDocument();

        // for (int i = 0; i < paras.size(); i++) {
        //     System.out.println(paras.get(i).getText());
        //     XWPFParagraph paragraph = newDoc.createParagraph();
        //     copyParagraph(paragraph, paras.get(i));
        //     paras.
        //
        //
        //     newDoc.setParagraph(paragraph, i);
        //     // System.out.println("remove>> " + ((XWPFParagraph) bodyElements.get(i)).getText());
        //     // doc.removeBodyElement(i);
        // }
        // newDoc.write(new FileOutputStream("C:\\Users\\User\\Desktop\\" + 1 + ".docx"));
        //
        // OutputStream out = new FileOutputStream("C:\\Users\\User\\Desktop\\" + 1 + ".docx");
        // doc.write(out);

        // for (int i = 0; i < bodyElements.size() / 2; i++) {
        //     // doc.removeBodyElement(i);
        //     // if (element.getElementType().equals(BodyElementType.PARAGRAPH)) {
        //     //
        //     // } else if (element.getElementType().equals(BodyElementType.TABLE)) {
        //     //
        //     // }
        //     System.out.println((++i) + " >>> " + bodyElements.get(i));
        //
        //     OutputStream out = new FileOutputStream("C:\\Users\\User\\Desktop\\" + i + ".docx");
        //     doc.write(out);
        // }
    }

    public static void splitWord() throws Exception {
        String path = "C:\\Users\\User\\Desktop\\kbase-media-2016.docx";
        InputStream is = new FileInputStream(path);
        XWPFDocument doc = new XWPFDocument(is);
        // 获取段落
        List<XWPFParagraph> paras = doc.getParagraphs();
        // 获取bodyElements
        List<IBodyElement> bodyElements = doc.getBodyElements();

        // 获取doc样式
        XWPFStyles styles = doc.getStyles();
        int j = 0;
        // /切割成的word 文件存储位置
        String patha = "C:\\Users\\User\\Desktop\\";
        // 根据大纲定义分割成的段落
        ArrayList<Integer> al_duanLuo = new ArrayList<>();
        // 大纲名称
        ArrayList al2_name = new ArrayList<>();
        // 存放生成wordId
        ArrayList<String> al6_wordId = new ArrayList<>();

        for (int i = 0; i < bodyElements.size(); i++) {
            IBodyElement bodyElement = bodyElements.get(i);
            if (j == 0) {
                al_duanLuo.add(i);
                j++;
                al2_name.add("首页");
                al6_wordId.add(java.util.UUID.randomUUID().toString());
            }
            if (bodyElement.getElementType() == BodyElementType.PARAGRAPH) {
                XWPFParagraph para = (XWPFParagraph) bodyElement;

                // 以标题创建第一个文件
                al_duanLuo.add(i);
                j++;
                al2_name.add(para.getParagraphText());
                al6_wordId.add(java.util.UUID.randomUUID().toString());
            }
        }
        // 定义存放父新id
        ArrayList al4_parentId = new ArrayList<>();
        XWPFDocument newDoc = doc;
        int max = bodyElements.size() - 1;
        al_duanLuo.add(max);
        for (int k = 0; k < al_duanLuo.size() - 1; k++) {
            path = "C:\\Users\\User\\Desktop\\kbase-media-2016.docx";
            is = new FileInputStream(path);
            doc = new XWPFDocument(is);
            // 移除多级列表，移除前面的编号，这里分割后是有编号的，不过这里如果你不移除的话，直接把代码注释掉即可
            if (k != 0) {
                XWPFParagraph para1 = (XWPFParagraph) doc.getBodyElements().get(al_duanLuo.get(k));
                String str1 = para1.getStyleID();
            }
            // 移除前0－－14，
            int temp = al_duanLuo.get(k);
            int tempCount = al_duanLuo.get(k + 1);

            for (int u = max; u > tempCount - 1; u--) {
                doc.removeBodyElement(u);
            }
            // 进行移除之前
            for (int l = temp - 1; l >= 0; l--) {
                doc.removeBodyElement(l);
            }
            OutputStream out = new FileOutputStream("C:\\Users\\User\\Desktop\\" + al6_wordId.get(k) + ".docx");
            doc.write(out);
        }
        System.out.println("over");
    }
}
*/
