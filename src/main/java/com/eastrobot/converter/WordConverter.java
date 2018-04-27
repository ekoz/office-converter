package com.eastrobot.converter;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.*;

import java.io.*;
import java.util.LinkedList;
import java.util.List;
import java.util.UUID;

/**
 * WordConverter office word转html (doc)
 * <p>
 * Section ->  Paragraph -> CharacterRun
 *
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-04-26 16:23
 */
public class WordConverter {
    private static String outputDirectory = "C:\\Users\\User\\Desktop\\demo\\";

    private static String outputImageDirectory = "C:\\Users\\User\\Desktop\\demo\\images\\";
    /**
     * 输入文件的绝对路径
     */
    private String wordPath;
    /**
     * 文档
     */
    private HWPFDocument doc;
    /**
     * 文档可读区域
     */
    private Range range;
    /**
     * 图片存储区
     */
    private PicturesTable picturesTable;
    /**
     * 表格迭代器
     */
    private TableIterator tableIterator;
    /**
     * 记录表格起始偏移
     */
    private int tbStartOffset[];
    /**
     * 记录表格
     */
    private List<Table> tableList;

    public WordConverter(String wordPath) {
        this.wordPath = wordPath;
    }

    /**
     * 准备环境
     */
    public void prepareEnv() throws Exception {
        this.doc = new HWPFDocument(new FileInputStream(wordPath));
        this.range = doc.getRange();
        this.picturesTable = doc.getPicturesTable();
        this.tableIterator = new TableIterator(range);
        // 确保路径一定存在
        new File(outputDirectory).mkdirs();
        new File(outputImageDirectory).mkdirs();

        this.tableList = new LinkedList<>();
        this.tbStartOffset = new int[100];
    }

    /**
     * 转换方法 请先调用{@link  WordConverter#prepareEnv}
     */
    public void convert() throws Exception {
        String htmlHead = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" /><title>"
                + doc.getSummaryInformation().getTitle()
                + "</title></head><body><div style='margin:60px;text-align:center;'><div style='width:620px;" +
                "text-align:left;line-height:24px;'>";
        StringBuilder htmlBuilder = new StringBuilder(htmlHead);
        StringBuilder scanBuilder = new StringBuilder();

        // 1. 预处理表格 找到其在文档中的相对位置
        boolean hasTable = false;
        int tablePos = 0;

        while (tableIterator.hasNext()) {
            Table table = tableIterator.next();
            tbStartOffset[tablePos++] = table.getStartOffset();
            tableList.add(table);
            hasTable = true;
        }

        // 2, 初始化目录处理器
        TocHandler tocHandler = new TocHandler(doc);
        tocHandler.init();
        System.out.println(doc.characterLength());

        int characterSize = 0;
        // 3. 开始处理每个字符
        tablePos = 0;
        for (int i = 0; i < range.numSections(); i++) {
            Section section = range.getSection(i);
            for (int j = 0; j < section.numParagraphs(); j++) {
                Paragraph paragraph = section.getParagraph(j);
                // 得到段落的对齐方式 居中的不成为目录和标题
                // int justification = paragraph.getJustification();
                // if (justification == WordConstant.CENTER_ALIGN) {
                //     htmlBuilder.append("<center>");
                //     handleCharacter(paragraph, hasTable, tablePos, htmlBuilder, scanBuilder);
                //     htmlBuilder.append("</center>");
                // } else {
                boolean isToc = tocHandler.convert(paragraph, htmlBuilder);
                if (!isToc) {
                    handleCharacter(paragraph, hasTable, tablePos, htmlBuilder, scanBuilder);
                }
                // }
            }
        }

        StringBuilder tocBuilder = tocHandler.getToc(true);
        writeFile(tocBuilder.append(htmlBuilder).toString());
    }


    private void handleCharacter(Paragraph paragraph, boolean hasTable, int tablePos, StringBuilder htmlBuilder,
                                 StringBuilder scanBuilder) throws Exception {
        for (int k = 0; k < paragraph.numCharacterRuns(); k++) {
            if (k == 0) {
                continue;
            }
            CharacterRun c1 = paragraph.getCharacterRun(k - 1);
            // 处理表格
            if (hasTable) {
                if (k == tbStartOffset[tablePos]) {
                    Table table = tableList.get(tablePos);
                    htmlBuilder.append(scanBuilder).append(tableToHtml(table));
                    k = table.getEndOffset() - 1;
                    tablePos++;
                    scanBuilder.setLength(0);
                }
            }
            // 图片
            if (picturesTable.hasPicture(c1)) {
                htmlBuilder.append(scanBuilder).append(pictureToHtml(picturesTable, c1));
                scanBuilder.setLength(0);
            } else {//正常文本
                CharacterRun c2 = paragraph.getCharacterRun(k);
                char c = c1.text().charAt(0);

                if (c == 13) {// 回车
                    scanBuilder.append("<br/>");
                } else if (c == 32) {// 空格符
                    scanBuilder.append(" ");
                } else if (c == 9) {     //水平制表符
                    scanBuilder.append("    ");
                }

                // 比较前后2个字符是否具有相同的格式
                boolean isSame = c1.isBold() == c2.isBold() && c1.isItalic() == c2.isItalic()
                        && c1.getFontName().equals(c2.getFontName()) && c1.getFontSize() ==
                        c2.getFontSize();
                if (isSame) {
                    scanBuilder.append(c1.text());
                } else {
                    StringBuilder fontStyleBuilder = new StringBuilder("<span style=\"font-family:")
                            .append(c1.getFontName()).append(";font-size:")
                            .append(c1.getFontSize() / 2).append("pt;");
                    if (c1.isBold())
                        fontStyleBuilder.append("font-weight:bold;");
                    if (c1.isItalic())
                        fontStyleBuilder.append("font-style:italic;");
                    if (c1.isStrikeThrough())
                        fontStyleBuilder.append("text-decoration:line-through;");

                    int fontColor = c1.getIco24();
                    int[] rgb = new int[3];
                    if (fontColor != -1) {
                        rgb[0] = (fontColor >> 0) & 0xff; // red;
                        rgb[1] = (fontColor >> 8) & 0xff; // green
                        rgb[2] = (fontColor >> 16) & 0xff; // blue
                    }
                    fontStyleBuilder.append("color: rgb(").append(rgb[0]).append(",").append(rgb[1])
                            .append(",")
                            .append(rgb[2]).append(");");
                    htmlBuilder.append(fontStyleBuilder).append("\">").append(scanBuilder).append(c1.text())
                            .append("</span>");
                    //文字段
                    scanBuilder.setLength(0);
                }
            }
        }
    }

    /**
     * 遍历生成表格
     */
    private String tableToHtml(Table table) {
        StringBuilder tableBuilder = new StringBuilder("<table border='1' cellpadding='0' cellspacing='0'>");
        for (int i = 0; i < table.numRows(); i++) {
            TableRow row = table.getRow(i);
            tableBuilder.append("<tr align='center'>");
            for (int j = 0; j < row.numCells(); j++) {
                TableCell td = row.getCell(j);
                int cellWidth = td.getWidth();
                // 得到具体表格内容
                for (int k = 0; k < td.numParagraphs(); k++) {
                    Paragraph para = td.getParagraph(k);
                    CharacterRun cr = para.getCharacterRun(0);
                    String fontStyle = "<span style=\"font-family:" + cr.getFontName() + ";font-size:"
                            + cr.getFontSize() / 2 + "pt;color:" + cr.getColor() + ";";

                    if (cr.isBold())
                        fontStyle += "font-weight:bold;";
                    if (cr.isItalic())
                        fontStyle += "font-style:italic;";

                    String fontSpan = fontStyle + "\">" + para.text().toString().trim() + "</span>";
                    tableBuilder.append("<td width=").append(cellWidth).append(">").append(fontSpan).append("</td>");
                } // end for
            } // end for
        } // end for
        tableBuilder.append("</table>");

        return tableBuilder.toString();
    }

    /**
     * 插入图片img
     */
    private String pictureToHtml(PicturesTable picturesTable, CharacterRun cr) throws Exception {
        StringBuilder pictureBuilder = new StringBuilder();
        // 提取图片
        Picture pic = picturesTable.extractPicture(cr, false);

        // 获取图片样式
        int picHeight = pic.getHeight() * pic.getVerticalScalingFactor() / 100;
        int picWidth = pic.getHorizontalScalingFactor() * pic.getWidth() / 100;
        if (picWidth > 500) {
            picHeight = 500 * picHeight / picWidth;
            picWidth = 500;
        }
        String style = " style='height:" + picHeight + "px;width:" + picWidth + "px'";

        String imageId = UUID.randomUUID().toString();
        String originName = pic.suggestFullFileName();
        String newImageName = outputImageDirectory + imageId + "." + FilenameUtils.getExtension(originName);
        pic.writeImageContent(new FileOutputStream(newImageName));
        pictureBuilder.append("<img ").append(style).append(" src=\"").append(newImageName).append("\"/>");
        if (pic.getWidth() > 450) {
            pictureBuilder.append("<br/>");
        }

        return pictureBuilder.toString();
    }

    /**
     * 写入文件
     */
    private void writeFile(String content) {
        FileOutputStream fos = null;
        BufferedWriter bw = null;
        content = content.replaceAll("EMBED", "").replaceAll("Equation.DSMT4", "") + "</div></div></body></html>";
        try {
            String baseName = FilenameUtils.getBaseName(wordPath);
            fos = new FileOutputStream(outputDirectory + baseName + ".html");
            bw = new BufferedWriter(new OutputStreamWriter(fos, "utf-8"));
            bw.append(content);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (bw != null)
                    bw.close();
                if (fos != null)
                    fos.close();
            } catch (IOException ie) {
                ie.printStackTrace();
            }
        }
    }
}
