package com.eastrobot.converter;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.model.StyleDescription;
import org.apache.poi.hwpf.model.StyleSheet;
import org.apache.poi.hwpf.usermodel.*;

import java.io.*;

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
     * 全文构建器
     */
    private StringBuilder htmlBuilder;

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
        // 确保路径一定存在
        new File(outputDirectory).mkdirs();
        new File(outputImageDirectory).mkdirs();

        this.htmlBuilder = new StringBuilder(2048);
    }

    /**
     * 转换方法 请先调用{@link  WordConverter#prepareEnv}
     */
    public void convert() throws Exception {
        String htmlHead = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" /><title>"
                + doc.getSummaryInformation().getTitle()
                + "</title></head><body><div style='margin:60px;text-align:center;'><div style='width:620px;" +
                "text-align:left;line-height:24px;'>";

        // 2, 初始化目录处理器
        TocHandler tocHandler = new TocHandler(doc);
        tocHandler.init();


        int page = 1;
        int currentTableLevel = Integer.MIN_VALUE;
        // 3. 开始处理每个字符
        for (int i = 0; i < range.numSections(); i++) {
            Section section = range.getSection(i);
            processParagraphs(section, tocHandler, currentTableLevel, page);
        }

        StringBuilder tocBuilder = tocHandler.getToc(true);
        writeFile(tocBuilder.append(htmlBuilder).toString());
    }

    private void processParagraphs(Section section, TocHandler tocHandler, int currentTableLevel, int page) throws
            Exception {
        StyleSheet docStyleSheet = doc.getStyleSheet();
        int paragraphs = range.numParagraphs();
        for (int p = 0; p < paragraphs; p++) {
            Paragraph paragraph = section.getParagraph(p);
            if (paragraph.isInTable() && paragraph.getTableLevel() != currentTableLevel) {

                Table table = section.getTable(paragraph);
                String tableHtml = tableToHtml(table);
                htmlBuilder.append(tableHtml);

                p += table.numParagraphs();
                p--;
                continue;
            }

            if (paragraph.text().equals("\u000c")) {
                htmlBuilder.append("<center>第").append(page++).append("页</center>");
            }

            // 段落的样式
            StyleDescription style = docStyleSheet.getStyleDescription(paragraph.getStyleIndex());
            String styleName = style.getName();
            // 段落的对齐方式 居中的不成为目录
            if (WordConstant.CENTER_ALIGN == paragraph.getJustification()) {
                htmlBuilder.append("<center>");
                if (styleName.contains("标题")) {
                    htmlBuilder.append("<h3>");
                }
                processNoTableCharacter(paragraph);
                htmlBuilder.append("<br/>");
                if (styleName.contains("标题")) {
                    htmlBuilder.append("</h3>");
                }
                htmlBuilder.append("</center>");
            } else {
                boolean isToc = tocHandler.convert(paragraph, htmlBuilder, styleName);
                if (!isToc) {
                    processNoTableCharacter(paragraph);
                    htmlBuilder.append("<br/>");
                }
            }
        }
    }

    /**
     * 处理非表格外的内容
     */
    private void processNoTableCharacter(Paragraph paragraph) throws Exception {
        int num = paragraph.numCharacterRuns();
        for (int i = 0; i < num; i++) {
            CharacterRun c = paragraph.getCharacterRun(i);
            // 图片
            if (picturesTable.hasPicture(c)) {
                htmlBuilder.append(pictureToHtml(picturesTable, c));
            } else {//正常文本
                String characterStyle = handleCharacterStyle(c);
                htmlBuilder.append(characterStyle);
            }
        }
    }

    /**
     * 处理字体的样式
     */
    private String handleCharacterStyle(CharacterRun c) {
        StringBuilder fontStyleBuilder = new StringBuilder("<span style=\"font-family:")
                .append(c.getFontName()).append(";font-size:")
                .append(c.getFontSize() / 2).append("pt;");
        if (c.isBold())
            fontStyleBuilder.append("font-weight:bold;");
        if (c.isItalic())
            fontStyleBuilder.append("font-style:italic;");
        if (c.isStrikeThrough())
            fontStyleBuilder.append("text-decoration:line-through;");

        int fontColor = c.getIco24();
        int[] rgb = new int[3];
        if (fontColor != -1) {
            rgb[0] = (fontColor >> 0) & 0xff; // red;
            rgb[1] = (fontColor >> 8) & 0xff; // green
            rgb[2] = (fontColor >> 16) & 0xff; // blue
        }
        fontStyleBuilder.append("color: rgb(").append(rgb[0]).append(",").append(rgb[1])
                .append(",")
                .append(rgb[2]).append(");").append("\">").append(c.text()).append("</span>");

        return fontStyleBuilder.toString();
    }

    /**
     * 生成表格html
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

        String originFullNameName = outputImageDirectory + pic.suggestFullFileName();
        pic.writeImageContent(new FileOutputStream(originFullNameName));
        pictureBuilder.append("<img ").append(style).append(" src=\"").append(originFullNameName).append("\"/>");
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
