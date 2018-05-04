package com.eastrobot.converter.word;

import com.eastrobot.converter.util.HtmlUtil;
import com.eastrobot.converter.util.OfficeUtil;
import com.eastrobot.converter.util.ResourceUtil;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.model.StyleDescription;
import org.apache.poi.hwpf.model.StyleSheet;
import org.apache.poi.hwpf.usermodel.*;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.tree.DefaultElement;
import sun.misc.BASE64Encoder;

import java.io.File;

import static org.apache.poi.hwpf.converter.AbstractWordUtils.TWIPS_PER_INCH;
import static org.apache.poi.hwpf.converter.AbstractWordUtils.TWIPS_PER_PT;

/**
 * WordConverter office word转html (doc)
 * <p>
 * Section ->  Paragraph -> CharacterRun
 *
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-04-26 16:23
 */
public class WordConverter {
    /**
     * 输入文件的绝对路径
     */
    private String wordPath;
    private String outputPath;
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

    private Document root;

    /**
     * @param wordPath doc文档路径
     * @param outputPath 输出文件路径
     */
    public WordConverter(String wordPath, String outputPath) {
        this.wordPath = wordPath;
        this.outputPath = outputPath;
    }

    /**
     * 准备环境
     */
    public WordConverter prepareEnv() throws Exception {
        this.doc = OfficeUtil.loadDoc(new File(this.wordPath));
        this.range = doc.getRange();
        this.picturesTable = doc.getPicturesTable();
        this.root = HtmlUtil.createHtmlDocument();

        return this;
    }

    /**
     * 转换方法 请先调用{@link  WordConverter#prepareEnv}
     */
    public void convert() throws Exception {
        Element html = root.addElement("html");
        Element head = html.addElement("head");
        Element body = html.addElement("body");
        HtmlUtil.charset(head);
        HtmlUtil.title(head, doc.getSummaryInformation().getTitle() != null ? doc.getSummaryInformation().getTitle() : "");

        Element mainDiv = new DefaultElement("div");
        mainDiv.addAttribute("style", "width:75%");

        // 初始化目录处理器
        TocHandler tocHandler = new TocHandler(doc);
        tocHandler.init();

        int currentTableLevel = Integer.MIN_VALUE;
        // 开始处理每个字符
        for (int i = 0; i < range.numSections(); i++) {
            Section section = range.getSection(i);
            processParagraphs(mainDiv, section, tocHandler, currentTableLevel);
        }

        Element toc = tocHandler.getToc(true);
        // 目录有内容才添加
        if (StringUtils.isNotBlank(toc.getText())) {
            body.add(toc);
        }
        body.add(mainDiv);

        ResourceUtil.writeFile(outputPath + FilenameUtils.getBaseName(wordPath) + ".html", root.asXML());
    }

    private void processParagraphs(Element mainDiv, Section section, TocHandler tocHandler, int currentTableLevel)
            throws Exception {
        StyleSheet docStyleSheet = doc.getStyleSheet();
        int paragraphs = section.numParagraphs();
        for (int p = 0; p < paragraphs; p++) {
            Paragraph paragraph = section.getParagraph(p);
            StyleDescription style = docStyleSheet.getStyleDescription(paragraph.getStyleIndex());
            String styleName = style.getName();

            // 跳过本身生成的目录 从正文结构生成
            if (styleName.contains(WordConstant.STYLE_TOC)) {
                continue;
            }

            // 处理表格
            if (paragraph.isInTable() && paragraph.getTableLevel() != currentTableLevel) {
                Table table = section.getTable(paragraph);
                Element tableEle = processTable(table);
                mainDiv.add(tableEle);

                p += table.numParagraphs(); // 跳过表格占用的段
                p--;
                continue;
            }

            // 段落的对齐方式 居中的不成为目录结构中的内容
            if (WordConstant.CENTER_ALIGN == paragraph.getJustification()) {
                // 跳过目录两个字的段
                if (paragraph.text().trim().equals(WordConstant.STYLE_TOC)) {
                    continue;
                }
                Element activeEle = mainDiv.addElement("center");
                if (styleName.contains("标题")) {
                    activeEle = activeEle.addElement("h3");
                }
                processNotTableCharacter(activeEle, paragraph);
            } else {
                boolean isToc = tocHandler.convert(paragraph, mainDiv, styleName);
                if (!isToc) {
                    processNotTableCharacter(mainDiv, paragraph);
                }
            }
        }
    }

    /**
     * 处理非表格外的内容
     */
    private void processNotTableCharacter(Element mainDiv, Paragraph paragraph) throws Exception {
        Element p = new DefaultElement("p");
        p.addAttribute("style", "text-indent:2em;");
        int num = paragraph.numCharacterRuns();
        for (int i = 0; i < num; i++) {
            CharacterRun c = paragraph.getCharacterRun(i);
            // 图片
            if (picturesTable.hasPicture(c)) {
                Element picEle = processPicture(c);
                mainDiv.add(picEle);
            } else {//正常文本
                // 超链接
                if (c.text().contains("HYPERLINK")) {
                    i = i + 2; //跳到link的文本段
                    CharacterRun link = paragraph.getCharacterRun(i);
                    String text = link.text().trim();
                    Element a = new DefaultElement("a");
                    a.addAttribute("style", "font-family:Times New Roman;font-size:10pt;color: rgb(0,0,255);")
                            .addElement("a")
                            .addAttribute("href", text)
                            .addText(text);
                    p.add(a);
                } else {
                    Element span = processCharacterRun(c);
                    p.add(span);
                }
            }
        }
        mainDiv.add(p);
    }

    /**
     * 处理 characterRun
     */
    private Element processCharacterRun(CharacterRun c) {
        Element span = new DefaultElement("span");
        StringBuilder fontStyleBuilder = new StringBuilder("font-family:")
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
                .append(rgb[2]).append(");");
        span.addAttribute("style", fontStyleBuilder.toString());
        span.addText(c.text().trim());

        return span;
    }

    /**
     * 处理表格元素
     */
    private Element processTable(Table table) {
        Element tableEle = new DefaultElement("table");
        tableEle.addAttribute("border", "1")
                .addAttribute("cellpadding", "1")
                .addAttribute("cellspacing", "0")
                .addAttribute("align", "center");
        for (int i = 0; i < table.numRows(); i++) {
            TableRow row = table.getRow(i);

            Element tr = tableEle.addElement("tr").addAttribute("align", "center");
            for (int j = 0; j < row.numCells(); j++) {
                TableCell td = row.getCell(j);
                float inch = td.getWidth() / TWIPS_PER_PT;

                Element outSpan = new DefaultElement("span");
                for (int k = 0; k < td.numCharacterRuns(); k++) {
                    CharacterRun cr = td.getCharacterRun(k);
                    Element span = new DefaultElement("span");
                    StringBuilder fontBuilder = new StringBuilder(40);
                    fontBuilder.append("font-family:").append(cr.getFontName())
                            .append(";font-size:").append(cr.getFontSize() / 2).append("pt;");
                    if (cr.isBold())
                        fontBuilder.append("font-weight:bold;");
                    if (cr.isItalic())
                        fontBuilder.append("font-style:italic;");

                    int[] rgb = new int[3];
                    int color = cr.getIco24();
                    if (color != -1) {
                        rgb[0] = (color >> 0) & 0xff; // red;
                        rgb[1] = (color >> 8) & 0xff; // green
                        rgb[2] = (color >> 16) & 0xff; // blue
                    }
                    fontBuilder.append("color: rgb(").append(rgb[0]).append(",").append(rgb[1])
                            .append(",")
                            .append(rgb[2]).append(");");

                    outSpan.addAttribute("style", fontBuilder.toString()).addText(cr.text().trim());
                }
                tr.addElement("td").addAttribute("width", inch + "pt;").add(outSpan);
            } // end for
        } // end for

        return tableEle;
    }

    /**
     * 处理图片元素 图片缩放比例 图片裁剪大小
     */
    private Element processPicture(CharacterRun cr) throws Exception {
        // 提取图片
        Picture picture = picturesTable.extractPicture(cr, true);

        // 图片原始缩放因子 千分位
        float aspectRatioX = picture.getHorizontalScalingFactor();
        float aspectRatioY = picture.getVerticalScalingFactor();

        //图片长宽
        float imageWidth;
        float imageHeight;

        // 图片裁剪因子
        float cropTop;
        float cropBottom;
        float cropLeft;
        float cropRight;

        // 处理图片纵横比
        if (aspectRatioX > 0) {
            imageWidth = picture.getDxaGoal() * aspectRatioX / 1000 / TWIPS_PER_INCH;
            cropRight = picture.getDxaCropRight() * aspectRatioX / 1000 / TWIPS_PER_INCH;
            cropLeft = picture.getDxaCropLeft() * aspectRatioX / 1000 / TWIPS_PER_INCH;
        } else {
            imageWidth = picture.getDxaGoal() / TWIPS_PER_INCH;
            cropRight = picture.getDxaCropRight() / TWIPS_PER_INCH;
            cropLeft = picture.getDxaCropLeft() / TWIPS_PER_INCH;
        }

        if (aspectRatioY > 0) {
            imageHeight = picture.getDyaGoal() * aspectRatioY / 1000 / TWIPS_PER_INCH;
            cropTop = picture.getDyaCropTop() * aspectRatioY / 1000 / TWIPS_PER_INCH;
            cropBottom = picture.getDyaCropBottom() * aspectRatioY / 1000 / TWIPS_PER_INCH;
        } else {
            imageHeight = picture.getDyaGoal() / TWIPS_PER_INCH;
            cropTop = picture.getDyaCropTop() / TWIPS_PER_INCH;
            cropBottom = picture.getDyaCropBottom() / TWIPS_PER_INCH;
        }

        BASE64Encoder encoder = new BASE64Encoder();
        String base64Img = "data:image/jpeg;base64," + encoder.encode(picture.getContent());

        Element imgDivEle = new DefaultElement("div");
        if (cropTop != 0 || cropRight != 0 || cropBottom != 0 || cropLeft != 0) {
            float visibleWidth = Math.max(0, imageWidth - cropLeft - cropRight);
            float visibleHeight = Math.max(0, imageHeight - cropTop - cropBottom);

            imgDivEle.addAttribute("style", "vertical-align:text-bottom;width:" + visibleWidth + "in;height:" +
                    visibleHeight + "in;");
            Element div = imgDivEle.addElement("div");
            div.addAttribute("style", "position:relative;width:" + visibleWidth + "in;height:" + visibleHeight + "in;" +
                    "overflow:hidden;");
            div.addElement("img").addAttribute("style", "position:absolute" +
                    ";left:-" + cropLeft +
                    ";top:-" + cropTop +
                    ";width:" + imageWidth + "in" +
                    ";height:" + imageHeight + "in;")
                    .addAttribute("src", base64Img);
        } else {
            imgDivEle.addElement("img")
                    .addAttribute("style", "width:" + imageWidth + "in;height:" + imageHeight + "in;" +
                            "vertical-align:text-bottom;")
                    .addAttribute("src", base64Img);
        }

        return imgDivEle;
    }
}
