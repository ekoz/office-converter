package com.eastrobot.poi;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.StyleDescription;
import org.apache.poi.hwpf.model.StyleSheet;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class TocUtil {

    public static void getWordTitles2003(String path) throws IOException {
        File file = new File(path);
        String filename = file.getName();
        InputStream is = new FileInputStream(path);
        HWPFDocument doc = new HWPFDocument(is);
        Range r = doc.getRange();

        // 文档样式
        StyleSheet docStyleSheet = doc.getStyleSheet();
        int docStyleNum = docStyleSheet.numStyles();
        // 定义最多9级标题
        Pattern pattern = Pattern.compile("H[1-9]");
        int headLevel = 9;
        int[] headNum = new int[headLevel];
        List<String> headList = new ArrayList<>(Arrays.asList("H1", "H2", "H3", "H4", "H5", "H6", "H7", "H8", "H9"));

        System.out.println(doc.characterLength());
        for (int i = 0; i < r.numParagraphs(); i++) {
            Paragraph p = r.getParagraph(i);
            int styleIndex = p.getStyleIndex();
            // 检查样式是否在范围内
            if (docStyleNum > styleIndex) {
                StyleDescription style = docStyleSheet.getStyleDescription(styleIndex);
                String styleName = style.getName();
                if (StringUtils.isNotBlank(styleName)) {
                    // 找到对应的是几级标题
                    Matcher matcher = pattern.matcher(styleName);
                    if (matcher.find()) {
                        String head = matcher.group();
                        int headPosition = headList.indexOf(head);
                        // 标题定位设置
                        for (int j = 0; j <= headPosition; j++) {
                            System.out.print("\t");
                        }
                        headNum[headPosition]++;
                        // 标题加上标题的数字
                        for (int j = 0; j <= headPosition; j++) {
                            System.out.print(headNum[j] + ".");
                        }
                        System.out.println(p.text());
                        // 当前级别后续的子级别重置
                        for (int j = headPosition + 1; j < headLevel; j++) {
                            headNum[j] = 0;
                        }
                    } else {
                        System.out.println(styleName + ">>>>>>" + p.text() + "<<<<<" + r.getTable(p));
                        // styleName = "正文";
                    }
                }
            }
        }
    }

    private static String preHandleToc(String path) throws Exception {
        File file = new File(path);
        String filename = file.getName();
        InputStream is = new FileInputStream(path);
        HWPFDocument doc = new HWPFDocument(is);
        Range range = doc.getRange();
        StringBuilder tocBuilder = new StringBuilder();
        Map<String, String> tocMap = new LinkedHashMap<>();
        // 文档样式
        StyleSheet docStyleSheet = doc.getStyleSheet();
        int docStyleNum = docStyleSheet.numStyles();
        // 定义最多9级标题
        Pattern pattern = Pattern.compile("H[1-9]");
        int headLevel = 9;
        int[] headNum = new int[headLevel];
        List<String> headList = new ArrayList<>(Arrays.asList("H1", "H2", "H3", "H4", "H5", "H6", "H7", "H8", "H9"));

        tocBuilder.append("<div>");
        for (int i = 0; i < range.numParagraphs(); i++) {
            Paragraph p = range.getParagraph(i);
            int styleIndex = p.getStyleIndex();
            // 检查样式是否在范围内
            if (docStyleNum > styleIndex) {
                StyleDescription style = docStyleSheet.getStyleDescription(styleIndex);
                String styleName = style.getName();
                if (StringUtils.isNotBlank(styleName)) {
                    // 找到对应的是几级标题
                    Matcher matcher = pattern.matcher(styleName);
                    if (matcher.find()) {
                        String head = matcher.group();
                        int headPosition = headList.indexOf(head);
                        tocBuilder.append("<p>");
                        // 标题定位设置
                        for (int j = 0; j <= headPosition; j++) {
                            tocBuilder.append("&emsp;");
                            // System.out.print("\t");
                        }
                        headNum[headPosition]++;
                        // 标题加上标题的数字
                        StringBuilder headId = new StringBuilder();
                        for (int j = 0; j <= headPosition; j++) {
                            headId.append(headNum[j]).append(".");
                            // System.out.print(headNum[j] + ".");
                        }

                        String key = headId.toString();
                        String text = p.text();
                        tocMap.put(key, text); // 存入map 后面找到标题替换标签锚点
                        tocBuilder.append("<a href=\"").append(key).append("\">").append(headId).append(text).append("</a>");
                        tocBuilder.append("</p>");
                        // System.out.println(p.text());
                        // 当前级别后续的子级别重置
                        for (int j = headPosition + 1; j < headLevel; j++) {
                            headNum[j] = 0;
                        }
                    } else {
                        // 该种方式遍历 文本会丢失
                        // styleName = "正文";
                    }
                }
            }
        }
        tocBuilder.append("</div>");
        tocMap.forEach((key, value)->{
            System.out.println(key + "<-->" + value);
        });

        return tocBuilder.toString();
    }

    public static List<String> getWordTitles2007(String path) throws Exception {
        InputStream is = new FileInputStream(path);
        OPCPackage p = POIXMLDocument.openPackage(path);
        XWPFWordExtractor e = new XWPFWordExtractor(p);
        // POIXMLDocument doc = e.getDocument();
        List<String> list = new ArrayList<String>();
        XWPFDocument doc = new XWPFDocument(is);
        List<XWPFParagraph> paras = doc.getParagraphs();
        for (XWPFParagraph graph : paras) {
            String text = graph.getParagraphText();
            String style = graph.getStyle();
            if ("1".equals(style)) {
                System.out.println(text + "--[" + style + "]");
            } else if ("2".equals(style)) {

                System.out.println(text + "--[" + style + "]");
            } else if ("3".equals(style)) {
                System.out.println(text + "--[" + style + "]");
            } else {
                continue;
            }
            list.add(text);
        }
        return list;
    }

    public static void main(String[] args) throws Exception {
        String path = "C:\\Users\\User\\Desktop\\demo\\201511181708182860.doc";
        List<String> list = new ArrayList<String>();

        // if (path.endsWith(".doc")) {
        //     getWordTitles2003(path);
        // } else if (path.endsWith(".docx")) {
        //     getWordTitles2007(path);
        // }
        System.out.println(preHandleToc(path));
    }
}