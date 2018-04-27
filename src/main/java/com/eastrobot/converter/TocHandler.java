package com.eastrobot.converter;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.StyleDescription;
import org.apache.poi.hwpf.model.StyleSheet;
import org.apache.poi.hwpf.usermodel.Paragraph;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * TocHandler 目录处理器
 *
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-04-27 16:11
 */
public class TocHandler {

    private HWPFDocument doc;
    /**
     * 至多6级标题
     */
    private int headLevel = 6;
    private int[] headNum;
    private List<String> headList;
    private List<String> mainTextTitleList;
    /**
     * 文档样式
     */
    private StyleSheet docStyleSheet;
    /**
     * 记录目录结构
     */
    private StringBuilder tocBuilder;

    public TocHandler(HWPFDocument doc) {
        this.doc = doc;
    }

    public void init() {
        this.tocBuilder = new StringBuilder("<div>");
        this.docStyleSheet = doc.getStyleSheet();
        this.headNum = new int[headLevel];
        this.headList = new ArrayList<>(Arrays.asList("H1", "H2", "H3", "H4", "H5", "H6"));
        this.mainTextTitleList = new ArrayList<>(Arrays.asList("标题 1", "标题 2", "标题 3", "标题 4", "标题 5", "标题 6"));
    }

    public StringBuilder getToc(boolean hasConvertedOver) {
        if (hasConvertedOver) {
            return tocBuilder.append("</div>");
        }

        return tocBuilder;
    }

    /**
     * 给我一个段落 我来判断是否是标题段落或者是目录,
     * 如果是目录我会保存目录html结构,并在正文中创建锚点 <=(≖ ‿ ≖)✧
     *
     * @return 是否是目录
     */
    public boolean convert(Paragraph paragraph, StringBuilder htmlBuilder) {
        int styleIndex = paragraph.getStyleIndex();
        StyleDescription style = docStyleSheet.getStyleDescription(styleIndex);
        String styleName = style.getName();
        if (StringUtils.isNotBlank(styleName)) {
            int position = mainTextTitleList.indexOf(styleName);
            if (position > 0) {
                tocBuilder.append("<p>");

                // 标题定位缩进
                for (int i = 0; i <= position; i++) {
                    tocBuilder.append("&emsp;");
                }
                // 标题序号
                headNum[position]++;
                StringBuilder headId = new StringBuilder();
                for (int i = 0; i <= position; i++) {
                    // 前置0去除
                    if (headNum[i] == 0 && headId.length() == 0) {
                        continue;
                    }
                    headId.append(headNum[i]).append(".");
                }
                // 生成对应toc
                String key = headId.toString();
                String text = paragraph.text();
                // 设置锚点链接
                tocBuilder.append("<a href=\"#").append(key).append("\">").append(headId).append(text).append
                        ("</a></p>");
                // 设置正文锚点
                String head = headList.get(position);
                htmlBuilder.append("<").append(head).append(" id=\"").append(key).append("\"").append(">")
                        .append(headId).append(text)
                        .append("</").append(head).append(">");

                // 当前级别后续的子级别重置
                for (int i = position + 1; i < headLevel; i++) {
                    headNum[i] = 0;
                }

                return true;
            }
        }

        return false;
    }

}
