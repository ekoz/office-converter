package com.eastrobot.converter.word;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.dom4j.Element;
import org.dom4j.tree.DefaultElement;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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
    private Pattern pattern = Pattern.compile("H[1-6]");
    /**
     * 记录目录结构
     */
    private Element tocElement;

    public TocHandler(HWPFDocument doc) {
        this.doc = doc;
        tocElement = new DefaultElement("div");
    }

    public void init() {
        this.headNum = new int[headLevel];
        this.headList = new ArrayList<String>(Arrays.asList("H1", "H2", "H3", "H4", "H5", "H6"));
        this.mainTextTitleList = new ArrayList<String>(Arrays.asList("标题 1", "标题 2", "标题 3", "标题 4", "标题 5", "标题 6"));
        tocElement.addAttribute("id", "tocNav");
    }

    public Element getToc(boolean hasConvertedOver) {
        if (hasConvertedOver) {
            return tocElement;
        }

        return null;
    }

    /**
     * 给我一个段落 我来判断是否是标题段落或者是目录,
     * 如果是目录我会保存目录html结构,并在正文中创建锚点 <=(≖ ‿ ≖)✧
     *
     * @return 是否是目录段落结构
     */
    public boolean convert(Paragraph paragraph, Element mainDiv, String styleName) {
        // 第一种是没有生成目录 但是文章中有标题结构 这样就直接去找对应的标题级别
        int position = mainTextTitleList.indexOf(styleName);
        if (position < 0) {
            //第二种是生成了目录 就匹配到对应的级别
            Matcher matcher = pattern.matcher(styleName);
            if (matcher.find()) {
                String group = matcher.group();
                position = headList.indexOf(group);
            }
        }
        if (position > -1) {
            Element p = tocElement.addElement("p");

            // 标题定位缩进
            int padingLeft = position * 30;
            p.addAttribute("style", "padding-left:" + padingLeft + "px");
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
            p.addElement("a").addAttribute("href", "#" + key).addText(headId + text);
            // 设置正文锚点
            String head = headList.get(position);
            String fontName = paragraph.getCharacterRun(0).getFontName();
            int fontSize = paragraph.getCharacterRun(0).getFontSize();

            mainDiv.addElement("span").addAttribute("id", key)
                    .addAttribute("style", "font-family:Times New Roman;" +
                            "font-size:" + fontSize + ";font-weight:bold;" +
                            "padding-left:" + (position > 0 ? 30 : 0) + "px").addText(headId + text);
            mainDiv.addElement("br");
            // 当前级别后续的子级别重置
            for (int i = position + 1; i < headLevel; i++) {
                headNum[i] = 0;
            }

            return true;
        }

        return false;
    }
}
