package com.eastrobot.converter;

/**
 * 定义word解析的常量
 *
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-04-27 15:43
 */
public interface WordConstant {
    /**
     * 表格容量
     */
    Integer TABLE_CAPACITY = 100;
    /**
     * 文字段落的样式 对于标题而言 用正则判断
     */
    String STYLE_TEXT = "正文";
    String STYLE_COVER_TITLE = "封面小标题";
    String STYLE_COVER_SIGNATURE = "封面签名";
    String STYLE_TOC = "目录";
    String STYLE_PARAGRAPH = "段";
    String STYLE_LIST_PARAGRAPHS = "列出段落";

    // 段落对齐方式
    int LEFT_ALIGN = 0;
    int CENTER_ALIGN = 1;
    int RIGHT_ALIGN = 2;
    int JUSTIFIED = 3;
}
