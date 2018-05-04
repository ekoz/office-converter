package com.eastrobot.converter.util;

import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.HTMLWriter;
import org.dom4j.io.OutputFormat;
import org.dom4j.tree.DefaultDocument;

import java.io.StringWriter;

/**
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-04-28 14:03
 */
public class HtmlUtil {

    public static Document createHtmlDocument() {
        return new DefaultDocument().addDocType("html", "", "");
    }

    /**
     * 设置编码
     */
    public static void charset(Element head) {
        head.addElement("meta")
                .addAttribute("http-equiv", "Content-Type")
                .addAttribute("content", "text/html; charset=utf-8");
    }

    /**
     * 设置标题
     */
    public static void title(Element head, String title) {
        head.addElement("title").addText(title);
    }

    /**
     * 格式化
     */
    public static String prettyHtml(Document document) throws Exception {
        StringWriter sw = new StringWriter();
        OutputFormat format = OutputFormat.createPrettyPrint();
        HTMLWriter writer;
        writer = new HTMLWriter(sw, format);
        writer.write(document);
        writer.flush();

        return sw.toString();
    }
}
