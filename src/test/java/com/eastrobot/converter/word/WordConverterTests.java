/*
 * Power by www.xiaoi.com
 */
package com.eastrobot.converter.word;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.dom4j.Document;
import org.junit.Test;

import com.eastrobot.converter.base.AbstractConverter;
import com.eastrobot.converter.base.ConverterFactory;
import com.eastrobot.converter.base.Type;
import com.eastrobot.converter.util.HtmlUtil;
import com.eastrobot.converter.util.OfficeUtil;

/**
 * @author <a href="mailto:eko.z@outlook.com">eko.zhan</a>
 * @date 2018年5月15日 上午9:04:48
 * @version 1.0
 */
public class WordConverterTests {

	@Test
	public void testConvertDoc() throws Exception{
		 AbstractConverter docConverter = ConverterFactory.getConverter(Type.WORD,
	                "E:\\ConvertTester\\CeairFile\\20151109\\201511091008144920.doc");
	        docConverter.convert();
	}
	
	@Test
	public void testConvertDocx() throws Exception{
		 AbstractConverter docConverter = ConverterFactory.getConverter(Type.WORDX,
	                "E:\\ConvertTester\\CeairFile\\20151109\\201511091015405980.docx");
	        docConverter.convert();
	}
	
	@Test
	public void testPoiDocx() throws IOException{
		File file = new File("E:\\ConvertTester\\CeairFile\\20151109\\201511091517597990.docx");
		XWPFDocument docx = OfficeUtil.loadDocx(file);
		
        XHTMLOptions options = XHTMLOptions.create();
        // 存放图片的文件夹
        options.setExtractor(new FileImageExtractor(new File("E:\\converter-html\\" + FilenameUtils.getBaseName(file.getName()) + "\\image")));
        // html中图片的路径
        options.URIResolver(new BasicURIResolver("image"));
        File dir = new File("E:\\converter-html\\" + FilenameUtils.getBaseName(file.getName()));
        if (!dir.exists()) dir.mkdirs();
        String filePath = "E:\\converter-html\\" + FilenameUtils.getBaseName(file.getName()) + "\\" + FilenameUtils.getBaseName(file.getName()) + ".html";
        XHTMLConverter xhtmlConverter = (XHTMLConverter) XHTMLConverter.getInstance();
        xhtmlConverter.convert(docx, new OutputStreamWriter(new FileOutputStream(filePath)), options);
		
	}
}
