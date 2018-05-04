package com.eastrobot.converter;

import com.eastrobot.converter.base.AbstractConverter;
import com.eastrobot.converter.base.ConverterFactory;
import com.eastrobot.converter.base.Type;

/**
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-05-04 17:45
 */
public class TestMain {
    public static void main(String[] args) throws Exception {

        testXls();
    }

    public static void testDoc() throws Exception {
        // Test convert Word
        AbstractConverter docConverter = ConverterFactory.getConverter(Type.WORD,
                "C:\\Users\\User\\Desktop\\demo\\kbase-media-2003.doc");
        docConverter.convert();
    }

    public static void testXls() throws Exception {
        // Test convert Excel
        AbstractConverter xlsConverter = ConverterFactory.getConverter(Type.EXCEL,
                "C:\\Users\\User\\Desktop\\demo\\201405301253002480.xls");
        xlsConverter.convert();
    }

    public static void testPpt() throws Exception {
        // Test convert PowerPoint
        AbstractConverter pptConverter = ConverterFactory.getConverter(Type.POWERPOINT,
                "C:\\Users\\User\\Desktop\\demo\\kbase-media-2003.doc");
        pptConverter.convert();
    }
}
