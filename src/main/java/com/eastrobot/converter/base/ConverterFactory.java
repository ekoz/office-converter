package com.eastrobot.converter.base;

import com.eastrobot.converter.excel.ExcelConverter;
import com.eastrobot.converter.powerpoint.PowerPointConverter;
import com.eastrobot.converter.word.WordConverter;

/**
 * ConverterFactory 未指定输出路径使用配置的输出路径
 *
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-05-04 17:25
 */
public class ConverterFactory {

    private static String outputPath = "E:\\converter-html\\";

    public static AbstractConverter getConverter(Type type, String inputFilePath) {
        AbstractConverter converter = null;
        switch (type) {
            case WORD:
                converter = new WordConverter(inputFilePath, outputPath);
                break;
            case EXCEL:
                converter = new ExcelConverter(inputFilePath, outputPath);
                break;
            case POWERPOINT:
                converter = new PowerPointConverter(inputFilePath, outputPath);
                break;
            default:
                break;
        }

        return converter;
    }

}
