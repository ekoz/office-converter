package com.eastrobot.converter.base;

import org.dom4j.Document;

public abstract class AbstractConverter {
    /**
     * dom4j构建 根root
     */
    protected Document root;

    /**
     * 输入文件路径
     */
    protected String inputFilePath;

    /**
     * 输出路径
     */
    protected String outputPath;

    public AbstractConverter(String inputFilePath, String outputPath) {
        this.inputFilePath = inputFilePath;
        this.outputPath = outputPath;
    }

    public void convert() throws Exception {
        this.prepareEnv();
        this.startConvert();
    }

    protected abstract AbstractConverter prepareEnv() throws Exception;

    protected abstract void startConvert() throws Exception;
}
