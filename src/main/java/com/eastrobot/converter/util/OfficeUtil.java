package com.eastrobot.converter.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.util.IOUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

/**
 * OfficeUtil
 *
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-05-02 13:57
 */
public class OfficeUtil {

    /**
     * 加载xls对象
     */
    public static HSSFWorkbook loadXls(File xlsFile) throws IOException {
        final FileInputStream inputStream = new FileInputStream(xlsFile);
        try {
            return new HSSFWorkbook(inputStream);
        } finally {
            IOUtils.closeQuietly(inputStream);
        }
    }

    /**
     * 加载doc对象
     */
    public static HWPFDocument loadDoc(File docFile) throws IOException {
        final FileInputStream inputStream = new FileInputStream(docFile);
        try {
            return new HWPFDocument(inputStream);
        } finally {
            IOUtils.closeQuietly(inputStream);
        }
    }
}
