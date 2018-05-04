package com.eastrobot.converter.util;

import org.apache.commons.io.FilenameUtils;

import java.io.*;

/**
 * ResourceUtil
 *
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-03-26 9:23
 */
public class ResourceUtil {

    /**
     * 获取文件夹路径, 若不存在则以文件名创建新的子文件夹
     *
     * @author Yogurt_lei
     * @date 2018-03-26 11:46
     */
    public static String getFolder(String inputPath, String subPath) {
        String folder = FilenameUtils.getFullPath(inputPath) + FilenameUtils.getBaseName(inputPath) + File.separator + subPath;

        File file = new File(folder);
        if (!file.exists()) {
            file.mkdirs();
        }

        return folder;
    }

    /**
     * 写入文件
     */
    public static void writeFile(String outputFilePath, String content) {
        FileOutputStream fos = null;
        BufferedWriter bw = null;
        try {
            fos = new FileOutputStream(outputFilePath);
            bw = new BufferedWriter(new OutputStreamWriter(fos, "UTF-8"));
            bw.append(content);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (bw != null)
                    bw.close();
                if (fos != null)
                    fos.close();
            } catch (IOException ie) {
                ie.printStackTrace();
            }
        }
    }
}
