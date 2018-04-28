package com.eastrobot.converter.util;

import org.apache.commons.io.FilenameUtils;

import java.io.File;

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
        String folder = FilenameUtils.getFullPath(inputPath) + FilenameUtils.getBaseName(inputPath) + File.separator
                + subPath;

        File file = new File(folder);
        if (!file.exists()) {
            file.mkdirs();
        }

        return folder;
    }
}
