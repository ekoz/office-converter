package com.eastrobot.converter;

import com.eastrobot.converter.word.WordConverter;

/**
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-04-27 16:42
 */
public class WordMainTest {
    public static void main(String[] args) throws Exception {
        long s = System.currentTimeMillis();
        WordConverter converter = new WordConverter("C:\\Users\\User\\Desktop\\demo\\kbase-media-2003.doc");
        converter.prepareEnv();
        converter.convert();
        long e = System.currentTimeMillis();
        System.out.println("use timed -> " + (e - s));
    }
}
