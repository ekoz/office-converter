package com.eastrobot.converter.powerpoint;

import com.eastrobot.converter.base.AbstractConverter;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;

/**
 * PowerPointConverter office powerpoint转html (ppt)
 *
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-05-04 17:21
 */
public class PowerPointConverter extends AbstractConverter {

    private HSLFSlideShow ppt;

    public PowerPointConverter(String inputFilePath, String outputPath) {
        super(inputFilePath, outputPath);
    }

    @Override
    public AbstractConverter prepareEnv() {
        return null;
    }

    @Override
    public void startConvert() {

    }
}
