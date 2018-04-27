package com.eastrobot.poi;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.StyleDescription;
import org.apache.poi.hwpf.model.StyleSheet;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-04-27 14:13
 */
public class TestMain {
    public static void main(String[] args) throws Exception {
        // String wordPath = "C:\\Users\\User\\Desktop\\demo\\201511181708182860.doc";
        // WordConverter converter = new WordConverter(wordPath);
        // converter.prepareEnv();
        // converter.convertToWord();
        foreachTest();
    }

    public static void foreachTest() throws Exception {
        HWPFDocument doc = new HWPFDocument(new FileInputStream
                ("C:\\Users\\User\\Desktop\\demo\\201511181708182860.doc"));
        Range range = doc.getRange();
        StyleSheet docStyleSheet = doc.getStyleSheet();
        int docStyleNum = docStyleSheet.numStyles();
        Pattern pattern = Pattern.compile("H[1-9]");
        int headLevel = 9;
        int[] headNum = new int[headLevel];
        List<String> headList = new ArrayList<>(Arrays.asList("H1", "H2", "H3", "H4", "H5", "H6", "H7", "H8", "H9"));


        for (int i = 0; i < range.numSections(); i++) {
            Section section = range.getSection(i);
            for (int j = 0; j < section.numParagraphs(); j++) {
                Paragraph paragraph = section.getParagraph(j);
                short styleIndex = paragraph.getStyleIndex();
                StyleDescription style = docStyleSheet.getStyleDescription(styleIndex);
                String styleName = style.getName();
                if (StringUtils.isNotBlank(styleName)) {
                    // 找到对应的是几级标题
                    Matcher matcher = pattern.matcher(styleName);
                    if (matcher.find()) {
                        String head = matcher.group();
                        int headPosition = headList.indexOf(head);
                        // 标题定位设置
                        System.out.println(styleName);
                        for (int x = 0; x <= headPosition; x++) {
                            System.out.print("\t");
                        }
                        headNum[headPosition]++;
                        // 标题加上标题的数字
                        StringBuilder headId = new StringBuilder();
                        for (int x = 0; x <= headPosition; x++) {
                            System.out.print(headNum[x] + ".");
                        }

                        System.out.println(paragraph.text());
                        // 当前级别后续的子级别重置
                        for (int x = headPosition + 1; x < headLevel; x++) {
                            headNum[x] = 0;
                        }
                        continue;
                    } else {
                        // 该种方式遍历 文本会丢失
                        // styleName = "正文"; //去遍历每个字符
                    }
                }

                for (int k = 0; k < paragraph.numCharacterRuns(); k++) {
                    CharacterRun c = paragraph.getCharacterRun(k);

                    System.out.println(styleName + "i=" + i + ",j=" + j + ",k=" + k + c.text());
                }
            }
        }
    }
}
