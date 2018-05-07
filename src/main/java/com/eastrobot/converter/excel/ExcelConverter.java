package com.eastrobot.converter.excel;

import com.eastrobot.converter.base.AbstractConverter;
import com.eastrobot.converter.util.HtmlUtil;
import com.eastrobot.converter.util.OfficeUtil;
import com.eastrobot.converter.util.ResourceUtil;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.converter.ExcelToHtmlUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.dom4j.Element;
import org.dom4j.tree.DefaultElement;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * ExcelConverter office excel转html (xls)
 * <p>
 *
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-05-02 9:10
 */
public class ExcelConverter extends AbstractConverter {

    private HSSFWorkbook workbook;

    public ExcelConverter(String inputFilePath, String outputPath) {
        super(inputFilePath, outputPath);
    }

    /**
     * 准备环境
     */
    @Override
    public ExcelConverter prepareEnv() throws Exception {
        this.workbook = OfficeUtil.loadXls(new File(this.inputFilePath));
        this.root = HtmlUtil.createHtmlDocument();

        return this;
    }

    /**
     * 转换主方法
     */
    @Override
    public void startConvert() throws Exception {
        Element html = root.addElement("html");
        Element head = html.addElement("head");
        Element body = html.addElement("body");
        HtmlUtil.charset(head);
        HtmlUtil.title(head, workbook.getSummaryInformation().getTitle() != null ? workbook.getSummaryInformation()
                .getTitle() : "");

        // 每个sheet
        Element sheet = new DefaultElement("div").addAttribute("class", "sheet-nav");
        Element sheetMenu = sheet.addElement("div").addAttribute("class", "menu");
        // sheet导航
        Element sheetNav = new DefaultElement("ul").addAttribute("id", "sheetNav");
        Element mainDiv = new DefaultElement("div").addAttribute("id", "mainDiv");
        processGeneralStyle(head);
        processWorkbook(workbook, sheetNav, mainDiv);
        sheetMenu.add(sheetNav);

        body.add(mainDiv);
        body.add(sheet);
        processScript(body);

        ResourceUtil.writeFile(this.outputPath + FilenameUtils.getBaseName(this.inputFilePath) + ".html", root.asXML());
    }

    /**
     * 处理总的css样式 base.css
     *
     * @param head head元素
     */
    private void processGeneralStyle(Element head) {
        Element css = head.addElement("link")
                .addAttribute("href", "./static/base-excel.css")
                .addAttribute("rel", "stylesheet");
    }

    /**
     * 处理excel表格
     *
     * @param workbook 当前工作簿
     * @param sheetNav 导航条
     * @param mainDiv  主内容区
     */
    private void processWorkbook(HSSFWorkbook workbook, Element sheetNav, Element mainDiv) {
        boolean isFirstSheet = true;
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            HSSFSheet sheet = workbook.getSheetAt(i);
            processAllPicture(sheet);

            // Element sheetDiv = processSheet(sheet);
            // if (StringUtils.isNotBlank(sheetDiv.getStringValue().trim())) {
            //     sheetDiv.addAttribute("id", sheet.getSheetName());
            //
            //     Element li = new DefaultElement("li").addText(sheet.getSheetName());
            //     // excel中sheet名称是唯一的
            //     sheetNav.add(li);
            //     // 默认第一个sheet激活
            //     if (isFirstSheet) {
            //         li.addAttribute("class", "menu-active");
            //         isFirstSheet = false;
            //     } else {
            //         // 其他的sheet先隐藏
            //         sheetDiv.addAttribute("style", "display:none");
            //     }
            //
            //     li.addAttribute("_id", sheet.getSheetName());
            //     mainDiv.add(sheetDiv);
            // }
        }
    }

    /**
     * 处理所有表格内嵌的图形对象
     */
    private Map<String, List<HSSFPictureData>> processAllPicture(HSSFSheet sheet) {
        Map<String, List<HSSFPictureData>> dataMap = new HashMap<String, List<HSSFPictureData>>();

        // TODO by Yogurt_lei :  判断非空
        HSSFPatriarch hssfPatriarch = sheet.getDrawingPatriarch();
        List<HSSFShape> shapeList = hssfPatriarch.getChildren();

        for (HSSFShape shape : shapeList) {
            List<HSSFPictureData> pictureDataList = null;

            if (shape instanceof HSSFPicture) {
                HSSFPicture picture = (HSSFPicture) shape;
                //获取图片数据
                HSSFPictureData pictureData = picture.getPictureData();
                //获取图片定位
                if (picture.getAnchor() instanceof HSSFClientAnchor) {
                    HSSFClientAnchor anchor = (HSSFClientAnchor) picture.getAnchor();
                    //获取图片所在行作为key值,插入图片时，默认图片只占一行的单个格子，不能超出格子边界
                    int row1 = anchor.getRow1();
                    String rowNum = String.valueOf(row1);

                    if (dataMap.get(rowNum) != null) {
                        pictureDataList = dataMap.get(rowNum);
                    } else {
                        pictureDataList = new ArrayList<HSSFPictureData>();
                    }
                    pictureDataList.add(pictureData);
                    dataMap.put(rowNum, pictureDataList);
                    // 测试部分
                    int row2 = anchor.getRow2();
                    short col1 = anchor.getCol1();
                    short col2 = anchor.getCol2();
                    int dx1 = anchor.getDx1();
                    int dx2 = anchor.getDx2();
                    int dy1 = anchor.getDy1();
                    int dy2 = anchor.getDy2();

                    System.out.println("row1: " + row1 + " , row2: " + row2 + " , col1: " + col1 + " , col2: " + col2);
                    System.out.println("dx1: " + dx1 + " , dx2: " + dx2 + " , dy1: " + dy1 + " , dy2: " + dy2);
                }
            }
        }

        System.out.println("********图片数量明细 START********");
        int t = 0;
        if (dataMap != null) {
            t = dataMap.keySet().size();
        }
        if (t > 0) {
            for (String key : dataMap.keySet()) {
                System.out.println("第 " + key + " 行， 有图片： " + dataMap.get(key).size() + " 张");
            }
        } else {
            System.out.println("Excel表中没有图片!");
        }
        System.out.println("********图片数量明细 END ********");

        return dataMap;
    }

    /**
     * 解析单个sheet
     */
    private Element processSheet(HSSFSheet sheet) {
        Element div = new DefaultElement("div");
        Element table = div.addElement("table");

        //遍历每行 输出结果
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            HSSFRow row = sheet.getRow(i);
            if (row == null)
                continue;

            Element tr = table.addElement("tr").addAttribute("style", "height:" + row.getHeight() / 20f + "pt;");
            // 记录行的跨行合并 跨列合并
            CellRangeAddress[][] mergedRanges = ExcelToHtmlUtils.buildMergedRangesMap(sheet);
            processRow(mergedRanges, row, tr);
        }

        return div;
    }

    /**
     * 解析每行
     */
    private void processRow(CellRangeAddress[][] mergedRanges, HSSFRow row, Element tr) {
        final HSSFSheet sheet = row.getSheet();

        for (int colIx = row.getFirstCellNum(); colIx < row.getLastCellNum(); colIx++) {

            CellRangeAddress range = ExcelToHtmlUtils.getMergedRange(mergedRanges, row.getRowNum(), colIx);
            if (range != null && (range.getFirstColumn() != colIx || range.getFirstRow() != row.getRowNum()))
                continue;

            HSSFCell cell = row.getCell(colIx);

            Element td = tr.addElement("td");

            // 设置是否跨行 跨列合并
            if (range != null) {
                if (range.getFirstColumn() != range.getLastColumn()) {
                    td.addAttribute("colspan", range.getLastColumn() - range.getFirstColumn() + 1 + "");
                }
                if (range.getFirstRow() != range.getLastRow()) {
                    td.addAttribute("rowspan", range.getLastRow() - range.getFirstRow() + 1 + "");
                }
            }

            if (cell != null) {
                // 获取宽度
                int columnWidth = sheet.getColumnWidth(colIx);
                int normalWidthPx = (columnWidth / 256) * 7;

                int offsetWidthUnits = columnWidth % 256;
                normalWidthPx += Math.round(offsetWidthUnits / ((float) 256 / 7));

                processCell(cell, td, normalWidthPx, row.getHeight() / 20f);
            }
        }
    }

    /**
     * 解析具体的单元格
     */
    private void processCell(HSSFCell cell, Element td, int normalWidthPx, float normalHeightPt) {

        String value;
        switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    value = String.valueOf(String.valueOf(cell.getDateCellValue()));
                } else {
                    value = String.valueOf(cell.getNumericCellValue());
                }
                break;
            case HSSFCell.CELL_TYPE_STRING:
                value = cell.getRichStringCellValue().getString();
                break;
            case HSSFCell.CELL_TYPE_FORMULA:
                FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                evaluator.evaluateFormulaCell(cell);
                CellValue cellValue = evaluator.evaluate(cell);
                value = String.valueOf(cellValue.getNumberValue());
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                value = String.valueOf(cell.getBooleanCellValue());
                break;
            case HSSFCell.CELL_TYPE_ERROR:
                value = ErrorEval.getText(cell.getErrorCellValue());
                break;
            default:
                value = "";
                break;
        }

        String tdStyle = processCellStyle(cell.getCellStyle(), cell.getSheet().getWorkbook());
        td.addAttribute("style", tdStyle);
        td.addText(value);
    }

    /**
     * 处理每个单元格的样式
     */
    private String processCellStyle(CellStyle cellStyle, HSSFWorkbook workbook) {
        StringBuilder tdStyle = new StringBuilder();
        tdStyle.append("white-space:pre-wrap;align:");
        switch (cellStyle.getAlignment()) {
            case CellStyle.ALIGN_CENTER:
                tdStyle.append("center;");
                break;
            case CellStyle.ALIGN_LEFT:
                tdStyle.append("left;");
                break;
            case CellStyle.ALIGN_RIGHT:
                tdStyle.append("right;");
                break;
            default:
                tdStyle.append("center;");
        }

        switch (cellStyle.getFillPattern()) {
            case 0:
                break;
            case 1:
                final Color foregroundColor = cellStyle.getFillForegroundColorColor();
                if (foregroundColor == null) break;
                String fgCol = ExcelToHtmlUtils.getColor(HSSFColor.toHSSFColor(foregroundColor));
                tdStyle.append("background-color:").append(fgCol).append(";");
                break;
            default:
                final Color backgroundColor = cellStyle.getFillBackgroundColorColor();
                if (backgroundColor == null) break;
                String bgCol = ExcelToHtmlUtils.getColor(HSSFColor.toHSSFColor(backgroundColor));
                tdStyle.append("background-color:").append(bgCol).append(";");
                break;
        }

        // 构建边框样式
        buildStyleBorder(workbook, tdStyle, "top", cellStyle.getBorderTop(), cellStyle.getTopBorderColor());
        buildStyleBorder(workbook, tdStyle, "right", cellStyle.getBorderRight(), cellStyle.getRightBorderColor());
        buildStyleBorder(workbook, tdStyle, "bottom", cellStyle.getBorderBottom(), cellStyle.getBottomBorderColor());
        buildStyleBorder(workbook, tdStyle, "left", cellStyle.getBorderLeft(), cellStyle.getLeftBorderColor());

        return tdStyle.toString();
    }

    // 构建单元格边框样式
    private void buildStyleBorder(HSSFWorkbook workbook, StringBuilder style, String type, short xlsBorder,
                                  short borderColor) {
        if (xlsBorder == HSSFCellStyle.BORDER_NONE) {
            return;
        }

        StringBuilder borderStyle = new StringBuilder();
        borderStyle.append(ExcelToHtmlUtils.getBorderWidth(xlsBorder));
        borderStyle.append(' ');
        borderStyle.append(ExcelToHtmlUtils.getBorderStyle(xlsBorder));

        HSSFColor color = workbook.getCustomPalette().getColor(borderColor);
        if (color != null) {
            borderStyle.append(' ');
            borderStyle.append(ExcelToHtmlUtils.getColor(color));
        }

        style.append("border-").append(type).append(":").append(borderStyle).append(";");
    }

    /**
     * 处理脚本 base.script 和 jquery-3.3.1.min.js
     */
    private void processScript(Element body) {
        Element jquery = body.addElement("script")
                .addAttribute("src", "./static/jquery-3.3.1.min.js")
                .addAttribute("type", "text/javascript")
                .addText("1");

        Element localScript = body.addElement("script")
                .addAttribute("src", "./static/base-excel.js")
                .addAttribute("type", "text/javascript")
                .addText("2");
    }
}
