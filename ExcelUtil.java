/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2024-2024. All rights reserved.
 */

package com.ljn.demo.util;

import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.text.DecimalFormat;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

/**
 * Excel数据处理工具类
 *
 * @author zWX1244029
 * @since 2024-10-29 16:22
 */
@Slf4j
public class ExcelUtil {
    private static final DecimalFormat DECIMAL_FORMAT = new DecimalFormat("###################.###########");

    private static final Integer NUMBER_ONE = 0;

    private static final String SHEET1 = "Sheet1";

    public static final String star = "*";

    private static final int DEFAULT_COLUMN_WIDTH = 3000;

    public static final String INVALID_STR = "-1";

    private static final short DEFAULT_ROW_HEIGHT = 255; // 默认行高：12.75磅（255二十分之一磅）

    /**
     * 前1行为模板固定内容
     */
    private static final Integer FIXED_ROWS = 1;

    /**
     * excel多级表头创建
     *
     * @param workbook workbook
     * @param sheet sheet
     * @param excelMultiHeadVO excelMultiHeadVO
     */
    public static void createMultiHeader(XSSFWorkbook workbook, Sheet sheet, ExcelMultiHeadVO excelMultiHeadVO) {
        List<List<String>> titleNameList = excelMultiHeadVO.getTitleNameList();
        if (CollectionUtils.isEmpty(titleNameList)) {
            throw new IllegalArgumentException("titleNameList is null.");
        }

        // 处理合并单元格
        List<CellRangeAddress> cellRangeAddressList = excelMultiHeadVO.getCellRangeAddressList();
        if (CollectionUtils.isNotEmpty(cellRangeAddressList)) {
            for (CellRangeAddress cellAddresses : cellRangeAddressList) {
                sheet.addMergedRegion(cellAddresses);
            }
        }

        CellStyle headerStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        headerStyle.setFont(font);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        int rowNum = 0;
        for (List<String> list : titleNameList) {
            Row row = sheet.createRow(rowNum++);
            int cellNum = 0;
            for (String str : list) {
                Cell cell = row.createCell(cellNum++);
                cell.setCellValue(str);
                cell.setCellStyle(headerStyle);
            }
        }

        XSSFCellStyle stringCellStyle = workbook.createCellStyle();
        stringCellStyle.setDataFormat(workbook.createDataFormat().getFormat("@"));
        // 调整列宽
        Row row = sheet.getRow(0);
        for (int i = 0; i < row.getLastCellNum(); i++) {
            sheet.setColumnWidth(i, DEFAULT_COLUMN_WIDTH);
            sheet.setDefaultColumnStyle(i, stringCellStyle); // 设置默认'单元格格式'为<文本>
        }
    }

    /**
     * excel数据写入
     *
     * @param sheet sheet
     * @param rowDataList 行数据集(行有序&值有序)
     * @param skipRowNum 跳过的行号
     */
    public static void writeRowData(Sheet sheet, List<List<Object>> rowDataList, int skipRowNum) {
        if (CollectionUtils.isEmpty(rowDataList)) {
            return;
        }
        for (int i = 0; i < rowDataList.size(); i++) {
            Row row = sheet.createRow(i + skipRowNum);
            List<Object> rowData = rowDataList.get(i);
            for (int j = 0; j < rowData.size(); j++) {
                Cell cell = row.createCell(j);
                Object cellData = rowData.get(j);
                String cellDataStr = cellData == null ? null : cellData.toString();
                cell.setCellValue(cellDataStr);
            }
        }
    }

    @Getter
    @Setter
    @Builder
    public static class TitleStyle {
        private boolean hasStar;

        private short fontColor; // 使用IndexedColors的索引值，例如 IndexedColors.RED.getIndex();

        private short backgroundColor; // 背景颜色

        private boolean bold; // 加粗
    }

    /**
     * EXCEL模板导出(动态标题行)
     *
     * @param response HttpServletResponse
     * @param fileName 文件名称(不带后缀)
     * @param headers 标题行字段清单
     * @param headerStyles 标题行对应style
     */
    public static void exportExcelTemplate(HttpServletResponse response, String fileName, List<String> headers,
        Map<String, TitleStyle> headerStyles) {
        XSSFWorkbook workbook = null;
        ServletOutputStream out = null;
        try {
            workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet(SHEET1);
            XSSFRow headerRow = sheet.createRow(0);
            XSSFCellStyle stringCellStyle = workbook.createCellStyle();
            stringCellStyle.setDataFormat(workbook.createDataFormat().getFormat("@"));
            // 标题行初始化
            int[] headerLen = initExcelHeaders(headers, headerStyles, headerRow, workbook);
            // 可选：自动调整列宽
            for (int i = 0; i < headers.size(); i++) {
                sheet.setColumnWidth(i, headerLen[i]);
                sheet.setDefaultColumnStyle(i, stringCellStyle); // 设置默认'单元格格式'为<文本>
            }

            // 文件写出
            response.reset();
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8");
            response.setHeader("Content-Disposition",
                "attachment; filename=" + URLEncoder.encode(fileName + ".xlsx", StandardCharsets.UTF_8.toString()));
            response.setCharacterEncoding("utf-8");
            out = response.getOutputStream();
            workbook.write(out);
            out.flush();
        } catch (Exception ex) {
            log.error("ExcelUtil.exportExcelTemplate isErr, ex=", ex);
            throw new IllegalArgumentException("exportExcelTemplate isErr");
        } finally {
            FileUtil.closeQuietly(workbook, out);
        }
    }

    private static int[] initExcelHeaders(List<String> headers, Map<String, TitleStyle> headerStyles, XSSFRow headerRow,
        XSSFWorkbook workbook) {
        int[] headerLen = new int[headers.size()];
        for (int i = 0; i < headers.size(); i++) {
            String header = headers.get(i);
            XSSFCell cell = headerRow.createCell(i);

            // 构造富文本标题
            String cellValue = header;
            XSSFRichTextString richText = null;
            if (headerStyles != null && headerStyles.containsKey(header) && headerStyles.get(header).isHasStar()) {
                // 1. star星标处理
                cellValue = header + "*";
                richText = new XSSFRichTextString(cellValue);
                XSSFFont starFont = workbook.createFont();
                starFont.setColor(IndexedColors.RED.getIndex());
                richText.applyFont(header.length(), cellValue.length(), starFont);
            } else {
                richText = new XSSFRichTextString(cellValue);
            }
            cell.setCellValue(richText);
            // 计算列宽
            headerLen[i] = calculateColumnWidth(cellValue);

            if (headerStyles != null && headerStyles.containsKey(header)) {
                TitleStyle style = headerStyles.get(header);
                // 2. 字体处理
                XSSFFont font = workbook.createFont();
                if (style.getFontColor() != 0) {
                    font.setColor(style.getFontColor());
                }
                // 加粗
                if (style.isBold()) {
                    font.setBold(true);
                }
                // 3. 单元格格式
                XSSFCellStyle cellStyle = workbook.createCellStyle();
                if (style.getBackgroundColor() != 0) {
                    // 背景颜色
                    cellStyle.setFillForegroundColor(style.getBackgroundColor());
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                }
                cellStyle.setFont(font);
                cell.setCellStyle(cellStyle);
            }
        }
        return headerLen;
    }

    private static int calculateColumnWidth(String text) {
        int width = 0;
        for (char c : text.toCharArray()) {
            if (isChinese(c)) {
                width += 2 * 256; // POI使用的是1/256个字符宽度作为单位, 2是系数
            } else {
                width += 1 * 256;
            }
        }
        return Math.min(width + 128, 65535); // 128 为缓冲，65535 为最大列宽
    }

    // 判断是否是中文字符
    private static boolean isChinese(char c) {
        return c >= '\u4e00' && c <= '\u9fa5';
    }

    /**
     * 给excel列限制下拉属性，不允许输入
     *
     * @param sheet 表格
     * @param validData 限制下拉的数据
     * @param columnStart 列的index start
     * @param columnEnd 列的index end
     * @param firstRow 需要从哪行开始
     * @param lastRow 需要从哪行结束
     */
    public static void setColumnDataValid(Sheet sheet, String[] validData, int columnStart, int columnEnd, int firstRow,
        int lastRow) {
        CellRangeAddressList rangeAddressList = new CellRangeAddressList(firstRow, lastRow, columnStart, columnEnd);
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = helper.createExplicitListConstraint(validData);
        DataValidation dataValidation = helper.createValidation(constraint, rangeAddressList);
        dataValidation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        dataValidation.createErrorBox("error", "请选择正确的数据");
        dataValidation.setShowErrorBox(true);
        sheet.addValidationData(dataValidation);
    }
}
