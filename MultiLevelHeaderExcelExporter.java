package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.TitleVO;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * Excel多级表头导出工具类（无框线版，修复方法引用错误）
 */
public class MultiLevelHeaderExcelExporter {

    /**
     * 导出多级表头Excel
     *
     * @param titles 表头结构定义
     * @param data 表格数据
     * @param sheetName 工作表名称
     * @return 导出的Excel字节数组
     * @throws IOException IO异常
     */
    public byte[] exportExcel(List<TitleVO> titles, List<List<Object>> data, String sheetName) throws IOException {
        // 创建工作簿
        try (Workbook workbook = new XSSFWorkbook()) {
            // 创建工作表
            Sheet sheet = workbook.createSheet(sheetName);

            // 计算表头的最大层级
            int headerMaxLevel = getHeaderMaxLevel(titles);

            // 计算总列数
            int totalColumnCount = getTotalColumnCount(titles);

            // 创建表头行
            createHeaderRows(sheet, titles, headerMaxLevel, 0, 0);

            // 填充数据
            fillData(sheet, data, headerMaxLevel, totalColumnCount);

            // 自动调整列宽
            autoSizeColumns(sheet, totalColumnCount);

            // 将工作簿写入字节数组
            try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
                workbook.write(outputStream);
                return outputStream.toByteArray();
            }
        }
    }

    /**
     * 计算表头的最大层级
     */
    private int getHeaderMaxLevel(List<TitleVO> titles) {
        int maxLevel = 1;
        for (TitleVO title : titles) {
            if (title.getChildren() != null && !title.getChildren().isEmpty()) {
                int currentLevel = 1 + getHeaderMaxLevel(title.getChildren());
                if (currentLevel > maxLevel) {
                    maxLevel = currentLevel;
                }
            }
        }
        return maxLevel;
    }

    /**
     * 计算总列数
     */
    private int getTotalColumnCount(List<TitleVO> titles) {
        int count = 0;
        for (TitleVO title : titles) {
            if (title.getChildren() != null && !title.getChildren().isEmpty()) {
                count += getTotalColumnCount(title.getChildren());
            } else {
                count++;
            }
        }
        return count;
    }

    /**
     * 创建表头行
     */
    private int createHeaderRows(Sheet sheet, List<TitleVO> titles, int maxLevel, int currentRow, int currentCol) {
        // 创建当前行
        Row row = sheet.getRow(currentRow);
        if (row == null) {
            row = sheet.createRow(currentRow);
        }

        CellStyle headerCellStyle = createHeaderCellStyle(sheet.getWorkbook());

        for (TitleVO title : titles) {
            // 计算当前标题需要跨多少行
            int rowSpan = calculateRowSpan(title, maxLevel - currentRow);

            // 计算当前标题需要跨多少列
            int colSpan = calculateColSpan(title);

            // 创建单元格
            Cell cell = row.createCell(currentCol);
            cell.setCellValue(title.getTitleName());
            cell.setCellStyle(headerCellStyle);

            // 设置合并区域
            if (rowSpan > 1 || colSpan > 1) {
                CellRangeAddress region = new CellRangeAddress(
                    currentRow,
                    currentRow + rowSpan - 1,
                    currentCol,
                    currentCol + colSpan - 1
                );
                sheet.addMergedRegion(region);
            }

            // 如果有子标题，递归创建
            if (title.getChildren() != null && !title.getChildren().isEmpty()) {
                currentCol = createHeaderRows(sheet, title.getChildren(), maxLevel, currentRow + 1, currentCol);
            } else {
                currentCol += colSpan;
            }
        }

        return currentCol;
    }

    /**
     * 计算行跨度
     */
    private int calculateRowSpan(TitleVO title, int remainingLevels) {
        if (title.getChildren() == null || title.getChildren().isEmpty()) {
            return remainingLevels;
        }
        return 1;
    }

    /**
     * 计算列跨度
     */
    private int calculateColSpan(TitleVO title) {
        if (title.getChildren() == null || title.getChildren().isEmpty()) {
            return 1;
        }
        return getTotalColumnCount(title.getChildren());
    }

    /**
     * 创建表头单元格样式（无框线）
     */
    private CellStyle createHeaderCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // 不设置任何边框

        // 设置对齐方式
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        // 设置字体
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);

        return style;
    }

    /**
     * 填充数据到表格
     */
    private void fillData(Sheet sheet, List<List<Object>> data, int headerRowCount, int totalColumnCount) {
        if (data == null || data.isEmpty()) {
            return;
        }

        CellStyle dataCellStyle = createDataCellStyle(sheet.getWorkbook());

        // 从表头下方开始填充数据
        int startRow = headerRowCount;

        for (int i = 0; i < data.size(); i++) {
            List<Object> rowData = data.get(i);
            Row row = sheet.createRow(startRow + i);

            // 使用表头总列数作为基准，确保数据列数与表头一致
            for (int j = 0; j < totalColumnCount; j++) {
                Cell cell = row.createCell(j);
                cell.setCellStyle(dataCellStyle);

                if (j < rowData.size()) {
                    Object value = rowData.get(j);
                    if (value != null) {
                        if (value instanceof Number) {
                            cell.setCellValue(((Number) value).doubleValue());
                        } else if (value instanceof Boolean) {
                            cell.setCellValue((Boolean) value);
                        } else {
                            cell.setCellValue(value.toString());
                        }
                    }
                }
            }
        }
    }

    /**
     * 从数据中获取总列数
     */
    private int getTotalColumnCountFromData(List<List<Object>> data) {
        int maxColumns = 0;
        for (List<Object> row : data) {
            if (row.size() > maxColumns) {
                maxColumns = row.size();
            }
        }
        return maxColumns;
    }

    /**
     * 创建数据单元格样式（无框线）
     */
    private CellStyle createDataCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // 不设置任何边框

        // 设置对齐方式
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        // 设置自动换行
        style.setWrapText(false);

        return style;
    }

    /**
     * 自动调整列宽
     */
    private void autoSizeColumns(Sheet sheet, int columnCount) {
        for (int i = 0; i < columnCount; i++) {
            sheet.autoSizeColumn(i);
            // 适当调整自动计算的列宽，避免过窄
            int width = sheet.getColumnWidth(i) + 256;
            sheet.setColumnWidth(i, Math.min(width, 65535)); // 最大列宽限制
        }
    }
}
