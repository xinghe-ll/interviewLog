package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.TitleVO;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * Excel多级表头导出工具类（标题行带框线版）
 */
public class MultiLevelHeaderExcelExporter2 {

    /**
     * 导出多级表头Excel
     *
     * @param titles    表头结构定义
     * @param data      表格数据
     * @param sheetName 工作表名称
     * @return 导出的Excel字节数组
     * @throws IOException IO异常
     */
    public byte[] exportExcel(List<TitleVO> titles, List<List<Object>> data, String sheetName) throws IOException {
        // 创建工作簿
        try (Workbook workbook = new XSSFWorkbook()) {
            // 创建工作表
            Sheet sheet = workbook.createSheet(sheetName);

            // 计算表头的最大层级和总列数
            int headerMaxLevel = getHeaderMaxLevel(titles);
            int totalColumnCount = getTotalColumnCount(titles);

            // 创建表头行（使用循环方式替代递归）
            createHeaderRowsWithLoop(sheet, titles, headerMaxLevel, totalColumnCount, workbook);

            // 填充数据
            fillData(sheet, data, headerMaxLevel, totalColumnCount, workbook);

            // 优化列宽调整，根据内容长度大致估算
            optimizeColumnWidths(sheet, headerMaxLevel, totalColumnCount, data);

            // 将工作簿写入字节数组
            try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
                workbook.write(outputStream);
                return outputStream.toByteArray();
            }
        }
    }

    /**
     * 计算表头的最大层级（使用循环实现）
     */
    private int getHeaderMaxLevel(List<TitleVO> titles) {
        if (titles == null || titles.isEmpty()) {
            return 0;
        }

        int maxLevel = 0;
        Queue<Map.Entry<TitleVO, Integer>> queue = new LinkedList<>();

        for (TitleVO title : titles) {
            queue.add(new AbstractMap.SimpleEntry<>(title, 1));
        }

        while (!queue.isEmpty()) {
            Map.Entry<TitleVO, Integer> entry = queue.poll();
            TitleVO title = entry.getKey();
            int currentLevel = entry.getValue();

            if (currentLevel > maxLevel) {
                maxLevel = currentLevel;
            }

            if (title.getChildren() != null && !title.getChildren().isEmpty()) {
                for (TitleVO child : title.getChildren()) {
                    queue.add(new AbstractMap.SimpleEntry<>(child, currentLevel + 1));
                }
            }
        }

        return maxLevel;
    }

    /**
     * 计算总列数（使用循环实现）
     */
    private int getTotalColumnCount(List<TitleVO> titles) {
        if (titles == null || titles.isEmpty()) {
            return 0;
        }

        int totalCount = 0;
        Stack<TitleVO> stack = new Stack<>();

        for (int i = titles.size() - 1; i >= 0; i--) {
            stack.push(titles.get(i));
        }

        while (!stack.isEmpty()) {
            TitleVO title = stack.pop();

            if (title.getChildren() != null && !title.getChildren().isEmpty()) {
                for (int i = title.getChildren().size() - 1; i >= 0; i--) {
                    stack.push(title.getChildren().get(i));
                }
            } else {
                totalCount++;
            }
        }

        return totalCount;
    }

    /**
     * 使用循环方式创建表头行（替代递归）
     */
    private void createHeaderRowsWithLoop(Sheet sheet, List<TitleVO> titles, int maxLevel,
        int totalColumnCount, Workbook workbook) {
        CellStyle headerCellStyle = createHeaderCellStyle(workbook);
        Stack<HeaderNode> stack = new Stack<>();

        int currentCol = 0;
        for (TitleVO title : titles) {
            stack.push(new HeaderNode(title, 0, currentCol));
            currentCol += calculateColSpan(title);
        }

        while (!stack.isEmpty()) {
            HeaderNode node = stack.pop();
            TitleVO title = node.title;
            int rowIndex = node.rowIndex;
            int colIndex = node.colIndex;

            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }

            int rowSpan = calculateRowSpan(title, maxLevel - rowIndex);
            int colSpan = calculateColSpan(title);

            Cell cell = row.createCell(colIndex);
            cell.setCellValue(title.getTitleName());
            cell.setCellStyle(headerCellStyle);

            if (rowSpan > 1 || colSpan > 1) {
                CellRangeAddress region = new CellRangeAddress(
                    rowIndex,
                    rowIndex + rowSpan - 1,
                    colIndex,
                    colIndex + colSpan - 1
                );
                sheet.addMergedRegion(region);

                // 为合并单元格设置边框
                setMergedRegionBorders(region, sheet, headerCellStyle);
            }

            if (title.getChildren() != null && !title.getChildren().isEmpty()) {
                int childCol = colIndex;
                for (int i = title.getChildren().size() - 1; i >= 0; i--) {
                    TitleVO child = title.getChildren().get(i);
                    stack.push(new HeaderNode(child, rowIndex + 1, childCol));
                    childCol += calculateColSpan(child);
                }
            }
        }
    }

    /**
     * 为合并单元格设置边框
     */
    private void setMergedRegionBorders(CellRangeAddress region, Sheet sheet, CellStyle style) {
        // 顶部边框
        for (int col = region.getFirstColumn(); col <= region.getLastColumn(); col++) {
            Cell cell = getOrCreateCell(sheet, region.getFirstRow(), col);
            cell.setCellStyle(style);
        }

        // 底部边框
        for (int col = region.getFirstColumn(); col <= region.getLastColumn(); col++) {
            Cell cell = getOrCreateCell(sheet, region.getLastRow(), col);
            cell.setCellStyle(style);
        }

        // 左边框
        for (int row = region.getFirstRow(); row <= region.getLastRow(); row++) {
            Cell cell = getOrCreateCell(sheet, row, region.getFirstColumn());
            cell.setCellStyle(style);
        }

        // 右边框
        for (int row = region.getFirstRow(); row <= region.getLastRow(); row++) {
            Cell cell = getOrCreateCell(sheet, row, region.getLastColumn());
            cell.setCellStyle(style);
        }
    }

    /**
     * 获取或创建单元格
     */
    private Cell getOrCreateCell(Sheet sheet, int rowIndex, int colIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        return cell;
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
     * 计算列跨度（使用循环实现）
     */
    private int calculateColSpan(TitleVO title) {
        if (title.getChildren() == null || title.getChildren().isEmpty()) {
            return 1;
        }

        int colSpan = 0;
        Stack<TitleVO> stack = new Stack<>();
        stack.push(title);

        while (!stack.isEmpty()) {
            TitleVO current = stack.pop();

            if (current.getChildren() != null && !current.getChildren().isEmpty()) {
                for (int i = current.getChildren().size() - 1; i >= 0; i--) {
                    stack.push(current.getChildren().get(i));
                }
            } else {
                colSpan++;
            }
        }

        return colSpan;
    }

    /**
     * 优化列宽调整，根据内容长度大致估算
     */
    private void optimizeColumnWidths(Sheet sheet, int headerRowCount, int totalColumnCount, List<List<Object>> data) {
        // 基础字符宽度系数（根据经验值设定）
        final int BASE_CHAR_WIDTH = 256 * 2; // 每个字符的大致宽度

        // 先让POI自动调整一次作为基础
        for (int i = 0; i < totalColumnCount; i++) {
            sheet.autoSizeColumn(i);
        }

        // 计算每列的最大内容长度
        int[] maxContentLengths = new int[totalColumnCount];

        // 检查表头内容
        for (int row = 0; row < headerRowCount; row++) {
            Row currentRow = sheet.getRow(row);
            if (currentRow == null) continue;

            for (int col = 0; col < totalColumnCount; col++) {
                Cell cell = currentRow.getCell(col);
                if (cell == null) continue;

                String content = getCellContent(cell);
                maxContentLengths[col] = Math.max(maxContentLengths[col], content.length());
            }
        }

        // 检查数据内容
        if (data != null && !data.isEmpty()) {
            int startRow = headerRowCount;
            for (int i = 0; i < data.size(); i++) {
                List<Object> rowData = data.get(i);
                for (int j = 0; j < rowData.size() && j < totalColumnCount; j++) {
                    Object value = rowData.get(j);
                    if (value == null) continue;

                    String content = value.toString();
                    maxContentLengths[j] = Math.max(maxContentLengths[j], content.length());
                }
            }
        }

        // 根据最大长度设置列宽
        for (int col = 0; col < totalColumnCount; col++) {
            // 计算宽度：基础宽度 + 字符数 * 字符宽度 + 适当缓冲
            int estimatedWidth = BASE_CHAR_WIDTH * (maxContentLengths[col] + 2);
            // 确保不小于自动调整的宽度，不超过Excel最大限制
            int autoSizeWidth = sheet.getColumnWidth(col);
            int finalWidth = Math.max(estimatedWidth, autoSizeWidth);
            finalWidth = Math.min(finalWidth, 65535);

            sheet.setColumnWidth(col, finalWidth);
        }
    }

    /**
     * 获取单元格内容字符串
     */
    private String getCellContent(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    /**
     * 创建表头单元格样式（带框线）
     */
    private CellStyle createHeaderCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // 设置边框
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        // 设置边框颜色为黑色
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());

        // 设置对齐方式
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);

        return style;
    }

    /**
     * 填充数据到表格
     */
    private void fillData(Sheet sheet, List<List<Object>> data, int headerRowCount,
        int totalColumnCount, Workbook workbook) {
        if (data == null || data.isEmpty()) {
            return;
        }

        // 数据行不设置样式，使用默认样式
        int startRow = headerRowCount;

        for (int i = 0; i < data.size(); i++) {
            List<Object> rowData = data.get(i);
            Row row = sheet.createRow(startRow + i);

            for (int j = 0; j < totalColumnCount; j++) {
                Cell cell = row.createCell(j);

                // 不设置单元格样式，使用默认样式
                if (j < rowData.size()) {
                    Object value = rowData.get(j);
                    if (value != null) {
                        if (value instanceof Number) {
                            cell.setCellValue(((Number) value).doubleValue());
                        } else if (value instanceof Boolean) {
                            cell.setCellValue((Boolean) value);
                        } else if (value instanceof Date) {
                            cell.setCellValue((Date) value);
                        } else {
                            cell.setCellValue(value.toString());
                        }
                    }
                }
            }
        }
    }

    /**
     * 辅助类：存储表头节点信息（用于循环处理）
     */
    private static class HeaderNode {
        TitleVO title;
        int rowIndex;
        int colIndex;

        HeaderNode(TitleVO title, int rowIndex, int colIndex) {
            this.title = title;
            this.rowIndex = rowIndex;
            this.colIndex = colIndex;
        }
    }
}
