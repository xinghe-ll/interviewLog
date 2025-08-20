package com.ljn.demo.util;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.util.AbstractMap;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Queue;
import java.util.Stack;

/**
 * Excel多级表头导出工具类
 */
@Slf4j
@Data
public class MultiLevelHeaderExcelHelper {
    // 标题行
    private List<TitleDTO> titles;

    // 排序的IT字段名(与标题行的最低层级对应, 用于映射数据行)
    private List<TitleDTO> sortedTitles = new ArrayList<>();

    // 标题行最大层级
    private int headerMaxLevel;

    // 总列数
    private int totalColumnCount;

    // 标题的单元格格式
    private CellStyle headerCellStyle;

    // 标题行是否带框线(默认:true)
    private boolean hasBorder = true;

    private boolean isInitCellTextFlag = true;

    public MultiLevelHeaderExcelHelper() {
    }

    public MultiLevelHeaderExcelHelper(List<TitleDTO> titles) {
        this.titles = titles;
        this.headerMaxLevel = getHeaderMaxLevel(titles);
        this.totalColumnCount = getTotalColumnCount(titles);
    }

    public void createHeader(Sheet sheet) throws IOException {
        try {
            // 创建表头行
            createHeaderRows(sheet, titles);

            // 设置单元格格式
            setCellStyle(sheet);

            // 优化列宽调整
            optimizeColumnWidths(sheet, headerMaxLevel, totalColumnCount);
        } catch (Exception ex) {
            log.error("MultiLevelHeaderExcelExporter3.exportExcel isErr, ex=", ex);
        }
    }

    /**
     * 使用循环方式创建表头行（替代递归）
     */
    private void createHeaderRows(Sheet sheet, List<TitleDTO> titles) {
        CellStyle headerCellStyle = createHeaderCellStyle(sheet.getWorkbook());

        Stack<HeaderNode> stack = new Stack<>();
        int currentCol = 0;
        for (TitleDTO title : titles) {
            stack.push(new HeaderNode(title, 0, currentCol));
            currentCol += calcColSpan(title);
        }

        while (!stack.isEmpty()) {
            HeaderNode node = stack.pop();
            TitleDTO title = node.title;
            int rowIndex = node.rowIndex;
            int colIndex = node.colIndex;

            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }
            // 设置单元格
            Cell cell = row.createCell(colIndex);
            cell.setCellValue(title.getTitle());
            cell.setCellStyle(headerCellStyle);

            // 行跨度
            int rowSpan = calcRowSpan(title, headerMaxLevel - rowIndex);
            // 列跨度
            int colSpan = calcColSpan(title);
            if (rowSpan > 1 || colSpan > 1) {
                int lastRow = rowIndex + rowSpan - 1;
                int lastCol = colIndex + colSpan - 1;
                CellRangeAddress region = new CellRangeAddress(rowIndex, lastRow, colIndex, lastCol);
                sheet.addMergedRegion(region);

                // 为合并单元格设置边框
                setMergedRegionBorders(region, sheet, headerCellStyle);
            }

            // 处理孩子节点
            List<TitleDTO> children = title.getChildren();
            if (CollectionUtils.isNotEmpty(children)) {
                int childCol = colIndex;
                for (TitleDTO child : children) {
                    stack.push(new HeaderNode(child, rowIndex + 1, childCol));
                    childCol += calcColSpan(child); // 每遍历一个 ->其'colIndex'需要加上'前列跨度'
                }
            }
        }
    }

    private void setCellStyle(Sheet sheet) {
        Workbook workbook = sheet.getWorkbook();
        CellStyle cellStyle = workbook.createCellStyle();// '常规'格式
        if (isInitCellTextFlag) {
            cellStyle.setDataFormat(workbook.createDataFormat().getFormat("@")); // '文本'格式
        }
        for (int i = 0; i < totalColumnCount; i++) {
            if (cellStyle != null) {
                sheet.setDefaultColumnStyle(i, cellStyle); // 设置默认'单元格格式'为<文本>
            }
        }
    }

    /**
     * 计算表头的最大层级
     */
    private int getHeaderMaxLevel(List<TitleDTO> titles) {
        if (CollectionUtils.isEmpty(titles)) {
            return 0;
        }

        int maxLevel = 0;
        Queue<Map.Entry<TitleDTO, Integer>> queue = new LinkedList<>();

        for (TitleDTO title : titles) {
            queue.add(new AbstractMap.SimpleEntry<>(title, 1)); // 第一层的深度就等于: 1
        }

        while (!queue.isEmpty()) {
            Map.Entry<TitleDTO, Integer> entry = queue.poll();
            TitleDTO title = entry.getKey();
            int currentLevel = entry.getValue();

            maxLevel = Math.max(maxLevel, currentLevel);

            List<TitleDTO> children;
            if (CollectionUtils.isNotEmpty(children = title.getChildren())) {
                for (TitleDTO child : children) { // 孩子节点level == 父节点level + 1
                    queue.add(new AbstractMap.SimpleEntry<>(child, currentLevel + 1));
                }
            }
        }

        return maxLevel;
    }

    /**
     * 计算总列数
     */
    private int getTotalColumnCount(List<TitleDTO> titles) {
        if (CollectionUtils.isEmpty(titles)) {
            return 0;
        }

        int totalCount = 0;
        Stack<TitleDTO> stack = new Stack<>();

        for (int i = titles.size() - 1; i >= 0; i--) {
            stack.push(titles.get(i));
        }

        while (!stack.isEmpty()) {
            TitleDTO currTitle = stack.pop();

            List<TitleDTO> children;
            if (CollectionUtils.isNotEmpty(children = currTitle.getChildren())) {
                for (int i = children.size() - 1; i >= 0; i--) {
                    stack.push(children.get(i));
                }
            } else {
                totalCount++;
                sortedTitles.add(new TitleDTO(currTitle.getTitle(), currTitle.getField()));
            }
        }
        return totalCount;
    }

    /**
     * 为合并单元格设置边框
     */
    private void setMergedRegionBorders(CellRangeAddress region, Sheet sheet, CellStyle style) {
        if (!hasBorder) {
            return;
        }

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
    private int calcRowSpan(TitleDTO title, int remainingLevels) {
        if (CollectionUtils.isEmpty(title.getChildren())) {
            return remainingLevels;
        }
        return 1; // 跨度==1 -> 表示不需要合并
    }

    /**
     * 计算列跨度
     */
    private int calcColSpan(TitleDTO title) {
        if (title.getChildren() == null || title.getChildren().isEmpty()) {
            return 1; // 跨度==1 -> 表示不需要合并
        }

        int colSpan = 0;
        Stack<TitleDTO> stack = new Stack<>();
        stack.push(title);

        while (!stack.isEmpty()) {
            TitleDTO current = stack.pop();

            List<TitleDTO> children;
            if (CollectionUtils.isNotEmpty(children = current.getChildren())) {
                children.forEach(stack::push);
            } else {
                colSpan++;
            }
        }
        return colSpan;
    }

    /**
     * 优化列宽调整，根据内容长度大致估算
     */
    private void optimizeColumnWidths(Sheet sheet, int headerRowCount, int totalColumnCount) {
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
            if (currentRow == null) {
                continue;
            }

            for (int col = 0; col < totalColumnCount; col++) {
                Cell cell = currentRow.getCell(col);
                if (cell == null) {
                    continue;
                }

                String content = getCellContent(cell);
                maxContentLengths[col] = Math.max(maxContentLengths[col], content.length());
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
        if (cell == null) {
            return "";
        }

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
        if (this.headerCellStyle != null) {
            return headerCellStyle;
        }

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
    private void fillData(Sheet sheet, List<List<Object>> data, int headerRowCount, int totalColumnCount,
        Workbook workbook) {
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
        TitleDTO title;

        int rowIndex;

        int colIndex;

        HeaderNode(TitleDTO title, int rowIndex, int colIndex) {
            this.title = title;
            this.rowIndex = rowIndex;
            this.colIndex = colIndex;
        }
    }
}