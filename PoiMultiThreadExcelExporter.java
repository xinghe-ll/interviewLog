package org.example;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.Future;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.ThreadFactory;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;
import java.util.function.BiConsumer;
import java.util.function.Function;

/**
 * POI多线程Excel导出通用工具类
 * 特点：
 * 1. 多线程并行处理不同Sheet，每个线程独立操作自己的Workbook
 * 2. 支持大数据量导出，采用SXSSFWorkbook减少内存占用
 * 3. 提供灵活的样式定制和数据映射接口
 * 4. 自动清理临时文件，避免磁盘空间占用
 * 5. 跨平台兼容Windows和Linux系统
 */
public class PoiMultiThreadExcelExporter<T> {
    // 内存中保留的行数，超过则写入临时文件
    private static final int ROW_ACCESS_WINDOW_SIZE = 1000;

    // 默认线程池大小
    private static final int DEFAULT_THREAD_POOL_SIZE = Math.min(Runtime.getRuntime().availableProcessors() + 1, 10);

    // 临时文件前缀
    private static final String TEMP_FILE_PREFIX = "poi_excel_temp_";

    // 临时文件后缀
    private static final String TEMP_FILE_SUFFIX = ".xlsx";

    private final String[] headers;

    private final Function<T, Object[]> dataMapper;

    private final String baseSheetName;

    private final int threadPoolSize;

    private final BiConsumer<Sheet, Integer> headerStyleConsumer;

    private final BiConsumer<Cell, Integer> cellStyleConsumer;

    /**
     * 构造函数
     *
     * @param headers 表头数组
     * @param dataMapper 数据映射函数，将T类型转换为Object数组
     * @param baseSheetName Sheet基础名称
     */
    public PoiMultiThreadExcelExporter(String[] headers, Function<T, Object[]> dataMapper, String baseSheetName) {
        this(headers, dataMapper, baseSheetName, DEFAULT_THREAD_POOL_SIZE, null, null);
    }

    /**
     * 构造函数（全参数）
     *
     * @param headers 表头数组
     * @param dataMapper 数据映射函数
     * @param baseSheetName Sheet基础名称
     * @param threadPoolSize 线程池大小
     * @param headerStyleConsumer 表头样式消费者
     * @param cellStyleConsumer 单元格样式消费者
     */
    public PoiMultiThreadExcelExporter(String[] headers, Function<T, Object[]> dataMapper, String baseSheetName,
        int threadPoolSize, BiConsumer<Sheet, Integer> headerStyleConsumer,
        BiConsumer<Cell, Integer> cellStyleConsumer) {
        Objects.requireNonNull(headers, "表头不能为空");
        Objects.requireNonNull(dataMapper, "数据映射函数不能为空");
        Objects.requireNonNull(baseSheetName, "Sheet基础名称不能为空");

        this.headers = Arrays.copyOf(headers, headers.length);
        this.dataMapper = dataMapper;
        this.baseSheetName = baseSheetName;
        this.threadPoolSize = Math.max(1, threadPoolSize);
        this.headerStyleConsumer = headerStyleConsumer != null ? headerStyleConsumer : this::defaultHeaderStyle;
        this.cellStyleConsumer = cellStyleConsumer != null ? cellStyleConsumer : this::defaultCellStyle;
    }

    /**
     * 导出Excel文件
     *
     * @param dataList 完整数据列表
     * @param filePath 目标文件路径
     * @param sheetDataSize 每个Sheet的数据量
     * @throws Exception 可能抛出的异常
     */
    public void export(List<T> dataList, String filePath, int sheetDataSize) throws Exception {
        if (dataList == null || dataList.isEmpty()) {
            throw new IllegalArgumentException("数据列表不能为空");
        }
        if (sheetDataSize <= 0) {
            throw new IllegalArgumentException("每个Sheet的数据量必须大于0");
        }

        // 数据分片
        List<List<T>> dataChunks = splitDataIntoChunks(dataList, sheetDataSize);
        int chunkCount = dataChunks.size();

        ThreadPoolExecutor executor = new ThreadPoolExecutor(threadPoolSize, threadPoolSize, 0L, TimeUnit.MILLISECONDS,
            new LinkedBlockingQueue<>(), new ThreadFactory() {
            private int counter = 0;

            @Override
            public Thread newThread(Runnable r) {
                Thread thread = new Thread(r, "excel-export-thread-" + counter++);
                thread.setDaemon(true);
                return thread;
            }
        }, new ThreadPoolExecutor.CallerRunsPolicy() // 当线程池满时，让提交任务的线程执行任务
        );

        // 创建线程池
        try {
            // 提交任务
            List<Future<File>> futures = new ArrayList<>(chunkCount);
            for (int i = 0; i < chunkCount; i++) {
                final int sheetIndex = i;
                List<T> chunk = dataChunks.get(i);
                futures.add(executor.submit(() -> createSheetTempFile(chunk, sheetIndex)));
            }

            // 收集临时文件并合并
            List<File> tempFiles = new ArrayList<>(chunkCount);
            try {
                for (Future<File> future : futures) {
                    tempFiles.add(future.get());
                }
                mergeTempFiles(tempFiles, filePath);
            } finally {
                // 确保临时文件被清理
                cleanupTempFiles(tempFiles);
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    /**
     * 每个线程创建一个包含单个Sheet的临时文件
     */
    private File createSheetTempFile(List<T> dataChunk, int sheetIndex) throws IOException {
        // 创建临时文件，使用try-with-resources确保资源释放
        File tempFile = File.createTempFile(TEMP_FILE_PREFIX + sheetIndex + "_", TEMP_FILE_SUFFIX);
        // 注册JVM退出时删除临时文件（作为最后保障）
        tempFile.deleteOnExit();

        // 使用SXSSFWorkbook处理大数据
        try (SXSSFWorkbook workbook = new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE)) {
            // 创建Sheet
            String sheetName = baseSheetName + (sheetIndex > 0 ? "_" + (sheetIndex + 1) : "");
            Sheet sheet = workbook.createSheet(sheetName);

            // 创建表头
            createHeader(sheet, workbook);

            // 写入数据
            writeSheetData(sheet, dataChunk);

            // 调整列宽
            autoSizeColumns(sheet);

            // 写入临时文件
            try (FileOutputStream fos = new FileOutputStream(tempFile)) {
                workbook.write(fos);
            }

            // 清理SXSSF的临时文件
            workbook.dispose();
        }

        return tempFile;
    }

    /**
     * 创建表头
     */
    private void createHeader(Sheet sheet, Workbook workbook) {
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            // 应用表头样式
            headerStyleConsumer.accept(sheet, i);
        }
    }

    /**
     * 写入Sheet数据
     */
    private void writeSheetData(Sheet sheet, List<T> dataChunk) {
        for (int i = 0; i < dataChunk.size(); i++) {
            Row row = sheet.createRow(i + 1); // 从1开始，0是表头
            T data = dataChunk.get(i);
            Object[] values = dataMapper.apply(data);

            if (values != null) {
                for (int j = 0; j < values.length && j < headers.length; j++) {
                    Cell cell = row.createCell(j);
                    setCellValue(cell, values[j]);
                    // 应用单元格样式
                    cellStyleConsumer.accept(cell, j);
                }
            }
        }
    }

    /**
     * 合并所有临时文件到最终Excel
     */
    private void mergeTempFiles(List<File> tempFiles, String targetFilePath) throws IOException {
        try (SXSSFWorkbook targetWorkbook = new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE)) {
            // 用于复制样式的映射表
            Map<Integer, CellStyle> styleMap = new HashMap<>();

            for (File tempFile : tempFiles) {
                try (InputStream is = new FileInputStream(tempFile); XSSFWorkbook tempWorkbook = new XSSFWorkbook(is)) {

                    // 复制临时文件中的第一个Sheet
                    Sheet tempSheet = tempWorkbook.getSheetAt(0);
                    Sheet targetSheet = targetWorkbook.createSheet(tempSheet.getSheetName());

                    // 复制Sheet内容
                    copySheetContent(tempWorkbook, targetWorkbook, tempSheet, targetSheet, styleMap);
                }
            }

            // 写入最终文件
            try (FileOutputStream fos = new FileOutputStream(targetFilePath)) {
                targetWorkbook.write(fos);
            }

            // 清理临时资源
            targetWorkbook.dispose();
        }
    }

    /**
     * 复制Sheet内容（包括数据和样式）
     */
    private void copySheetContent(XSSFWorkbook srcWorkbook, SXSSFWorkbook destWorkbook, Sheet srcSheet, Sheet destSheet,
        Map<Integer, CellStyle> styleMap) {
        // 复制列宽
        for (int i = 0; i < srcSheet.getRow(0).getLastCellNum(); i++) {
            destSheet.setColumnWidth(i, srcSheet.getColumnWidth(i));
        }

        // 复制行和单元格
        for (int rowNum = 0; rowNum <= srcSheet.getLastRowNum(); rowNum++) {
            Row srcRow = srcSheet.getRow(rowNum);
            if (srcRow == null) {
                continue;
            }

            Row destRow = destSheet.createRow(rowNum);

            for (int cellNum = 0; cellNum < srcRow.getLastCellNum(); cellNum++) {
                Cell srcCell = srcRow.getCell(cellNum);
                if (srcCell == null) {
                    continue;
                }

                Cell destCell = destRow.createCell(cellNum);
                // 复制单元格值
                setCellValue(destCell, getCellValue(srcCell));
                // 复制单元格样式
                copyCellStyle(srcWorkbook, destWorkbook, srcCell, destCell, styleMap);
            }
        }
    }

    /**
     * 复制单元格样式
     */
    private void copyCellStyle(XSSFWorkbook srcWorkbook, SXSSFWorkbook destWorkbook, Cell srcCell, Cell destCell,
        Map<Integer, CellStyle> styleMap) {
        CellStyle srcStyle = srcCell.getCellStyle();
        int styleIndex = srcStyle.getIndex();

        CellStyle destStyle = styleMap.get(styleIndex);
        if (destStyle == null) {
            destStyle = destWorkbook.createCellStyle();
            destStyle.cloneStyleFrom(srcStyle);
            styleMap.put(styleIndex, destStyle);
        }

        destCell.setCellStyle(destStyle);
    }

    /**
     * 获取单元格值
     */
    private Object getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                return cell.getCellFormula();
            default:
                return null;
        }
    }

    /**
     * 设置单元格值
     */
    private void setCellValue(Cell cell, Object value) {
        if (value == null) {
            cell.setCellValue("");
            return;
        }

        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    /**
     * 自动调整列宽
     */
    private void autoSizeColumns(Sheet sheet) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) {
            return;
        }

        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            sheet.autoSizeColumn(i);
            // 适当增加宽度，避免内容被截断
            int width = sheet.getColumnWidth(i) + 256;
            sheet.setColumnWidth(i, Math.min(width, 65535)); // 最大列宽限制
        }
    }

    /**
     * 数据分片
     */
    private List<List<T>> splitDataIntoChunks(List<T> dataList, int chunkSize) {
        List<List<T>> chunks = new ArrayList<>();
        int totalSize = dataList.size();
        int index = 0;

        while (index < totalSize) {
            int end = Math.min(index + chunkSize, totalSize);
            chunks.add(new ArrayList<>(dataList.subList(index, end)));
            index = end;
        }

        return chunks;
    }

    /**
     * 清理临时文件
     */
    private void cleanupTempFiles(List<File> tempFiles) {
        if (tempFiles == null) {
            return;
        }

        for (File file : tempFiles) {
            if (file != null && file.exists() && !file.delete()) {
                // 如果删除失败，标记为JVM退出时删除
                file.deleteOnExit();
            }
        }
    }

    /**
     * 默认表头样式
     */
    private void defaultHeaderStyle(Sheet sheet, int columnIndex) {
        Workbook workbook = sheet.getWorkbook();
        Cell cell = sheet.getRow(0).getCell(columnIndex);

        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        cell.setCellStyle(style);
    }

    /**
     * 默认单元格样式
     */
    private void defaultCellStyle(Cell cell, int columnIndex) {
        Workbook workbook = cell.getSheet().getWorkbook();
        CellStyle style = workbook.createCellStyle();

        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        cell.setCellStyle(style);
    }
}
