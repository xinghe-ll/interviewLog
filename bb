import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class DynamicMultiHeaderExcel {
    // 存储自动计算的单元格合并区域（修复：确保无重叠）
    private static final List<CellRangeAddress> MERGE_REGIONS = new ArrayList<>();
    // 表头总列数（最底层表头的数量）
    private static int TOTAL_COLUMN = 0;
    // 表头总行数（先通过单独递归确定，避免动态变化）
    private static int TOTAL_HEADER_ROW = 0;

    public static void main(String[] args) {
        // 1. 构造测试用的3行表头（与之前一致，确保结构正确）
        List<Title> titleList = buildTestTitleList();

        // 2. 生成Excel（修复：先算总行数，再算合并规则）
        try {
            generateExcel(titleList, "D:/MultiHeaderExcel.xlsx");
            System.out.println("Excel生成成功！");
            System.out.println("表头总行数：" + TOTAL_HEADER_ROW + "，表头总列数：" + TOTAL_COLUMN);
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Excel生成失败：" + e.getMessage());
        } finally {
            // 重置静态变量（避免多次调用污染）
            MERGE_REGIONS.clear();
            TOTAL_COLUMN = 0;
            TOTAL_HEADER_ROW = 0;
        }
    }

    /**
     * 构造测试用的3行表头（结构不变，确保层级正确）
     */
    private static List<Title> buildTestTitleList() {
        // 第3行（最底层，无child）
        Title q1 = new Title("Q1");
        Title q2 = new Title("Q2");
        Title q3 = new Title("Q3");
        Title newUserCount = new Title("新增数");
        Title oldUserCount = new Title("活跃数");

        // 第2行（child=第3行）
        Title title2024 = new Title("2024", List.of(q1, q2));
        Title title2023 = new Title("2023", List.of(q3));
        Title titleNewUser = new Title("新用户", List.of(newUserCount));
        Title titleOldUser = new Title("老用户", List.of(oldUserCount));

        // 第1行（child=第2行）
        Title salesData = new Title("销售数据", List.of(title2024, title2023));
        Title userData = new Title("用户数据", List.of(titleNewUser, titleOldUser));

        return List.of(salesData, userData);
    }

    /**
     * 核心方法：生成Excel（修复：拆分“算总行数→算合并规则→填内容”三步）
     */
    public static void generateExcel(List<Title> firstRowTitles, String outputPath) throws IOException {
        // 步骤1：先递归计算「表头总行数」和「总列数」（仅算深度和列数，不算合并规则）
        calculateHeaderDepthAndColumn(firstRowTitles, 0);

        // 步骤2：再递归计算「合并规则」（此时总行数已确定，避免范围重叠）
        calculateMergeRegions(firstRowTitles, 0, 0);

        // 步骤3：创建工作簿和工作表
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("多级表头数据");

        // 步骤4：创建表头样式
        CellStyle headerStyle = createHeaderStyle(workbook);

        // 步骤5：填充表头内容
        fillHeaderContent(sheet, firstRowTitles, 0, 0, headerStyle);

        // 步骤6：执行合并（此时合并规则无重叠）
        for (CellRangeAddress region : MERGE_REGIONS) {
            sheet.addMergedRegion(region);
        }

        // 步骤7：调整列宽
        adjustColumnWidth(sheet);

        // 步骤8：写入文件
        try (FileOutputStream fos = new FileOutputStream(outputPath)) {
            workbook.write(fos);
        } finally {
            workbook.close();
        }
    }

    /**
     * 步骤1-1：单独递归计算「表头总行数」和「总列数」（仅算深度，不涉及合并）
     * @param currentTitles 当前行表头列表
     * @param currentRow 当前行号（从0开始）
     */
    private static int calculateHeaderDepthAndColumn(List<Title> currentTitles, int currentRow) {
        // 更新总行数（当前行+1 = 已解析层级）
        TOTAL_HEADER_ROW = Math.max(TOTAL_HEADER_ROW, currentRow + 1);

        int currentCol = 0;
        for (Title title : currentTitles) {
            List<Title> childTitles = title.getChild();
            if (childTitles == null || childTitles.isEmpty()) {
                // 最底层：列数+1
                TOTAL_COLUMN = Math.max(TOTAL_COLUMN, currentCol + 1);
                currentCol++;
            } else {
                // 有子级：递归算子级的列数，累加当前列数
                int childColCount = calculateHeaderDepthAndColumn(childTitles, currentRow + 1);
                currentCol += childColCount;
            }
        }
        return currentCol;
    }

    /**
     * 步骤2-1：单独递归计算「合并规则」（总行数已确定，避免重叠）
     * @param currentTitles 当前行表头列表
     * @param currentRow 当前行号
     * @param startCol 当前行起始列号
     */
    private static int calculateMergeRegions(List<Title> currentTitles, int currentRow, int startCol) {
        int currentCol = startCol;

        for (Title title : currentTitles) {
            List<Title> childTitles = title.getChild();
            if (childTitles == null || childTitles.isEmpty()) {
                // 最底层：无合并，列数+1
                currentCol++;
            } else {
                // 有子级：先算子级的列数和合并规则
                int childColCount = calculateMergeRegions(childTitles, currentRow + 1, currentCol);

                // 计算当前表头的合并范围（关键修复：父表头合并到「总行数-1」，子表头从「当前行+1」开始）
                int parentEndRow = TOTAL_HEADER_ROW - 1; // 父表头结束行=总行数-1（固定）
                int parentEndCol = currentCol + childColCount - 1; // 父表头结束列=当前列+子列数-1
                MERGE_REG

/**
      * 步骤2-1：单独递归计算「合并规则」（总行数已确定，避免重叠）
      * @param currentTitles 当前行表头列表
      * @param currentRow 当前行号
      * @param startCol 当前行起始列号
      */
     private static int calculateMergeRegions(List<Title> currentTitles, int currentRow, int startCol) {
         int currentCol = startCol;
         for (Title title : currentTitles) {
             List<Title> childTitles = title.getChild();
             if (childTitles == null || childTitles.isEmpty()) {
                 // 最底层：无合并，列数+1
                 currentCol++;
             } else {
                 // 有子级：先算子级的列数和合并规则
                 int childColCount = calculateMergeRegions(childTitles, currentRow + 1, currentCol);
                 // 计算当前表头的合并范围（关键修复：父表头合并到「总行数-1」，子表头从「当前行+1」开始）
                 int parentEndRow = TOTAL_HEADER_ROW - 1; // 父表头结束行=总行数-1（固定）
                 int parentEndCol = currentCol + childColCount - 1; // 父表头结束列=当前列+子列数-1
                 MERGE_REGIONS.add(new CellRangeAddress(currentRow, parentEndRow, currentCol, parentEndCol));
                 currentCol += childColCount;
             }
         }
         return currentCol - startCol;
     }


/**
      * 填充表头内容（逻辑不变）
      */
     private static int fillHeaderContent(Sheet sheet, List<Title> currentTitles, int currentRow, int startCol, CellStyle style) {
         int currentCol = startCol;
         for (Title title : currentTitles) {
             List<Title> childTitles = title.getChild();
             // 创建当前行和单元格
             Row row = sheet.getRow(currentRow);
             if (row == null) row = sheet.createRow(currentRow);
             Cell cell = row.createCell(currentCol);
             cell.setCellValue(title.getTitleName());
             cell.setCellStyle(style);
             if (childTitles != null && !childTitles.isEmpty()) {
                 // 递归填充子表头
                 currentCol = fillHeaderContent(sheet, childTitles, currentRow + 1, currentCol, style);
             } else {
                 currentCol++;
             }
         }
         return currentCol;
     }



* 创建表头样式（逻辑不变）
      */
     private static CellStyle createHeaderStyle(Workbook workbook) {
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
         return style;
     }
     /**
      * 调整列宽（逻辑不变）
      */
     private static void adjustColumnWidth(Sheet sheet) {
         for (int i = 0; i < TOTAL_COLUMN; i++) {
             sheet.autoSizeColumn(i);
             int finalWidth = Math.min(sheet.getColumnWidth(i) + 1200, 65535);
             sheet.setColumnWidth(i, finalWidth);
         }
     }





/**
      * Title实体类（与之前一致，确保属性正确）
      */
     public static class Title {
         private String titleName;
         private List<Title> child;
         public Title() {}
         public Title(String titleName) {
             this.titleName = titleName;
             this.child = null;
         }
         public Title(String titleName, List<Title> child) {
             this.titleName = titleName;
             this.child = child;
         }
         // Getter + Setter
         public String getTitleName() { return titleName; }
         public void setTitleName(String titleName) { this.titleName = titleName; }
         public List<Title> getChild() { return child; }
         public void setChild(List<Title> child) { this.child = child; }