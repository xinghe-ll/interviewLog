package org.example;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Excel导出示例
 */
public class ExcelExportDemo {
    public static void main(String[] args) {
        try {
            // 1. 构建多级表头
            List<TitleVO> titles = buildMultiLevelTitles();
            
            // 2. 准备数据
            List<List<Object>> data = prepareData();
            
            // 3. 导出Excel
            MultiLevelHeaderExcelExporter exporter = new MultiLevelHeaderExcelExporter();
            byte[] excelBytes = exporter.exportExcel(titles, data, "多级表头示例");
            
            // 4. 保存到文件
            try (FileOutputStream fos = new FileOutputStream("multi_level_header_example-3.xlsx")) {
                fos.write(excelBytes);
                System.out.println("Excel导出成功！");
            }
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Excel导出失败：" + e.getMessage());
        }
    }
    
    /**
     * 构建多级表头示例
     */
    private static List<TitleVO> buildMultiLevelTitles() {
        // 构建三级表头示例
        TitleVO idTitle = new TitleVO("ID");
        
        TitleVO basicInfo = new TitleVO("基本信息");
        basicInfo.setChildren(Arrays.asList(
            new TitleVO("姓名"),
            new TitleVO("性别"),
            new TitleVO("年龄")
        ));
        
        TitleVO contactInfo = new TitleVO("联系方式");
        contactInfo.setChildren(Arrays.asList(
            new TitleVO("电话"),
            new TitleVO("邮箱")
        ));
        
        TitleVO scoreInfo = new TitleVO("成绩信息");
        TitleVO mathScore = new TitleVO("数学");
        mathScore.setChildren(Arrays.asList(
            new TitleVO("期中"),
            new TitleVO("期末")
        ));
        
        TitleVO englishScore = new TitleVO("英语");
        englishScore.setChildren(Arrays.asList(
            new TitleVO("期中"),
            new TitleVO("期末")
        ));
        
        scoreInfo.setChildren(Arrays.asList(mathScore, englishScore));
        
        return Arrays.asList(idTitle, basicInfo, contactInfo, scoreInfo);
    }
    
    /**
     * 准备示例数据
     */
    private static List<List<Object>> prepareData() {
        List<List<Object>> data = new ArrayList<>();
        
        // 注意：数据列数需要与表头总列数一致
        data.add(Arrays.asList(1, "张三", "男", 20, "13800138000", "zhangsan@example.com", 85, 92, 78, 88));
        data.add(Arrays.asList(2, "李四", "女", 21, "13900139000", "lisi@example.com", 90, 95, 82, 91));
        data.add(Arrays.asList(3, "王五", "男", 22, "13700137000", "wangwu@example.com", 78, 85, 90, 93));
        
        return data;
    }
}
