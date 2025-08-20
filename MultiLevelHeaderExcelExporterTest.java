package com.ljn.demo.util;/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2024-2025. All rights reserved.
 */

import lombok.extern.slf4j.Slf4j;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

/**
 * Excel导出示例
 */
@Slf4j
public class MultiLevelHeaderExcelExporterTest {

    @Test
    public void createHeaderTest() {
        try {
            // 1. 构建多级表头
            List<TitleDTO> titles = buildTestTitleList();
            List<TitleDTO> titles3 = buildMultiLevelTitles();
            List<TitleDTO> titles4 = buildOneRowTitleList();

            // 2. 导出Excel
            try (Workbook workbook = new XSSFWorkbook();
                FileOutputStream fos = new FileOutputStream("multi_level_header_example-11.xlsx")) {
                Sheet headerSheet = workbook.createSheet("sheet0");
                MultiLevelHeaderExcelHelper exporter = new MultiLevelHeaderExcelHelper(titles);
                exporter.createHeader(headerSheet);

                workbook.write(fos);
                log.error("Excel导出成功!");
            }
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Excel导出失败：" + e.getMessage());
        }
    }

    private static List<TitleDTO> buildTestTitleList() {
        // 第3行（最底层，无child）
        TitleDTO q1 = new TitleDTO("Q1");
        TitleDTO q2 = new TitleDTO("Q2");
        TitleDTO q3 = new TitleDTO("Q3");
        TitleDTO newUserCount1 = new TitleDTO("新增数1");
        TitleDTO newUserCount2 = new TitleDTO("新增数2");
        TitleDTO oldUserCount = new TitleDTO("活跃数");

        // 第2行（child=第3行）
        TitleDTO title2024 = new TitleDTO("2024", Arrays.asList(q1, q2));
        TitleDTO title2023 = new TitleDTO("2023", Arrays.asList(q3));
        TitleDTO titleNewUser = new TitleDTO("新用户", Arrays.asList(newUserCount1, newUserCount2));
        TitleDTO titleOldUser = new TitleDTO("老用户", Arrays.asList(oldUserCount));

        // 第1行（child=第2行）
        TitleDTO salesData = new TitleDTO("销售数据", Arrays.asList(title2023, title2024));
        TitleDTO metricKK = new TitleDTO("指标kk");
        TitleDTO userData = new TitleDTO("用户数据", Arrays.asList(titleNewUser, titleOldUser));

        return Arrays.asList(salesData, metricKK, userData);
    }

    private static List<TitleDTO> buildOneRowTitleList() {
        // 第3行（最底层，无child）
        TitleDTO q1 = new TitleDTO("Q1");
        TitleDTO q2 = new TitleDTO("Q2");
        TitleDTO q3 = new TitleDTO("Q3");
        return Arrays.asList(q1, q2, q3);
    }

    /**
     * 构建多级表头示例
     */
    private static List<TitleDTO> buildMultiLevelTitles() {
        // 构建三级表头示例
        TitleDTO idTitle = new TitleDTO("ID");

        TitleDTO basicInfo = new TitleDTO("基本信息");
        basicInfo.setChildren(Arrays.asList(new TitleDTO("姓名"), new TitleDTO("性别"), new TitleDTO("年龄")));

        TitleDTO contactInfo = new TitleDTO("联系方式");
        contactInfo.setChildren(Arrays.asList(new TitleDTO("电话"), new TitleDTO("邮箱")));

        TitleDTO scoreInfo = new TitleDTO("成绩信息");
        TitleDTO mathScore = new TitleDTO("数学");
        mathScore.setChildren(Arrays.asList(new TitleDTO("期中"), new TitleDTO("期末")));

        TitleDTO englishScore = new TitleDTO("英语");
        englishScore.setChildren(Arrays.asList(new TitleDTO("期中"), new TitleDTO("期末")));

        scoreInfo.setChildren(Arrays.asList(mathScore, englishScore));

        return Arrays.asList(idTitle, basicInfo, contactInfo, scoreInfo);
    }
}