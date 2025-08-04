package com.ljn.demo.util;

import com.alibaba.fastjson.TypeReference;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class MultiLevelHeaderExcel {

    /**
     * 构造excel多级表头vo对象
     * 注: titles有序
     *
     * @param titles titles
     * @return ExcelMultiHeadVO
     */
    private static ExcelMultiHeadVO buildExcelMultiHeadVO(List<TitleVO> titles) {
        List<List<String>> titleNameList = new ArrayList<>();
        List<String> header1 = new ArrayList<>(); // 标题行1
        List<String> header2 = new ArrayList<>(); // 标题行2
        titleNameList.add(header1);
        titleNameList.add(header2);
        int rowNum1 = 0;
        int rowNum2 = 1;

        // 总数据行: it字段(可能重复) + 列索引(唯一)
        List<ImportDataHeaderVO> headerVOList = new ArrayList<>();

        List<CellRangeAddress> cellRangeAddresseList = new ArrayList<>();
        int columnIdx = 0;
        for (int i = 0; i < titles.size(); i++) {
            TitleVO titleVO = titles.get(i);

            List<TitleVO> children = titleVO.getChildren();
            if (CollectionUtils.isNotEmpty(children)) { // 存在子标题
                for (TitleVO childTitle : children) {
                    header1.add(titleVO.getTitle());
                    header2.add(childTitle.getTitle());
                    headerVOList.add(initImportDataHeaderVO(childTitle, columnIdx++));
                }
                // 处理单元格合并情况
                int start = columnIdx - 1 - (children.size() - 1);
                int end = columnIdx - 1;
                cellRangeAddresseList.add(new CellRangeAddress(rowNum1, rowNum1, start, end));
            } else { // 没有子标题
                header1.add(titleVO.getTitle());
                header2.add("-1");
                // 处理单元格合并情况
                cellRangeAddresseList.add(new CellRangeAddress(rowNum1, rowNum2, columnIdx, columnIdx));
                headerVOList.add(initImportDataHeaderVO(titleVO, columnIdx++));
            }
        }

        ExcelMultiHeadVO excelMultiHeadVO = new ExcelMultiHeadVO();
        excelMultiHeadVO.setTitleNameList(titleNameList);
        excelMultiHeadVO.setCellRangeAddressList(cellRangeAddresseList);
        excelMultiHeadVO.setHeaderVOList(headerVOList);
        return excelMultiHeadVO;
    }

    private static ImportDataHeaderVO initImportDataHeaderVO(TitleVO titleVO, int columnIndex) {
        ImportDataHeaderVO headerVO = new ImportDataHeaderVO();
        headerVO.setDisplayName(titleVO.getTitle());
        headerVO.setFieldName(titleVO.getField());
        headerVO.setColumnIndex(columnIndex);
        return headerVO;
    }

    public static void main(String[] args) throws IOException {
        // 创建工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 创建工作表
        Sheet sheet = workbook.createSheet("多级表头示例");

        String fileName = "multi_level_header-1.xlsx";

        String strJson
            = "[{\"title\":\"国家\",\"field\":\"country_cn_name\"},{\"title\":\"kkk\",\"field\":\"combination0\",\"children\":[{\"title\":\"kkk-国家\",\"field\":\"country_cn_name\"},{\"title\":\"kkk-城市\",\"field\":\"city\"},{\"title\":\"kkk-城市2\",\"field\":\"city2\"}]},{\"title\":\"城市\",\"field\":\"city\"},{\"title\":\"ggg\",\"field\":\"combination1\",\"children\":[{\"title\":\"ggg-国家\",\"field\":\"country_cn_name\"},{\"title\":\"ggg-城市\",\"field\":\"city\"}]}]";
        List<TitleVO> titles = JsonUtils.convert(strJson, new TitleVOListType());
        ExcelMultiHeadVO excelMultiHeadVO = buildExcelMultiHeadVO(titles);
        System.out.println("excelMultiHeadVO = " + excelMultiHeadVO);

        ExcelUtil.createMultiHeader(workbook, sheet, excelMultiHeadVO);

        // 保存文件
        try (FileOutputStream fos = new FileOutputStream(fileName)) {
            workbook.write(fos);
        } finally {
            workbook.close();
        }
    }

    public static class TitleVOListType extends TypeReference<List<TitleVO>> { }
}