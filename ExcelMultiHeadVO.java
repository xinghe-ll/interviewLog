/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2024-2025. All rights reserved.
 */

package com.ljn.demo.util;

import lombok.Data;

import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

/**
 * excel多级表头VO类
 */
@Data
public class ExcelMultiHeadVO {
    private List<ImportDataHeaderVO> headerVOList; // 表头字段有序清单(columnIndex唯一)

    private List<List<String>> titleNameList; // 多级表头名称list

    private List<CellRangeAddress> cellRangeAddressList; // 单元格合并清单
}
