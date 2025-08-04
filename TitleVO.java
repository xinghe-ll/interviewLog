/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2024-2025. All rights reserved.
 */

package com.ljn.demo.util;

import lombok.Data;

import java.util.List;

@Data
public class TitleVO {
    private String title;
    private String field; // it字段名
    private List<TitleVO> children;
}