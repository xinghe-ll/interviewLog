/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2024-2025. All rights reserved.
 */

package com.ljn.demo.util;

import lombok.Data;

import java.util.List;

/**
 * 标题行对象
 */
@Data
public class TitleDTO {
    private String title; // 标题显示名称

    private String field; // it字段名

    private List<TitleDTO> children; // 孩子节点

    public TitleDTO() {
    }

    public TitleDTO(String title, String field, List<TitleDTO> children) {
        this.title = title;
        this.field = field;
        this.children = children;
    }

    public TitleDTO(String title, List<TitleDTO> children) {
        this.title = title;
        this.children = children;
    }

    public TitleDTO(String title) {
        this.title = title;
    }

    public TitleDTO(String title, String field) {
        this.title = title;
        this.field = field;
    }
}