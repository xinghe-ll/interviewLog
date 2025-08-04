/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2024-2024. All rights reserved.
 */

package com.ljn.demo.util;

import com.fasterxml.jackson.annotation.JsonIgnore;

import lombok.Data;
import lombok.EqualsAndHashCode;

import java.io.Serializable;

@Data
@EqualsAndHashCode
public class ImportDataHeaderVO implements Serializable {
    private static final long serialVersionUID = 3529622431814903535L;

    private String displayName;

    private String fieldName;

    @JsonIgnore
    private Integer columnIndex;
}
