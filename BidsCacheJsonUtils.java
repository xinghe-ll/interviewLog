/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2024-2025. All rights reserved.
 */

package com.huawei.fin.bfd.goams.domain.common.util;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import lombok.extern.slf4j.Slf4j;

/**
 * json相关工具
 */
@Slf4j
public class BidsCacheJsonUtils {
    private static final ThreadLocal<ObjectMapper> OBJECT_MAPPER_THREAD_LOCAL = ThreadLocal.withInitial(() -> {
        ObjectMapper objectMapper = new ObjectMapper();
        // json转对象时，json里有对象不存在的属性时，会把对象里有的属性复制过去，没有的忽略，不报错
        objectMapper.configure(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES, false);
        // 数值类型不识别, 保持原样
        objectMapper.configure(DeserializationFeature.USE_BIG_DECIMAL_FOR_FLOATS, false);
        objectMapper.configure(SerializationFeature.WRITE_ENUMS_USING_TO_STRING, false);
        return objectMapper;
    });

    /**
     * 反序列化
     *
     * @param obj
     * @param typeReference
     * @param <T>
     * @return
     */
    public static <T> T convert(Object obj, TypeReference<T> typeReference) {
        try {
            ObjectMapper objectMapper = OBJECT_MAPPER_THREAD_LOCAL.get();
            return objectMapper.readValue(toStr(obj), typeReference);
        } catch (JsonProcessingException ex) {
            log.error("BidsCacheJsonUtils.convert isErr, ex=", ex);
        }
        return null;
    }

    /**
     * 反序列化
     *
     * @param obj
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> T convert(Object obj, Class<T> clazz) {
        try {
            ObjectMapper objectMapper = OBJECT_MAPPER_THREAD_LOCAL.get();
            return objectMapper.readValue(toStr(obj), clazz);
        } catch (JsonProcessingException ex) {
            log.error("BidsCacheJsonUtils.convert isErr, ex=", ex);
        }
        return null;
    }

    /**
     * 转jsonStr
     *
     * @param obj obj
     * @return String
     */
    public static String toStr(Object obj) {
        if (obj instanceof String) {
            return obj.toString();
        } else {
            try {
                return OBJECT_MAPPER_THREAD_LOCAL.get().writeValueAsString(obj);
            } catch (JsonProcessingException ex) {
                log.error("BidsCacheJsonUtils.toStr isErr, ex=", ex);
            }
        }
        return null;
    }
}
