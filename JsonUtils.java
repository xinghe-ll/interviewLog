/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2024-2025. All rights reserved.
 */

package com.ljn.demo.util;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.TypeReference;
import com.alibaba.fastjson.serializer.SerializerFeature;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * json相关工具
 */
public class JsonUtils {
    /**
     * 反序列化
     *
     * @param obj
     * @param typeReference
     * @param <T>
     * @return
     */
    public static <T> T convert(Object obj, TypeReference<T> typeReference) {
        String str = toStr(obj);
        return JSON.parseObject(str, typeReference);
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
        String str = toStr(obj);
        return JSON.parseObject(str, clazz);
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
            return JSON.toJSONString(obj, SerializerFeature.WriteMapNullValue);
        }
    }

    /**
     * 从非标准json字符串中读取键值
     *
     * @param sourceStr 非标准json字符串
     * @param key 键
     * @return 值
     */
    public static String readValue(String sourceStr, String key) {
        try {
            String patternStr = "\\\\*\"(" + key + ")\\\\*\"\\s*:\\s*\\\\*\"([^\"^\\\\]*)\\\\*\"";
            Pattern pattern = Pattern.compile(patternStr);
            Matcher matcher = pattern.matcher(sourceStr);
            if (matcher.find()) {
                return matcher.group(2);
            }
        } catch (Exception ex) {
            return null;
        }
        return null;
    }
}
