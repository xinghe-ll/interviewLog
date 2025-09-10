/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2019-2022. All rights reserved.
 */

package com.huawei.it.hap.preprocess.config;

import java.nio.charset.Charset;

import org.apache.logging.log4j.core.Layout;
import org.apache.logging.log4j.core.LogEvent;
import org.apache.logging.log4j.core.config.Node;
import org.apache.logging.log4j.core.config.plugins.Plugin;
import org.apache.logging.log4j.core.config.plugins.PluginAttribute;
import org.apache.logging.log4j.core.config.plugins.PluginElement;
import org.apache.logging.log4j.core.config.plugins.PluginFactory;
import org.apache.logging.log4j.core.layout.AbstractStringLayout;
import org.apache.logging.log4j.core.layout.PatternLayout;
import org.apache.logging.log4j.core.pattern.RegexReplacement;

/**
 * Log4j自定义Layout
 *
 * @author l00612340
 * @since 2022-08-05
 */
@Plugin(name = "CustomPatternLayout", category = Node.CATEGORY, elementType = Layout.ELEMENT_TYPE, printObject = true)
public class CustomPatternLayout extends AbstractStringLayout {
    private PatternLayout patternLayout;
    private Boolean sensitive;
    private RegexReplacement[] replaces;

    /**
     * CustomPatternLayout构造器
     *
     * @param charset charset
     * @param pattern pattern
     * @param sensitive sensitive
     * @param replaces replaces
     */
    protected CustomPatternLayout(Charset charset, String pattern, Boolean sensitive, RegexReplacement[] replaces) {
        super(charset);
        patternLayout = PatternLayout.newBuilder().withPattern(pattern).build();
        this.sensitive = sensitive;
        this.replaces = replaces;
    }

    /**
     * 插件构造工厂方法
     *
     * @param pattern   输出pattern
     * @param charset   字符集
     * @param sensitive 是否开启脱敏
     * @param replaces  脱敏规则
     * @return Layout<String>
     */
    @PluginFactory
    public static Layout<String> createLayout(
            @PluginAttribute(value = "pattern") final String pattern,
            @PluginAttribute(value = "charset", defaultString = "UTF-8") final Charset charset,
            @PluginAttribute(value = "sensitive") final Boolean sensitive,
            @PluginElement("replace") final RegexReplacement[] replaces) {
        return new CustomPatternLayout(charset, pattern, sensitive, replaces);
    }

    /**
     * toSerializable
     *
     * @param event event
     * @return 脱敏后的日志
     */
    @Override
    public String toSerializable(LogEvent event) {
        // 原日志信息
        String msg = this.patternLayout.toSerializable(event);

        if (Boolean.FALSE.equals(this.sensitive)) {
            // 不脱敏，直接返回
            return msg;
        }

        if (this.replaces == null || this.replaces.length == 0) {
            throw new IllegalArgumentException("Desensitization rule is not configured, please check the configuration and try again");
        }

        for (RegexReplacement replace : this.replaces) {
            // 遍历脱敏正则 & 替换敏感数据
            msg = replace.format(msg);
        }

        // 脱敏后的日志
        return msg;
    }
}
