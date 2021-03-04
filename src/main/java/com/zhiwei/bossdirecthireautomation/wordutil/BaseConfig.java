package com.zhiwei.bossdirecthireautomation.wordutil;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;

/**
 * 基本配置
 *
 * @author fuck-aszswaz
 * @date 2021-02-22 14:00:42
 */
public class BaseConfig {
    /**
     * 默认字体
     */
    public static final String DEFAULT_FONT = "微软雅黑";
    /**
     * 默认大小
     */
    public static final int DEFAULT_FONT_SIZE = ContentFont.小四.getPoundValue();
    /**
     * Boss直聘特殊情感
     */
    public static final String BOSS_EXCEPTION_EMOTION = "负面+敏感";
    /**
     * 特权媒体
     */
    public static final List<String> priorities = Collections.unmodifiableList(
            Arrays.asList("微信", "网媒")
    );

    /**
     * 品牌顺序
     */
    public static final List<String> brandShort = Collections.unmodifiableList(
            Arrays.asList("Boss直聘", "智联招聘", "58同城", "前程无忧/51job", "猎聘为顺序排放")
    );
}
