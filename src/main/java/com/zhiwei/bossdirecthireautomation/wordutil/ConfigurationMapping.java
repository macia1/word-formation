package com.zhiwei.bossdirecthireautomation.wordutil;

import java.util.HashMap;
import java.util.Map;

/**
 * 映射配置
 *
 * @author fuck-aszswaz
 * @date 2021-02-20 17:34:47
 */
public class ConfigurationMapping {
    /**
     * 数字与中文字符的映射
     */
    private static final Map<Integer, String> NUMBER_MAPPING = new HashMap<>();
    /**
     * 阿拉伯数字与中文数字的权重映射
     */
    private static final Map<Integer, String> PLACE_VALUE = new HashMap<>();

    static {
        NUMBER_MAPPING.put(0, "零");
        NUMBER_MAPPING.put(1, "一");
        NUMBER_MAPPING.put(2, "二");
        NUMBER_MAPPING.put(3, "三");
        NUMBER_MAPPING.put(4, "四");
        NUMBER_MAPPING.put(5, "五");
        NUMBER_MAPPING.put(6, "六");
        NUMBER_MAPPING.put(7, "七");
        NUMBER_MAPPING.put(8, "八");
        NUMBER_MAPPING.put(9, "九");

        PLACE_VALUE.put(1, "");// 个位
        PLACE_VALUE.put(2, "十");// 十位
        PLACE_VALUE.put(3, "百");// 百位
        PLACE_VALUE.put(4, "千");// 千位
        PLACE_VALUE.put(5, "万");// 万位
        PLACE_VALUE.put(6, "十");// 十万位
        PLACE_VALUE.put(7, "百");// 百万位
        PLACE_VALUE.put(8, "千");// 千万位
        PLACE_VALUE.put(9, "亿");// 亿位
    }

    public static String getNumberMapping(int number) {
        return NUMBER_MAPPING.get(number);
    }

    public static String getPlaceValue(int placeValue) {
        return PLACE_VALUE.get(placeValue);
    }
}
