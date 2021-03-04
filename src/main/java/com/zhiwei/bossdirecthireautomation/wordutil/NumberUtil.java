package com.zhiwei.bossdirecthireautomation.wordutil;

import lombok.Setter;

import java.util.Arrays;

/**
 * 数字工具类
 *
 * @author fuck-aszswaz
 * @date 2021-02-20 17:48:28
 */
public class NumberUtil {
    @Setter
    private int value;

    public NumberUtil() {
    }

    public NumberUtil(int value) {
        this.value = value;
    }

    /**
     * 转换为中文数字
     */
    public String toChineseNumber() {
        int[] ints = this.toArray();// 获得int每一位的数字
        StringBuilder builder = new StringBuilder();
        for (int i = ints.length - 1; i >= 0; i--) {
            int placeValue = i + 1;// 计算位值
            // 不为零的个位数
            if (i == 0 && ints[i] != 0) {
                builder.append(ConfigurationMapping.getNumberMapping(ints[0]));
                break;
            } else if (ints.length == 1) {
                builder.append(ConfigurationMapping.getNumberMapping(ints[i]));
                break;
            }
            // 如果数不为个位数, 跳过末尾的零
            if (ints.length > 1 && i == 0) {
                continue;
            }
            /*
            1. 不能出现连续的零, 零只能有一个
            2. 权重是万位或亿位的, 需要添加"亿"或"万"
             */
            if (ints[i - 1] == 0 && ints[i] == 0 && placeValue != 5 && placeValue != 9) {
                continue;
            } else if (ints[i] == 0 && placeValue != 5 && placeValue != 9) {
                // 添加不在万位或亿位的中间的零
                builder.append(ConfigurationMapping.getNumberMapping(ints[i]));
            } else if (ints[i] != 0) {
                // 确保非零的数字被添加
                builder.append(ConfigurationMapping.getNumberMapping(ints[i]));
            } else continue;
            // 中间的零没有权重
            if (ints[i] != 0) {
                String chinesePlaceValue = ConfigurationMapping.getPlaceValue(placeValue);
                builder.append(chinesePlaceValue);
            } else if (placeValue == 5 || placeValue == 9) {
                // 不管"万位"或"亿位"的数字是否为零, 都要添加"万"或"亿"
                String chinesePlaceValue = ConfigurationMapping.getPlaceValue(placeValue);
                builder.append(chinesePlaceValue);
            }
        }
        return builder.toString();
    }

    /**
     * 获得int的每一位数字, 结果是倒序的
     */
    public int[] toArray() {
        int[] toArray = new int[0];
        int index = 0, number = this.value, size = 0;
        do {
            toArray = Arrays.copyOf(toArray, size + 1);
            toArray[size++] = number % 10;

            number /= 10;
        } while (number > 0);
        return toArray;
    }
}
