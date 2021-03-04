package com.zhiwei.bossdirecthireautomation.wordutil;

import lombok.Getter;
import lombok.ToString;

/**
 * 打印机的字号与磅值的换算
 *
 * @author 朽木不可雕也
 * @date 2020/9/18 11:52
 * @week 星期五
 */
@SuppressWarnings({"NonAsciiCharacters", "JavaDoc"})
@Getter
@ToString
public enum ContentFont {
    初号(42),
    小初(36),
    一号(26),
    小一(24),
    二号(22),
    小二(18),
    三号(16),
    小三(15),
    四号(14),
    小四(12),
    五号(10),
    小五(9),
    六号(7),
    小六(6),
    /**
     * 实际上是5.5，但是在程序中的磅值都是int，无法应用小数
     */
    八号(5),
    七号(5);
    /**
     * 磅值
     */
    private final int poundValue;

    ContentFont(int poundValue) {
        this.poundValue = poundValue;
    }
}
