package com.zhiwei.bossdirecthireautomation.wordutil;

import lombok.extern.log4j.Log4j2;
import org.junit.jupiter.api.Test;

import java.util.Arrays;

/**
 * 多叉树的实现工具类
 *
 * @author fuck-aszswaz
 * @date 2021-02-20 18:13:44
 */
@Log4j2
class NumberUtilTest {

    @Test
    void toArray() {
        NumberUtil numberUtil = new NumberUtil(12030);
        log.info(Arrays.toString(numberUtil.toArray()));
    }

    @Test
    void toChineseNumber() {
        NumberUtil numberUtil = new NumberUtil(0);
        log.info(numberUtil.toChineseNumber());
        for (int i = 1; i < 1000; i++) {
            numberUtil.setValue(i);
            log.info(numberUtil.toChineseNumber());
        }
    }
}