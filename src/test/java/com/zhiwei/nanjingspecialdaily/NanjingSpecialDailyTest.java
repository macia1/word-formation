package com.zhiwei.nanjingspecialdaily;

import org.junit.jupiter.api.Test;

import java.io.IOException;

/**
 * @author aszswaz
 * @date 2021/5/28 15:44:19
 * @IDE IntelliJ IDEA
 */
@SuppressWarnings("JavaDoc")
class NanjingSpecialDailyTest {

    /**
     * 南京专项报告
     */
    @Test
    void start() throws NanjingSpecialDailyException, IOException {
        String writePath = "target/" + NanjingSpecialDaily.class.getSimpleName() + ".docx";
        NanjingSpecialDaily.start("source/试 爱美客 标题聚类（6月1日）.xlsx", writePath);
    }
}