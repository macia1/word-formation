package com.zhiwei.nanjingspecialdaily;

import org.junit.jupiter.api.Test;

import java.awt.*;
import java.io.File;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

/**
 * @author aszswaz
 * @date 2021/5/28 15:44:19
 * @IDE IntelliJ IDEA
 */
@SuppressWarnings("JavaDoc")
class NanjingSpecialDailyTest {

    @Test
    void start() throws NanjingSpecialDailyException, IOException {
        String writePath = "target/" + NanjingSpecialDaily.class.getSimpleName() + ".docx";
        NanjingSpecialDaily.start("source/试 爱美客 标题聚类（6月1日）.xlsx", writePath);
    }
}