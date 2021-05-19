package com.zhiwei.bossdirecthireautomation.tree;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.event.SyncReadListener;
import lombok.extern.log4j.Log4j2;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.List;

/**
 * 多叉树的实现工具类
 *
 * @author fuck-aszswaz
 * @date 2021-02-20 14:19:17
 */
@Log4j2
class PolyTreeUtilTest {
    private PolyTreeUtil treeUtil;

    PolyTreeUtilTest() {
    }

    @BeforeEach
    void setUp() {
        // 读取文件
        List<EventExcelEntity> eventExcelEntities = EasyExcel.read("source/Boss直聘及竞品数据.xlsx", EventExcelEntity.class, new SyncReadListener()).doReadAllSync();
        log.info("获得{}条数据", eventExcelEntities.size());
        PolyTreeUtil polyTreeUtil = new PolyTreeUtil();
        // 添加到树中
        eventExcelEntities.forEach(polyTreeUtil::add);
        this.treeUtil = polyTreeUtil;
    }

    /**
     * 排序测试
     */
    @Test
    void sort() {
        this.treeUtil.sort();
    }
}