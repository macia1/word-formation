package com.zhiwei.bossdirecthireautomation;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.event.SyncReadListener;
import com.zhiwei.bossdirecthireautomation.exceptions.BossDirectHireAutomationException;
import com.zhiwei.bossdirecthireautomation.tree.EventExcelEntity;
import lombok.extern.log4j.Log4j2;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.List;

@Log4j2
class BossDirectHireAutomationTest {

    /**
     * 第一种方式
     *
     * @throws BossDirectHireAutomationException 数据文件存在问题
     */
    @Test
    public void demo01() throws BossDirectHireAutomationException, IOException {
        // 直接传入数据源文件和渠道文件，
        BossDirectHireAutomation bossDirectHireAutomation = BossDirectHireAutomation.build("source/work/boss-补充.xlsx", "source/work/渠道-补充.xlsx");
        // 生成world并输出到指定路径
        bossDirectHireAutomation.generateWord("source/work/boss-direct-hire-automation.docx");
    }

    /**
     * 第二种
     *
     * @throws BossDirectHireAutomationException 数据文件存在问题
     */
    @Test
    public void demo02() throws BossDirectHireAutomationException, IOException {
        // 直接传入对象，但是需要渠道以及影响力不能为空，否则不能通过数据检查
        List<EventExcelEntity> entities = EasyExcel.read("boss-direct-hire-automation/BOSS直聘周报测试.xlsx", EventExcelEntity.class, new SyncReadListener()).sheet(0).doReadSync();
        List<EventExcelEntity.Influence> influences = EasyExcel.read("boss-direct-hire-automation/渠道生成.xlsx", EventExcelEntity.Influence.class, new SyncReadListener()).sheet(0).doReadSync();
        // 匹配影响力
        for (EventExcelEntity entity : entities) {
            for (EventExcelEntity.Influence influence : influences) {
                if (entity.getChannel().equals(influence.getChannel()) && entity.getPlatform().equals(influence.getPlatform())) {
                    entity.setInfluence(influence.getInfluence());
                }
            }
        }
        BossDirectHireAutomation automation = BossDirectHireAutomation.build(entities);
        //  生成world并输出到指定路径
        automation.generateWord("boss-direct-hire-automation/boss-direct-hire-automation.docx");
    }

    /**
     * 第三种
     * 如果在数据的源文件里面就已经包含渠道影响力和渠道，可以使用该方式
     *
     * @throws BossDirectHireAutomationException 数据文件存在问题
     */
    @Test
    public void demo03() throws BossDirectHireAutomationException, IOException {
        BossDirectHireAutomation automation = BossDirectHireAutomation.build("boss-direct-hire-automation/BOSS直聘周报测试.xlsx");
        automation.generateWord("boss-direct-hire-automation/boss-direct-hire-automation.docx");
    }
}