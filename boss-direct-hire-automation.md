# Boss直聘周报生成程序

## 主要业务逻辑

```txt
表格部分
一、表格内必要包含的抬头有：平台、来源、渠道、时间、标题、文本、地址、是否原发、情感倾向、品牌、事件、稿件，这部分由周报制作人员整理
Word部分
一、字体：微软雅黑，字号：小四
二、整体分为7大部分（以下面所列部分为优先级顺序排放）
第一部分为Boss直聘正面事件，第二部分为Boss直聘负面+敏感事件，第三部分为Boss直聘个人投诉类舆情分析（周报制作人员自行制作），第四部分为竞品正面事件，第五部分为竞品负面事件，第六部分为竞品个人投诉类舆情分析（周报制作人员自行制作），第七部分为行业负面事件，除备注周报制作人员自行制作部分，其余均需技术层面支持
三、每部分具体排列顺序
Boss直聘部分
1、各部分按调性正面和负面敏感区分（Boss直聘有负面和敏感，竞品则只有负面），先放正面模块，再放负面+敏感模块，每个模块排序均以事件传播量降序排放，需要放上事件标题，和传播量，并且加粗（传播量只列举≥5篇的稿件，若一起事件下多篇稿件传播量均小于5篇，则事件自动过滤不列举）
2、每起事件下再按稿件传播量降序排放稿件，每篇稿件需增加超链接和传播量，链接选取标准为：以原发优先，无原发情况下按渠道和平台影响力优先选取（如第一财经在微信和其他自媒体平台上均有发文，需优先选取微信链接），每篇稿件前增加【原发：XXX】标签，如无原发，则为【原发：未知】，稿件下一行有高频/重点参与渠道列举，渠道列举比例：1个原发，3个按影响力获取，3个按出现次数获取

竞品部分
1、竞品有多个品牌，以智联招聘、58同城、前程无忧/51job、猎聘为顺序排放，其余排放规则和顺序同Boss直聘部分相同
行业部分
1、行业部分无稿件，只有事件，按事件传播量降序排放，需要放上事件标题、超链接、传播量，此部分没有原发转载标签， 超链接选取权重媒体即可，传播量加粗
2、事件标题下一行有高频/重点参与渠道列举，列举权重媒体、高频参与渠道，稿件下一行有高频/重点参与渠道列举，渠道列举比例：1个原发，3个按影响力获取，3个按出现次数获取

总结
整体优先级以先Boss直聘，再竞品，最后为行业排放；竞品中以智联招聘、58同城、前程无忧/51job、猎聘为顺序排放，Boss直聘及各竞品具体模块下，以先正面，再负面事件排放，均按事件传播量降序，各事件下再按稿件传播量降序排放
```

## 使用

添加maven依赖

```xml
<dependency>
    <groupId>com.zhiwei</groupId>
    <artifactId>word-formation</artifactId>
    <version>4.0-SNAPSHOT</version>
</dependency>
```

[代码使用演示](src/test/java/com/zhiwei/bossdirecthireautomation/BossDirectHireAutomationTest.java)

```java
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
        BossDirectHireAutomation bossDirectHireAutomation = BossDirectHireAutomation.build("boss-direct-hire-automation/BOSS直聘周报测试.xlsx", "boss-direct-hire-automation/渠道生成.xlsx");
        // 生成world并输出到指定路径
        bossDirectHireAutomation.generateWord("boss-direct-hire-automation/boss-direct-hire-automation.docx");
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
```