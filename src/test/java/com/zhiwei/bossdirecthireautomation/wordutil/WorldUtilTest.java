package com.zhiwei.bossdirecthireautomation.wordutil;

import com.zhiwei.bossdirecthireautomation.BossDirectHireAutomation;
import com.zhiwei.bossdirecthireautomation.exceptions.BossDirectHireAutomationException;
import com.zhiwei.bossdirecthireautomation.tree.PolyTreeNode;
import com.zhiwei.bossdirecthireautomation.tree.PolyTreeUtil;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

/**
 * @author aszswaz
 * @date 2021/3/5 12:07:28
 */
class WorldUtilTest {
    private WorldUtil worldUtil;
    private PolyTreeUtil polyTreeUtil;

    @BeforeEach
    void setUp() throws BossDirectHireAutomationException {
        BossDirectHireAutomation automation = BossDirectHireAutomation.build("source/work/BOSS直聘周报.xlsx", "source/work/渠道匹配.xlsx");
        this.polyTreeUtil = new PolyTreeUtil();
        automation.getBrandData().forEach(this.polyTreeUtil::add);
        this.worldUtil = new WorldUtil(polyTreeUtil);
    }

    @Test
    void chooseChannel() {
        PolyTreeNode labelNode = this.polyTreeUtil.getRootNode().getChildNode("Boss直聘");
        PolyTreeNode brandNode = labelNode.getChildNode("Boss直聘");
        PolyTreeNode emotionNode = brandNode.getChildNode("正面");
        PolyTreeNode eventNode = emotionNode.getChildNode("BOSS直聘《2021年春节复工首周就业趋势观察》");
        PolyTreeNode manuscript = eventNode.getChildNode("节后市场平均薪资超8K，用人需求同比翻倍丨2021年春节复工首周就业趋势观察");
        this.worldUtil.chooseChannel(manuscript, null).forEach(System.out::println);
    }
}