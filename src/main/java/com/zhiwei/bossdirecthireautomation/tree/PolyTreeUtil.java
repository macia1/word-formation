package com.zhiwei.bossdirecthireautomation.tree;

import com.zhiwei.bossdirecthireautomation.wordutil.BaseConfig;
import lombok.Getter;
import lombok.ToString;
import org.jetbrains.annotations.NotNull;

import java.util.Collections;
import java.util.Comparator;
import java.util.List;

/**
 * 多叉树的实现工具类
 *
 * @author fuck-aszswaz
 */
@ToString
public class PolyTreeUtil {
    /**
     * 根节点
     */
    @Getter
    private final PolyTreeNode rootNode = new PolyTreeNode();

    {
        this.rootNode.setNodeName("root");
    }

    /**
     * 添加数据
     *
     * @param data 一条数据
     */
    public void add(@NotNull EventExcelEntity data) {
        // 处理例外品牌
        this.makeException(data);
        // 标签分类
        final PolyTreeNode labelNode = this.rootNode.getChildNodeOrNew(data.getLabel());
        // 从根节点获得品牌节点
        final PolyTreeNode brandNode = labelNode.getChildNodeOrNew(data.getBrand());
        // 从品牌节点获取情感节点
        final PolyTreeNode emotionNode = brandNode.getChildNodeOrNew(data.getEmotion());
        // 从情感节点获取事件节点
        final PolyTreeNode eventNode = emotionNode.getChildNodeOrNew(data.getEvent());
        // 获取稿件节点
        final PolyTreeNode manuscriptNode = eventNode.getChildNodeOrNew(data.getManuscript());
        // 将数据作为数据节点放入多叉树
        data.setNodeName(data.getBrand() + "-" + data.getEmotion() + "-" + data.getEvent() + "-" + data.getManuscript());
        manuscriptNode.add(data);
        // 分别自增
        brandNode.setDataNodeSize(brandNode.getDataNodeSize() + 1);
        emotionNode.setDataNodeSize(emotionNode.getDataNodeSize() + 1);
        eventNode.setDataNodeSize(eventNode.getDataNodeSize() + 1);
        manuscriptNode.setDataNodeSize(manuscriptNode.getDataNodeSize() + 1);
        data.setDataNodeSize(1);
    }

    /**
     * 处理例品牌
     *
     * @param eventExcelEntity 一条数据
     */
    public void makeException(@NotNull EventExcelEntity eventExcelEntity) {
        // 处理Boss直聘例外
        final String exception = "Boss直聘";
        if (eventExcelEntity.getBrand().equalsIgnoreCase(exception)) {
            if ("负面".equalsIgnoreCase(eventExcelEntity.getEmotion()) || "敏感".equalsIgnoreCase(eventExcelEntity.getEmotion())) {
                eventExcelEntity.setEmotion(BaseConfig.BOSS_EXCEPTION_EMOTION);
            }
        }
    }

    /**
     * 排序
     */
    private void sort(PolyTreeNode labelNode) {
        final List<PolyTreeNode> brandNodes = labelNode.getNodes();
        {
            // 对品牌进行排序
            for (int i = 0; i < BaseConfig.brandShort.size(); i++) {
                if (i >= brandNodes.size()) break;// 防止下标越界
                for (int j = 0; j < brandNodes.size(); j++) {
                    PolyTreeNode node = brandNodes.get(j);
                    if (j != 0 && node.getNodeName().equalsIgnoreCase(BaseConfig.brandShort.get(i))) {
                        brandNodes.add(i, brandNodes.remove(j));
                        break;
                    }
                }
            }
        }

        // 排序
        {
            for (PolyTreeNode brand : brandNodes) {
                List<PolyTreeNode> emotions = brand.getNodes();
                for (PolyTreeNode emotion : emotions) {
                    List<PolyTreeNode> events = emotion.getNodes();// 事件
                    events.sort(Collections.reverseOrder(Comparator.comparingInt(PolyTreeNode::getDataNodeSize)));// 事件按照传播量降序

                    for (PolyTreeNode event : events) {
                        List<PolyTreeNode> manuscripts = event.getNodes();// 稿件
                        // 稿件按照传播量降序
                        manuscripts.sort(Collections.reverseOrder(Comparator.comparingInt(PolyTreeNode::getDataNodeSize)));

                        for (PolyTreeNode manuscript : manuscripts) {
                            // 排序每一条数据， 将“原发”的数据置顶
                            List<PolyTreeNode> data = manuscript.getNodes();
                            for (int i = 0; i < data.size(); i++) {
                                EventExcelEntity value = (EventExcelEntity) data.get(i);
                                if ("原发".equals(value.getWhether())) {
                                    data.add(0, data.remove(i));
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * 排序
     */
    public void sort() {
        this.rootNode.getNodes().forEach(this::sort);
    }
}
