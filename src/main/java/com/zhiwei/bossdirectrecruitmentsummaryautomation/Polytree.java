package com.zhiwei.bossdirectrecruitmentsummaryautomation;

import lombok.Getter;

import java.util.List;

/**
 * 多叉树
 *
 * @author aszswaz
 * @date 2021/5/11 18:01:54
 */
public class Polytree {
    /**
     * 根节点
     */
    @Getter
    private final PolytreeNode root = new PolytreeNode("root");

    /**
     * 添加节点
     */
    public void add(ExcelReadEntity readEntity) {
        // 媒介
        final PolytreeNode mediaTageNode = this.getNodeByName(this.root, readEntity.getMediaTage());
        // 媒体（渠道）
        final PolytreeNode mediaTypeNode = this.getNodeByName(mediaTageNode, readEntity.getChannel());
        // 来源
        final PolytreeNode sourceNode = this.getNodeByName(mediaTypeNode, readEntity.getSource());
        // 数据的实际节点
        sourceNode.add(readEntity);
    }

    /**
     * 根据节点的名称获取节点
     */
    public PolytreeNode getNodeByName(PolytreeNode polytreeNode, String nodeName) {
        List<PolytreeNode> polytreeNodes = polytreeNode.getPolytreeNodes();
        for (PolytreeNode element : polytreeNodes) {
            if (nodeName.equals(element.getNodeName())) {
                return element;
            }
        }
        polytreeNode = new PolytreeNode(nodeName);
        polytreeNodes.add(polytreeNode);
        return polytreeNode;
    }
}
