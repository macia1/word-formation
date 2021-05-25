package com.zhiwei.bossdirectrecruitmentsummaryautomation;

import lombok.Getter;

import java.util.Comparator;
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
        // 排序
        List<PolytreeNode> sourceNodes = mediaTypeNode.getPolytreeNodes();
        sourceNodes.sort(Comparator.comparing(sourceNodeCom -> {
            String source = sourceNodeCom.getNodeName();
            return Weights.getWeights(source);
        }));
        // 数据的实际节点
        sourceNode.add(readEntity);
    }

    /**
     * 根据节点的名称获取节点
     */
    public PolytreeNode getNodeByName(PolytreeNode polytreeNode, String nodeName) {
        List<PolytreeNode> polytreeNodes = polytreeNode.getPolytreeNodes();
        for (PolytreeNode element : polytreeNodes) {
            if (nodeName.equalsIgnoreCase(element.getNodeName())) {
                return element;
            }
        }
        polytreeNode = new PolytreeNode(nodeName);
        polytreeNodes.add(polytreeNode);
        return polytreeNode;
    }
}
