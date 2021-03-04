package com.zhiwei.bossdirecthireautomation.tree;

import com.zhiwei.bossdirecthireautomation.exceptions.NodeException;
import lombok.*;
import org.jetbrains.annotations.NotNull;

import java.util.ArrayList;
import java.util.List;

@ToString(callSuper = true)
@EqualsAndHashCode
@Data
public class PolyTreeNode {
    /**
     * 节点名称
     */
    @Setter
    private String nodeName;
    /**
     * 数据节点的数量
     */
    private int dataNodeSize = 0;
    private final List<PolyTreeNode> nodes = new ArrayList<>();

    /**
     * 判断当前节点是否已包含该子节点节点名称
     *
     * @param name 子节点名称
     */
    public boolean containNodeByName(@NotNull String name) {
        return this.nodes.stream().anyMatch(polyTreeNode -> polyTreeNode.getNodeName().equalsIgnoreCase(name));
    }

    /**
     * 添加节点
     *
     * @param polyTreeNode 新节点
     */
    public void add(@NotNull PolyTreeNode polyTreeNode) {
        if (this.containNodeByName(polyTreeNode.getNodeName()) && !(polyTreeNode instanceof EventExcelEntity)) {
            throw new NodeException("节点重复添加");
        }
        this.nodes.add(polyTreeNode);
    }

    /**
     * 获取节点的子节点， 如果节点不存在则创建
     *
     * @param nodeName 节点名称
     */
    public PolyTreeNode getChildNodeOrNew(@NotNull String nodeName) {
        if (!this.nodes.isEmpty() && this.containNodeByName(nodeName)) {
            return this.getChildNode(nodeName);
        } else {
            PolyTreeNode polyTreeNode = new PolyTreeNode();
            polyTreeNode.setNodeName(nodeName);
            // 添加一级节点到根节点
            this.add(polyTreeNode);
            return polyTreeNode;
        }
    }

    /**
     * 获取节点的子节点
     *
     * @param nodeName 节点名称
     */
    public PolyTreeNode getChildNode(String nodeName) {
        return this.nodes.stream().filter(polyTreeNode -> polyTreeNode.getNodeName().equalsIgnoreCase(nodeName)).findFirst().orElse(null);
    }

    /**
     * 是否为空
     */
    public boolean isEmpty() {
        return this.nodes.isEmpty();
    }
}
