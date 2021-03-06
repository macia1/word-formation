package com.zhiwei.bossdirectrecruitmentsummaryautomation;

import com.alibaba.excel.annotation.ExcelIgnore;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;

/**
 * 多茶树节点
 *
 * @author aszswaz
 * @date 2021/5/11 17:55:21
 */
@Data
public class PolytreeNode {
    /**
     * 节点名称
     */
    @ExcelIgnore
    private final String nodeName;

    /**
     * 子节点
     */
    @ExcelIgnore
    private final List<PolytreeNode> polytreeNodes = new ArrayList<>();

    public PolytreeNode(String nodeName) {
        this.nodeName = nodeName;
    }

    public void add(PolytreeNode polytreeNode) {
        this.polytreeNodes.add(polytreeNode);
    }
}
