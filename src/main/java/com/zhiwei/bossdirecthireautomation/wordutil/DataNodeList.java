package com.zhiwei.bossdirecthireautomation.wordutil;

import com.zhiwei.bossdirecthireautomation.tree.EventExcelEntity;

import java.util.ArrayList;

/**
 * 数据集合
 *
 * @author fuck-aszswaz
 * @date 2021-02-25 14:57:56
 */
public class DataNodeList extends ArrayList<EventExcelEntity> {

    /**
     * 生成
     */
    @Override
    public boolean add(EventExcelEntity eventExcelEntity) {
        // 取出完全相同的对象
        if (super.contains(eventExcelEntity)) return false;

        for (int i = 0; i < super.size(); i++) {
            if (super.get(i).getChannel().equals(eventExcelEntity.getChannel())) {
                EventExcelEntity result = this.comparison(eventExcelEntity, super.get(i));
                // 已包含相同的渠道和平台
                super.set(i, result);
                return true;
            }
        }

        return super.add(eventExcelEntity);
    }

    /**
     * 比对
     */
    private EventExcelEntity comparison(EventExcelEntity o1, EventExcelEntity o2) {
        // 判断是否原发
        if ("原发".equals(o1.getWhether()) && !"原发".equals(o2.getWhether())) return o1;
        // 比较影响力
        return o1.getInfluence() > o2.getInfluence() ? o1 : o2;
    }
}
