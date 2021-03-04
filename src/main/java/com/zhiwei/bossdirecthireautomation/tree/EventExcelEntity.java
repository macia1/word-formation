package com.zhiwei.bossdirecthireautomation.tree;

import com.alibaba.excel.annotation.ExcelProperty;
import com.zhiwei.bossdirecthireautomation.NotNull;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.ToString;

/**
 * 事件实体类
 *
 * @author fuck-aszswaz
 */
@EqualsAndHashCode(callSuper = true)
@ToString(callSuper = true)
@Data
public class EventExcelEntity extends PolyTreeNode {
    @ExcelProperty(value = "平台")
    @NotNull
    private String platform;
    @ExcelProperty(value = "来源")
    private String source;
    @ExcelProperty(value = "渠道")
    @NotNull
    private String channel;
    @ExcelProperty(value = "时间")
    @NotNull
    private String time;
    @ExcelProperty(value = "标题")
    private String title;
    @ExcelProperty(value = "事件")
    @NotNull
    private String event;
    @ExcelProperty(value = "稿件")
    @NotNull
    private String manuscript;
    @ExcelProperty(value = "文本")
    private String text;
    @ExcelProperty(value = "链接")
    @NotNull
    private String url;
    @ExcelProperty(value = "原发/转载")
    @NotNull
    private String whether;
    @ExcelProperty(value = "情感倾向")
    @NotNull
    private String emotion;
    @ExcelProperty(value = "品牌")
    @NotNull
    private String brand;
    @NotNull
    @ExcelProperty(value = "影响力")
    private double influence;
    /**
     * 数据标签
     */
    @ExcelProperty(value = "品牌区分")
    @NotNull
    private String label;

    /**
     * 平台和渠道的影响力
     */
    @Data
    public static class Influence {
        /**
         * 渠道
         */
        @ExcelProperty(value = "渠道")
        private String channel;
        /**
         * 平台
         */
        @ExcelProperty(value = "平台")
        private String platform;
        /**
         * 影响力
         */
        @ExcelProperty(value = "影响力")
        private double influence;
    }
}
