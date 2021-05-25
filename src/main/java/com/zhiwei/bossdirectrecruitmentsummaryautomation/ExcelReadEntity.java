package com.zhiwei.bossdirectrecruitmentsummaryautomation;

import com.alibaba.excel.annotation.ExcelProperty;
import com.zhiwei.bossdirecthireautomation.NotNull;
import lombok.*;

import java.util.Date;

/**
 * @author aszswaz
 * @date 2021/5/11 14:24:57
 */
@Getter
@Setter
@EqualsAndHashCode(callSuper = true)
@ToString(callSuper = true)
public class ExcelReadEntity extends PolytreeNode {
    @ExcelProperty(value = "平台")
    private String platform;

    @ExcelProperty(value = "来源")
    @NotNull
    private String source;

    @ExcelProperty(value = "渠道")
    @NotNull
    private String channel;

    @ExcelProperty(value = "时间")
    private Date time;

    @ExcelProperty(value = "标题")
    @NotNull
    private String title;

    @ExcelProperty(value = "文本")
    private String text;

    @ExcelProperty(value = "地址")
    @NotNull
    private String url;

    @ExcelProperty(value = "是否源发")
    private String original;

    @ExcelProperty(value = "情感倾向")
    private String emotion;

    @ExcelProperty(value = "媒体区分")
    private String mediaType;

    @ExcelProperty(value = "媒介")
    private String mediaTage;


    public ExcelReadEntity() {
        super(null);
    }
}
