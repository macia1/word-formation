package com.zhiwei.nanjingspecialdaily;

import com.alibaba.excel.annotation.ExcelProperty;
import com.zhiwei.bossdirecthireautomation.NotNull;
import lombok.Data;

import java.io.Serializable;
import java.util.Date;

/**
 * 南京专项日报数据实体
 *
 * @author aszswaz
 * @date 2021/5/28 14:47:03
 * @IDE IntelliJ IDEA
 */
@Data
@SuppressWarnings("JavaDoc")
public class NanjingSpecialDailyEntity implements Serializable {
    @ExcelProperty(value = "渠道")
    @NotNull
    private String channel;

    @ExcelProperty(value = "聚合标题")
    @NotNull
    private String title;

    @ExcelProperty(value = "文本")
    @NotNull
    private String text;

    @ExcelProperty(value = "地址")
    @NotNull
    private String url;

    @ExcelProperty(value = "命中词")
    @NotNull
    private String word;

    @ExcelProperty(value = "时间")
    @NotNull
    private Date time;
}
