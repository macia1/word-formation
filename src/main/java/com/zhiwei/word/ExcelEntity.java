package com.zhiwei.word;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.ToString;

/**
 * excel实体
 *
 * @author 朽木不可雕也
 * @date 2020/9/27 13:15
 * @week 星期日
 */
@SuppressWarnings("unused")
@EqualsAndHashCode
@ToString
@Data
public class ExcelEntity {
    /**
     * 序号
     */
    private String serialNumber;
    /**
     * 标题
     */
    private String title;
    /**
     * 全文
     */
    private String fullText;
    /**
     * 韩文标题
     */
    private String koreanTitle;
    /**
     * 韩文摘要
     */
    private String koreanAbstract;
    /**
     * 渠道
     */
    private String channel;
    /**
     * 渠道域名
     */
    private String channelDomainName;
    /**
     * 时间
     */
    private String time;
    /**
     * 中文品牌分类
     */
    private String chineseBrandClassification;
    /**
     * 韩文品牌分类
     */
    private String koreanBrandClassification;
    /**
     * 书签，由jar包内部赋值
     */
    private BaseWord.Bookmark bookmark;

    void setBookmark(BaseWord.Bookmark bookmark) {
        this.bookmark = bookmark;
    }
}