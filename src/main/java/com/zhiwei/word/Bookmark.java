package com.zhiwei.word;

import lombok.Data;

/**
 * @author 朽木不可雕也
 * @date 2020/9/18 15:01
 * @week 星期五
 */
@Data
public class Bookmark {
    /**
     * 目标书签，文本为需要跳转的目标文本
     */
    private String target;
    /**
     * 返回书签，用户返回目录索引处
     */
    private String returnMark;

    public Bookmark(String target, String returnMark) {
        this.target = target;
        this.returnMark = returnMark;
    }
}
