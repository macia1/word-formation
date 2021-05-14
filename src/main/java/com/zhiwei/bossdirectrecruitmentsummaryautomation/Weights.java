package com.zhiwei.bossdirectrecruitmentsummaryautomation;

/**
 * 获取平台的权重
 *
 * @author aszswaz
 * @date 2021/5/14 15:00:52
 */
public class Weights {
    public static int getWeights(String source) {
        if ("微信".equals(source) || "微信公众号".equals(source)) {
            return 1;
        } else if ("微博".equals(source) || "新浪微博".equals(source)) {
            return 2;
        } else if ("网媒".equals(source)) {
            return 4;
        } else {
            // 其他
            return 3;
        }
    }
}
