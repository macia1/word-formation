package com.zhiwei.exception;

/**
 * word模板版本不支持
 *
 * @author 朽木不可雕也
 * @date 2020/9/25 10:27
 * @week 星期五
 */
public class WordTemplateVersionException extends RuntimeException {
    public WordTemplateVersionException(String message) {
        super(message);
    }
}
