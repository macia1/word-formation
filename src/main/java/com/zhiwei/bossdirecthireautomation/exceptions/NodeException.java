package com.zhiwei.bossdirecthireautomation.exceptions;

/**
 * 节点异常
 *
 * @author fuck-aszswaz
 * @date 2021-02-20 11:05:26
 */
@SuppressWarnings("unused")
public class NodeException extends RuntimeException {
    public NodeException(String message) {
        super(message);
    }

    public NodeException(String message, Throwable cause) {
        super(message, cause);
    }
}
