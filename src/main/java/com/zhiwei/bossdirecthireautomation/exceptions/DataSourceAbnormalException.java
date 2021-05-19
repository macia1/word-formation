package com.zhiwei.bossdirecthireautomation.exceptions;

/**
 * 数据源异常
 */
@SuppressWarnings("unused")
public class DataSourceAbnormalException extends BossDirectHireAutomationException {
    public DataSourceAbnormalException(String message) {
        super(message);
    }

    public DataSourceAbnormalException(String message, Throwable cause) {
        super(message, cause);
    }
}
