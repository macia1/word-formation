package com.zhiwei.bossdirecthireautomation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 标注excel中不得为空的字段
 *
 * @author fuck-aszswaz
 * @date 2021-02-20 15:15:48
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface NotNull {
}
