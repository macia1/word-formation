package com.zhiwei.bossdirecthireautomation;

import com.zhiwei.bossdirecthireautomation.tree.EventExcelEntity;
import org.junit.jupiter.api.Test;

class ClassUtilTest {

    @Test
    void getGetMethod() {
        EventExcelEntity eventExcelEntity = new EventExcelEntity();
        ClassUtil<EventExcelEntity> classUtil = new ClassUtil<>(EventExcelEntity.class);
        classUtil.getGetMethod().forEach((key, value) -> System.out.println(key + ": " + value));
    }
}