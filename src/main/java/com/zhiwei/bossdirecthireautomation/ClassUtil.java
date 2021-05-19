package com.zhiwei.bossdirecthireautomation;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.Map;

/**
 * 类初始化工具
 */
public class ClassUtil<T> {
    private final Class<T> thisClass;

    public ClassUtil(Class<T> thisClass) {
        this.thisClass = thisClass;
    }

    /**
     * 获得属性对应的get方法
     */
    public Map<Field, Method> getGetMethod() {
        Map<Field, Method> fieldMethodMap = new HashMap<>();
        Method[] methods = thisClass.getMethods();
        Field[] fields = this.thisClass.getDeclaredFields();

        for (Method method : methods) {
            String methodName = method.getName();
            boolean isGetMethod = methodName.contains("get") && !methodName.equals("getClass");// 是否get方法
            if (!isGetMethod) continue;
            methodName = methodName.substring(3);// 去除前三位字母
            for (Field field : fields) {
                String fieldName = field.getName();
                if (methodName.equalsIgnoreCase(fieldName)) {
                    fieldMethodMap.put(field, method);
                    break;
                }
            }
        }
        return fieldMethodMap;
    }
}
