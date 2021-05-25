package com.zhiwei.word.huawei;

/**
 * @author aszswaz
 * @date 2021/4/2 11:19:50
 */
public enum HuaWeiStyleEnum {
    /**
     * 格式一
     */
    STYLE_ONE("格式一"),

    /**
     * 格式er
     */
    STYLE_TWO("格式二");

    HuaWeiStyleEnum(String style){
        this.style = style;
    }

    public String style;

    public static String match(String styleName){
        HuaWeiStyleEnum[] values = HuaWeiStyleEnum.values();
        for (HuaWeiStyleEnum value : values) {
            if (value.style.equals(styleName)){
                return value.getStyle();
            }
        }
        return null;
    }

    public String getStyle(){
        return style;
    }
}
