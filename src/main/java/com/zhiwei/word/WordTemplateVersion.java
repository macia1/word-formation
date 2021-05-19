package com.zhiwei.word;

/**
 * word版本样式
 *
 * @author 朽木不可雕也
 * @date 2020/9/15 15:25
 * @week 星期二
 */
public enum WordTemplateVersion {
    /**
     * 1、SCS版本格式（全部中文+韩文版）：
     * a. 四个板块分别为：三星关联、SCS关联及陕西新闻、事业及竞争对手、中国新闻，字体微软雅黑小四；
     * b. 中文微软雅黑、韩文BatangChe、英文及日期Arial（除正文以外，正文里英文和数字格式与当前文本格式保持一致），字号均为小五，单倍行距，两端对齐；
     * c. 返回首页与对应标题书签可交叉引用；
     */
    SAS("SAS-template.docx"),
    /**
     * 2、SAS版本格式（8条中文+韩文翻译版）：
     * a. 四个板块韩文版：삼성、섬서성(SCS)、사업환경 & 경쟁사、정치 & 경제，字体BatangChe，字号四号；
     * b. 中文微软雅黑11号、韩文BatangChe小四，英文及日期Arial11号（除正文以外，正文里英文和数字格式与当前文本格式保持一致），单倍行距，两端对齐；
     * c. 表格中日期右对齐，全文部分标题加粗；
     * d. 返回首页与对应标题书签可交叉引用；
     */
    SCS("SCS-template.docx"),
    /**
     * 3、SCS转图片版本（韩文版）
     * a. 四个板块韩文版跟SAS版本保持一致，字体BatangChe，字号小二；
     * b. 韩文资讯标题字体BatangChe，字号小四，中文标题字体仿宋，字号11，日期及域名字体仿宋，字号10，韩文正文字体BatangChe，字号小五，单倍行距，两端对齐；
     */
    SCS_IMAGE("SCS-image-template.docx");

    private final String templateFile;

    WordTemplateVersion(String templatePath) {
        this.templateFile = templatePath;
    }

    public String getTemplateFile() {
        return templateFile;
    }
}
