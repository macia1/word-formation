package com.zhiwei.word;

import com.alibaba.excel.EasyExcel;
import com.zhiwei.config.ExcelTitles;
import com.zhiwei.exception.WordTemplateVersionException;
import com.zhiwei.util.ExcelMonitor;
import com.zhiwei.word.wordversion.SASWord;
import com.zhiwei.word.wordversion.SCSWord;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.jetbrains.annotations.NotNull;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * word基本属性
 *
 * @author 朽木不可雕也
 * @date 2020/9/17 15:52
 * @week 星期四
 */
@SuppressWarnings({"SameParameterValue", "DuplicatedCode"})
public abstract class BaseWord extends XWPFDocument {
    protected ExcelMonitor monitor;
    protected OutputStream wordOut;

    protected BaseWord(String path) throws IOException {
        super(Objects.requireNonNull(BaseWord.class.getClassLoader().getResourceAsStream(path)));
    }

    /**
     * 初始化word对象
     *
     * @param templateVersion  word模板的版本
     * @param excelInputStream excel文件流
     * @param wordOut          word文件输出流
     * @throws IOException                   io异常
     * @throws ExcelTitleIncompleteException excel的标题缺失异常
     */
    public static void start(
            @NotNull WordTemplateVersion templateVersion,
            @NotNull InputStream excelInputStream,
            @NotNull OutputStream wordOut
    ) throws IOException, ExcelTitleIncompleteException {
        BaseWord baseWord;
        switch (templateVersion) {
            case SAS:
                baseWord = SASWord.init();
                break;
            case SCS:
                baseWord = SCSWord.init();
                break;
            /*case SCS_IMAGE:
                baseWord = SCSImageWord.init();
                break;*/
            default:
                throw new WordTemplateVersionException("不支持的版本：" + templateVersion.toString());
        }
        //读取excel文件
        baseWord.readExcel(excelInputStream, templateVersion);
        baseWord.wordOut = wordOut;
        //调用子类start开始解析模板
        baseWord.start();
        //输出word
        baseWord.write(wordOut);
    }

    /**
     * 启动word读写操作
     */
    protected abstract void start();

    /**
     * 读取excel文件
     *
     * @throws ExcelTitleIncompleteException excel文件的标题缺失异常
     */
    protected final void readExcel(@NotNull InputStream excelInputStream, WordTemplateVersion templateVersion) throws ExcelTitleIncompleteException {
        //开始读取数据文件
        ExcelMonitor monitor = new ExcelMonitor();
        EasyExcel.read(excelInputStream, monitor).sheet(0).doRead();
        this.monitor = monitor;
        //excel文件标题检查
        if (templateVersion == WordTemplateVersion.SCS) {
            List<ExcelTitles> error = new ArrayList<>();
            for (ExcelTitles titles : ExcelTitles.values()) {
                if (!monitor.getHeaders().containsKey(titles.name())) {
                    //缺失标题
                    error.add(titles);
                }
            }
            if (!error.isEmpty()) {
                StringBuilder errorMessageBuilder = new StringBuilder("excel文件标题缺失：");
                Iterator<ExcelTitles> iterator = error.iterator();
                while (iterator.hasNext()) {
                    errorMessageBuilder.append(iterator.next().name());
                    if (iterator.hasNext()) {
                        errorMessageBuilder.append(",");
                    }
                }
                throw new ExcelTitleIncompleteException(errorMessageBuilder.toString());
            }
        }
    }

    /**
     * 读取数据并进行分类
     *
     * @return 数据分类后
     */
    @NotNull
    protected final Map<String, List<List<Object>>> classification(String classificationColumn) {
        //进行数据筛选
        Map<String, List<List<Object>>> classificationMap = new LinkedHashMap<>(this.monitor.getDataList().size() / 2);
        Integer column = this.monitor.getHeaders().get(classificationColumn);//获得筛选依据的列
        for (List<Object> data : this.monitor.getDataList()) {
            Object object = data.get(column);
            if (object instanceof String) {
                String classification = (String) object;
                if (classificationMap.containsKey(classification)) {
                    List<List<Object>> dataList = classificationMap.get(classification);
                    dataList.add(data);
                } else {
                    List<List<Object>> lists = new ArrayList<>();
                    lists.add(data);
                    classificationMap.put(classification, lists);
                }
            }
        }
        return classificationMap;
    }

    /**
     * 创建书签
     *
     * @param paragraph    段落
     * @param bookmarkName 书签名称
     */
    protected void createBookmark(@NotNull XWPFParagraph paragraph, String bookmarkName) {
        CTBookmark ctBookmark = paragraph.getCTP().addNewBookmarkStart();
        ctBookmark.setName(bookmarkName);
        ctBookmark.setId(BigInteger.valueOf(0));
        paragraph.getCTP().addNewBookmarkEnd().setId(BigInteger.valueOf(0));
    }

    /**
     * 根据指定的单元格数量，创建单元格。
     *
     * @param tableRow 表格行
     * @param number   创建单元格的数量
     */
    protected final void createCells(@NotNull XWPFTableRow tableRow, int number) {
        XWPFTableCell cell = tableRow.getCell(0);
        //去除边框
        CTTcBorders tcBorders = CTTcBorders.Factory.newInstance();
        tcBorders.addNewLeft().setVal(STBorder.NIL);
        tcBorders.addNewRight().setVal(STBorder.NIL);
        tcBorders.addNewTop().setVal(STBorder.NIL);
        tcBorders.addNewBottom().setVal(STBorder.NIL);
        cell.getCTTc().addNewTcPr().setTcBorders(tcBorders);
        for (int i = 0; i < number; i++) {
            cell = tableRow.createCell();
            cell.getCTTc().addNewTcPr().setTcBorders(tcBorders);
        }
    }

    /**
     * excel文件标题缺失
     */
    public static class ExcelTitleIncompleteException extends Exception {
        public ExcelTitleIncompleteException(String message) {
            super(message);
        }
    }
}
