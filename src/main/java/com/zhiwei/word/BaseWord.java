package com.zhiwei.word;

import com.zhiwei.exception.WordTemplateVersionException;
import com.zhiwei.word.wordversion.SASWord;
import com.zhiwei.word.wordversion.SCSImageWord;
import com.zhiwei.word.wordversion.SCSWord;
import lombok.Data;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.jetbrains.annotations.Contract;
import org.jetbrains.annotations.NotNull;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
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
@SuppressWarnings({"SameParameterValue", "DuplicatedCode", "unused"})
public abstract class BaseWord extends XWPFDocument {
    protected List<ExcelEntity> dataList = new ArrayList<>();
    protected OutputStream wordOut;

    protected BaseWord(String path) throws IOException {
        super(Objects.requireNonNull(BaseWord.class.getClassLoader().getResourceAsStream(path)));
    }

    /**
     * 初始化word对象
     *
     * @param templateVersion word模板的版本
     * @param dataList        excel数据
     * @param wordOut         word文件输出流
     * @throws IOException                  io异常
     * @throws WordTemplateVersionException word模板的版本不支持
     */
    public static void start(
            @NotNull WordTemplateVersion templateVersion,
            @NotNull List<ExcelEntity> dataList,
            @NotNull OutputStream wordOut
    ) throws IOException {
        BaseWord baseWord;
        switch (templateVersion) {
            case SAS:
                baseWord = SASWord.init();
                baseWord.conversion(dataList);
                break;
            case SCS:
                baseWord = SCSWord.init();
                baseWord.conversion(dataList);
                break;
            case SCS_IMAGE:
                baseWord = SCSImageWord.init();
                baseWord.conversion(dataList);
                break;
            default:
                throw new WordTemplateVersionException("不支持的版本：" + templateVersion.toString());
        }
        baseWord.wordOut = wordOut;
        //调用子类start开始解析模板
        baseWord.start();
        //输出word
        baseWord.write(wordOut);
        baseWord.close();
    }

    /**
     * 数据转换
     *
     * @param dataList 数据集
     */
    private void conversion(@NotNull List<ExcelEntity> dataList) {
        //转换数据存储方式，并生成书签
        for (int i = 0; i < dataList.size(); i++) {
            ExcelEntity dataColumn = dataList.get(i);
            dataColumn.setBookmark(new Bookmark("t" + i, "r" + i));
            this.dataList.add(dataColumn);
        }
    }

    /**
     * 将数据转换成word文档
     *
     * @param wordTemplateVersion word模板的版本
     * @param dataList            数据集
     * @param location            输出地址
     */
    public static void start(
            @NotNull WordTemplateVersion wordTemplateVersion,
            @NotNull List<ExcelEntity> dataList,
            @NotNull String location
    ) throws IOException {
        start(wordTemplateVersion, dataList, new FileOutputStream(location));
    }

    /**
     * 将数据转换成word文档
     *
     * @param wordTemplateVersion word模板的版本
     * @param dataList            数据集
     * @param outFile             输出文件
     */
    public static void start(
            @NotNull WordTemplateVersion wordTemplateVersion,
            @NotNull List<ExcelEntity> dataList,
            @NotNull File outFile
    ) throws IOException {
        start(wordTemplateVersion, dataList, new FileOutputStream(outFile));
    }

    /**
     * 启动word读写操作
     */
    protected abstract void start();

    /**
     * 读取数据并进行分类
     *
     * @return 数据分类后
     */
    @NotNull
    protected final Map<String, List<ExcelEntity>> classification(Language language) {
        //进行数据筛选
        Map<String, List<ExcelEntity>> classificationMap = new LinkedHashMap<>(this.dataList.size() / 2);
        for (ExcelEntity data : this.dataList) {
            String classification = null;
            switch (language) {
                case chinese:
                    classification = data.getChineseBrandClassification();
                    break;
                case korean:
                    classification = data.getKoreanBrandClassification();
                    break;
                default:
                    break;
            }
            if (classificationMap.containsKey(classification)) {
                List<ExcelEntity> dataList = classificationMap.get(classification);
                dataList.add(data);
            } else {
                List<ExcelEntity> lists = new ArrayList<>();
                lists.add(data);
                classificationMap.put(classification, lists);
            }
        }
        return classificationMap;
    }

    /**
     * 语言
     */
    protected enum Language {
        chinese,
        korean
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

    @Data
    protected static class Bookmark {
        /**
         * 目标书签，文本为需要跳转的目标文本
         */
        private String target;
        /**
         * 返回书签，用户返回目录索引处
         */
        private String returnMark;

        @Contract(pure = true)
        public Bookmark(String target, String returnMark) {
            this.target = target;
            this.returnMark = returnMark;
        }
    }
}
