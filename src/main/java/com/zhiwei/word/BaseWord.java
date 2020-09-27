package com.zhiwei.word;

import com.zhiwei.exception.WordTemplateVersionException;
import com.zhiwei.word.wordversion.SASWord;
import com.zhiwei.word.wordversion.SCSImageWord;
import com.zhiwei.word.wordversion.SCSWord;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;

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
@SuppressWarnings({"SameParameterValue", "DuplicatedCode"})
public abstract class BaseWord extends XWPFDocument {
    protected static Map<String, Integer> headerMap = new LinkedHashMap<>();
    protected static List<List<Object>> dataList = new ArrayList<>();
    protected OutputStream wordOut;

    protected BaseWord(String path) throws IOException {
        super(Objects.requireNonNull(BaseWord.class.getClassLoader().getResourceAsStream(path)));
    }

    /**
     * 初始化word对象
     *
     * @param templateVersion word模板的版本
     * @param headerMap       excel文件标题
     * @param dataList        excel数据
     * @param wordOut         word文件输出流
     * @throws IOException                  io异常
     * @throws WordTemplateVersionException word模板的版本不支持
     */
    public static void start(
            WordTemplateVersion templateVersion,
            Map<String, Integer> headerMap,
            List<List<String>> dataList,
            OutputStream wordOut
    ) throws IOException {
        //转换标题
        headerMap.forEach((headerTitle, column) -> {
            BaseWord.headerMap.put(headerTitle, column);
        });
        //转换数据存储方式，并生成书签
        for (int i = 0; i < dataList.size(); i++) {
            List<Object> dataColumn = new ArrayList<>(dataList.get(i));
            dataColumn.add(new Bookmark("t" + i, "r" + i));
            BaseWord.dataList.add(dataColumn);
        }
        BaseWord baseWord;
        switch (templateVersion) {
            case SAS:
                baseWord = SASWord.init();
                break;
            case SCS:
                baseWord = SCSWord.init();
                break;
            case SCS_IMAGE:
                baseWord = SCSImageWord.init();
                break;
            default:
                throw new WordTemplateVersionException("不支持的版本：" + templateVersion.toString());
        }
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
     * 读取数据并进行分类
     *
     * @return 数据分类后
     */
    protected final Map<String, List<List<Object>>> classification(String classificationColumn) {
        //进行数据筛选
        Map<String, List<List<Object>>> classificationMap = new LinkedHashMap<>(BaseWord.dataList.size() / 2);
        Integer column = BaseWord.headerMap.get(classificationColumn);//获得筛选依据的列
        for (List<Object> data : BaseWord.dataList) {
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
    protected void createBookmark(XWPFParagraph paragraph, String bookmarkName) {
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
    protected final void createCells(XWPFTableRow tableRow, int number) {
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
}
