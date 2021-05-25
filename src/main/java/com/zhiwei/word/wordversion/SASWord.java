package com.zhiwei.word.wordversion;

import com.deepoove.poi.util.TableTools;
import com.zhiwei.config.ContentFont;
import com.zhiwei.util.Util;
import com.zhiwei.word.BaseWord;
import com.zhiwei.word.ExcelEntity;
import com.zhiwei.word.WordTemplateVersion;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.jetbrains.annotations.Contract;
import org.jetbrains.annotations.NotNull;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;

import java.awt.Color;
import java.io.IOException;
import java.math.BigInteger;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

/**
 * @author 朽木不可雕也
 * @date 2020/9/17 16:10
 * @week 星期四
 */
@SuppressWarnings({"DuplicatedCode", "ForLoopReplaceableByForEach"})
public class SASWord extends BaseWord {
    private SASWord() throws IOException {
        super(WordTemplateVersion.SAS.getTemplateFile());
    }

    @Override
    public void start() {
        //删除无韩文翻译的文本
        super.dataList.removeIf(dataColumn ->
                Objects.isNull(dataColumn.getKoreanTitle()));
        //创建开头三个空白段落
        for (int i = 0; i < 2; i++) {
            super.createParagraph();
        }

        //根据分类写出表格
        this.addTable();
    }

    /**
     * 根据分类写出表格
     */
    private void addTable() {
        Map<String, List<ExcelEntity>> classificationMap = super.classification(Language.korean);
        String tileBackgroundColor = Util.getColorString(new Color(0, 68, 129));//标题栏颜色
        BigInteger titleHeight = new BigInteger(Integer.toString((int) Util.getPixel(0.89D, Util.LengthUnit.centimeter)));//计算标题栏行高
        XWPFParagraph paragraph;
        Set<String> keySet = classificationMap.keySet();
        for (String classification : keySet) {
            List<ExcelEntity> dataList = classificationMap.get(classification);
            //创建表格
            XWPFTable table = super.createTable();
            table.setWidth((int) Util.getPixel(15.55D, Util.LengthUnit.centimeter));

            //创建表格的标题栏
            XWPFTableRow tableRow = table.getRow(0);
            tableRow.getCtRow().addNewTrPr().addNewTrHeight().setVal(titleHeight);
            //设置背景颜色
            //创建单元格
            super.createCells(tableRow, 1);
            XWPFTableCell cell = tableRow.getCell(0);
            cell.setColor(tileBackgroundColor);
            cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            paragraph = cell.getParagraphArray(0);
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            paragraph.setVerticalAlignment(TextAlignment.CENTER);
            XWPFRun xwpfRun = paragraph.createRun();
            xwpfRun.setText(classification);
            xwpfRun.setFontFamily("BatangChe");
            xwpfRun.setFontSize(ContentFont.四号.getPoundValue());
            xwpfRun.setBold(true);
            xwpfRun.setColor(Util.getColorString(Color.WHITE));
            TableTools.mergeCellsHorizonal(table, table.getNumberOfRows() - 1, 0, 1);

            //设置空行
            tableRow = table.createRow();
            //去边框
            super.createCells(tableRow, 0);
            tableRow.setHeight((int) Util.getPixel(0.5D, Util.LengthUnit.centimeter));

            //输入表格内容
            for (int i = 0; i < dataList.size(); i++) {
                tableRow = table.createRow();
                super.createCells(tableRow, 1);
                tableRow.setHeight((int) Util.getPixel(1.84D, Util.LengthUnit.centimeter));
                ExcelEntity dataColumn = dataList.get(i);
                cell = tableRow.getCell(0);
                cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                paragraph = cell.getParagraphArray(0);//使用单元格自带段落
                paragraph.setAlignment(ParagraphAlignment.BOTH);
                //获得生成的书签
                BaseWord.Bookmark bookmark = dataColumn.getBookmark();
                xwpfRun = paragraph.createHyperlinkRun("#" + bookmark.getTarget());

                //尝试获得韩文的标题
                String text = dataColumn.getKoreanTitle();
                //判断韩文的标题是否存在
                if (Objects.nonNull(text)) {
                    //存在韩文标题，输出韩文标题
                    xwpfRun.setText(text);
                    xwpfRun.setFontFamily("BatangChe");
                    xwpfRun.setFontSize(ContentFont.小四.getPoundValue());
                    xwpfRun.setColor(Util.getColorString(new Color(0, 81, 144)));
                    super.createBookmark(paragraph, bookmark.getReturnMark());
                    //输出中文标题
                    text = dataColumn.getTitle();
                    if (text != null) {
                        paragraph = cell.addParagraph();//在单元格中创建段落
                        paragraph.setAlignment(ParagraphAlignment.LEFT);
                        xwpfRun = paragraph.createRun();
                        xwpfRun.setFontFamily("微软雅黑");
                        xwpfRun.setFontSize(11);
                        xwpfRun.setColor(Util.getColorString(new Color(128, 128, 128)));
                        xwpfRun.setText(text);
                    }
                }

                cell = tableRow.getCell(1);
                cell.setWidth(Integer.toString((int) Util.getPixel(2.75D, Util.LengthUnit.centimeter)));
                cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                paragraph = cell.getParagraphArray(0);
                paragraph.setAlignment(ParagraphAlignment.RIGHT);
                xwpfRun = paragraph.createRun();
                text = dataColumn.getTime();
                text = Objects.isNull(text) ? "" : text;
                text = text.replaceAll("/", "-");
                xwpfRun.setText(text);
                xwpfRun.setFontFamily("Calibri");
                xwpfRun.setFontSize(ContentFont.小四.getPoundValue());
                xwpfRun.setColor(Util.getColorString(new Color(0, 81, 144)));
            }

            //表格间的空行
            super.createParagraph();
            super.createParagraph();
        }

        //创建空白段落
        SASWord.super.createParagraph();
        paragraph = SASWord.super.createParagraph();
        //分页
        paragraph.createRun().addBreak(BreakType.PAGE);
        super.createParagraph();

        List<ExcelEntity> dataList = super.dataList;
        Iterator<ExcelEntity> iterator = dataList.iterator();
        while (iterator.hasNext()) {
            ExcelEntity dataColumn = iterator.next();
            XWPFTable table = SASWord.super.createTable(5, 2);

            //去除边框
            CTTcBorders tcBorders = CTTcBorders.Factory.newInstance();
            tcBorders.addNewLeft().setVal(STBorder.NIL);
            tcBorders.addNewRight().setVal(STBorder.NIL);
            tcBorders.addNewTop().setVal(STBorder.NIL);
            tcBorders.addNewBottom().setVal(STBorder.NIL);
            for (XWPFTableRow tableRow : table.getRows()) {
                for (XWPFTableCell cell : tableRow.getTableCells()) {
                    cell.getCTTc().addNewTcPr().setTcBorders(tcBorders);
                }
            }

            XWPFTableRow tableRow = table.getRow(0);
            List<XWPFTableCell> tableCells = tableRow.getTableCells();
            XWPFTableCell cell = tableCells.get(0);
            paragraph = cell.getParagraphArray(0);
            XWPFRun xwpfRun = paragraph.createRun();
            xwpfRun.setFontFamily("BatangChe");
            xwpfRun.setFontSize(10);
            xwpfRun.setBold(true);
            xwpfRun.setText("뉴스 타이틀:");

            cell = tableCells.get(1);
            cell.setWidth(Integer.toString((int) Util.getPixel(13.28D, Util.LengthUnit.centimeter)));
            cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            paragraph = cell.getParagraphArray(0);
            paragraph.setAlignment(ParagraphAlignment.BOTH);
            xwpfRun = paragraph.createRun();
            xwpfRun.setFontFamily("BatangChe");
            xwpfRun.setFontSize(ContentFont.小四.getPoundValue());
            //获得韩文标题
            String text = (String) dataColumn.getKoreanTitle();
            if (Objects.nonNull(text)) {
                xwpfRun.setText(text);
                //创建超链接
                Bookmark bookmark = dataColumn.getBookmark();
                xwpfRun = paragraph.createHyperlinkRun("#" + bookmark.getReturnMark());
                xwpfRun.setText("返回首页");
                xwpfRun.setFontFamily("微软雅黑");
                xwpfRun.setFontSize(11);
                xwpfRun.setColor(Util.getColorString(new Color(0, 81, 144)));
                super.createBookmark(paragraph, bookmark.getTarget());
                //添加中文标题
                paragraph = cell.addParagraph();
                xwpfRun = paragraph.createRun();
                xwpfRun.setFontFamily("微软雅黑");
                xwpfRun.setFontSize(11);
                xwpfRun.setText(dataColumn.getTitle());
            }

            tableRow = table.getRow(1);
            tableCells = tableRow.getTableCells();
            cell = tableCells.get(0);
            cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            paragraph = cell.getParagraphArray(0);
            xwpfRun = paragraph.createRun();
            xwpfRun.setFontFamily("BatangChe");
            xwpfRun.setFontSize(10);
            xwpfRun.setBold(true);
            xwpfRun.setText("출판사:");

            cell = tableCells.get(1);
            paragraph = cell.getParagraphArray(0);
            xwpfRun = paragraph.createRun();
            xwpfRun.setFontFamily("Arial");
            xwpfRun.setFontSize(11);
            text = dataColumn.getChannelDomainName() + " / ";
            xwpfRun.setText(text);
            xwpfRun = paragraph.createRun();
            xwpfRun.setFontFamily("微软雅黑");
            xwpfRun.setFontSize(11);
            xwpfRun.setText(dataColumn.getChannel());

            tableRow = table.getRow(2);
            tableCells = tableRow.getTableCells();
            cell = tableCells.get(0);
            cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            paragraph = cell.getParagraphArray(0);
            xwpfRun = paragraph.createRun();
            xwpfRun.setFontFamily("BatangChe");
            xwpfRun.setFontSize(10);
            xwpfRun.setBold(true);
            xwpfRun.setText("기사날짜:");

            cell = tableCells.get(1);
            paragraph = cell.getParagraphArray(0);
            xwpfRun = paragraph.createRun();
            xwpfRun.setFontSize(11);
            xwpfRun.setFontFamily("Arial");
            text = dataColumn.getTime();
            text = Objects.isNull(text) ? "" : text;
            text = text.replaceAll("/", "-");
            xwpfRun.setText(text);

            tableRow = table.getRow(3);
            tableCells = tableRow.getTableCells();
            cell = tableCells.get(0);
            cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            paragraph = cell.getParagraphArray(0);
            xwpfRun = paragraph.createRun();
            xwpfRun.setFontFamily("BatangChe");
            xwpfRun.setFontSize(10);
            xwpfRun.setBold(true);
            xwpfRun.setText("개요:");
            TableTools.mergeCellsHorizonal(table, 3, 0, 1);

            TableTools.mergeCellsHorizonal(table, 4, 0, 1);
            tableRow = table.getRow(4);
            tableCells = tableRow.getTableCells();
            cell = tableCells.get(0);
            cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            paragraph = cell.getParagraphArray(0);
            paragraph.setAlignment(ParagraphAlignment.BOTH);
            text = dataColumn.getKoreanAbstract();
            if (Objects.nonNull(text)) {
                //设置悬挂缩进1.5字符保持与行首对齐
                CTPPr ctpPr = Objects.isNull(paragraph.getCTP().getPPr()) ? paragraph.getCTP().addNewPPr() : paragraph.getCTP().getPPr();
                CTInd ctInd = Objects.isNull(ctpPr.getInd()) ? ctpPr.addNewInd() : ctpPr.getInd();
                ctInd.setHangingChars(new BigInteger("150"));
                //设置段落编号
                CTAbstractNum ctAbstractNum = CTAbstractNum.Factory.newInstance();
                ctAbstractNum.setAbstractNumId(BigInteger.valueOf(0));
                CTLvl ctLvl = ctAbstractNum.addNewLvl();//创建一级编号
                ctLvl.addNewNumFmt().setVal(STNumberFormat.BULLET);
                ctLvl.addNewLvlText().setVal("-");
                XWPFAbstractNum abstractNum = new XWPFAbstractNum(ctAbstractNum);
                XWPFNumbering xwpfNumbering = super.createNumbering();
                BigInteger abstractNumID = xwpfNumbering.addAbstractNum(abstractNum);
                BigInteger numID = xwpfNumbering.addNum(abstractNumID);
                paragraph.setNumID(numID);
                //插入项目编号
                xwpfRun = paragraph.createRun();
                xwpfRun.setFontFamily("BatangChe");
                xwpfRun.setFontSize(ContentFont.小四.getPoundValue());
                xwpfRun.setText(text);
                //创建空白段落
                super.createParagraph();
                super.createParagraph();
                //输入中文
                paragraph = super.createParagraph();
                xwpfRun = paragraph.createRun();
                text = dataColumn.getTitle();
                xwpfRun.setFontSize(11);
                xwpfRun.setBold(true);
                xwpfRun.setFontFamily("微软雅黑");
                xwpfRun.setText(Objects.isNull(text) ? "" : text);

                paragraph = super.createParagraph();
                xwpfRun = paragraph.createRun();
                xwpfRun.setFontFamily("微软雅黑");
                xwpfRun.setFontSize(11);
                text = dataColumn.getTime();
                text = Objects.isNull(text) ? "" : text;
                text = text.replaceAll("/", "-") +
                        "来源于：" + (Objects.isNull(dataColumn.getChannel()) ?
                        "" : dataColumn.getChannel());
                xwpfRun.setText(text);

                text = dataColumn.getFullText();
                String[] texts = text.split("\n");
                for (String sentence : texts) {
                    paragraph = super.createParagraph();
                    paragraph.setAlignment(ParagraphAlignment.BOTH);//两端对齐
                    //设置首行缩进2个字符
                    ctpPr = Objects.isNull(paragraph.getCTP().getPPr()) ? paragraph.getCTP().addNewPPr() : paragraph.getCTP().getPPr();
                    ctInd = Objects.isNull(ctpPr.getInd()) ? ctpPr.addNewInd() : ctpPr.getInd();
                    ctInd.setFirstLineChars(new BigInteger("200"));
                    xwpfRun = paragraph.createRun();
                    xwpfRun.setFontFamily("微软雅黑");
                    xwpfRun.setFontSize(11);
                    xwpfRun.setText(sentence);
                }
            }
            for (XWPFTableRow xwpfTableRow : table.getRows()) {
                try {
                    xwpfTableRow.getCell(0).setWidth(Integer.toString((int) Util.getPixel(2.8D, Util.LengthUnit.centimeter)));
                } catch (NullPointerException ignored) {
                }
            }

            if (iterator.hasNext()) {
                //分页，并在页首添加空格
                paragraph.createRun().addBreak(BreakType.PAGE);
                super.createParagraph();
            }
        }
    }


    @NotNull
    @Contract(" -> new")
    public static SASWord init() throws IOException {
        return new SASWord();
    }
}
