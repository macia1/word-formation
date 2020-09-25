package com.zhiwei.word.wordversion;

import com.deepoove.poi.xwpf.XWPFParagraphWrapper;
import com.sun.istack.internal.NotNull;
import com.zhiwei.config.ContentFont;
import com.zhiwei.config.ExcelTitles;
import com.zhiwei.util.Util;
import com.zhiwei.word.BaseWord;
import com.zhiwei.word.Bookmark;
import com.zhiwei.word.WordTemplateVersion;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

import java.awt.Color;
import java.io.IOException;
import java.math.BigInteger;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * @author 朽木不可雕也
 * @date 2020/9/17 16:10
 * @week 星期四
 */
@SuppressWarnings("DuplicatedCode")
public class SCSWord extends BaseWord {
    private SCSWord() throws IOException {
        super(WordTemplateVersion.SCS.getTemplateFile());
    }

    @Override
    public void start() {
        //创建第一个段落
        XWPFParagraph paragraph = super.createParagraph();
        //设置段落居中
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        //创建文档标题
        paragraph = super.createParagraph();
        //设置标题居中
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        /*
        设置段落文本
         */
        XWPFRun title = paragraph.createRun();
        title.setFontFamily("Arial");//设置字体
        title.setFontSize(18);//设置字号（小二）
        title.setText("Daily News");//文本
        title.setBold(true);//粗体字
        //设置下划线
        CTUnderline underline = title.getCTR().getRPr().addNewU();
        underline.setVal(STUnderline.Enum.forInt(1));
        //写入日期
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy.MM.dd");
        paragraph = super.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);//居中
        XWPFRun content = paragraph.createRun();
        content.setFontFamily("Arial");
        content.setText(dateFormat.format(new Date()) + "（木）");
        content.setFontSize(14);
        //创建一个空行
        super.createParagraph();
        //开始写入文档的主体部分
        this.createWordBody();
    }

    /**
     * 创建word的主体部分
     */
    private void createWordBody() {
        //读取数据并进行分类筛选
        Map<String, List<List<Object>>> classificationMap = super.classification(ExcelTitles.中文品牌分类.name());
        classificationMap.forEach((classification, dataList) -> {
            //输出分类标题
            XWPFParagraph paragraph = super.createParagraph();
            XWPFRun xwpfRun = paragraph.createRun();
            xwpfRun.setFontSize(ContentFont.小四.getPoundValue());
            xwpfRun.setFontFamily("微软雅黑");
            xwpfRun.setText(classification);
            xwpfRun.setBold(true);
            //创建分类表格
            XWPFTable table = super.createTable(dataList.size() + 1, 4);
            table.setWidth((int) Util.getPixel(18.41D, Util.LengthUnit.centimeter));

            //标题栏插入标题
            List<XWPFTableRow> tableRows = table.getRows();//获得表格行
            XWPFTableRow titleRow = tableRows.get(0);
            //获得标题栏的单元格
            List<XWPFTableCell> tableCells = titleRow.getTableCells();
            String colorStr = Util.getColorString(new Color(99, 154, 206));
            this.setTableTitle(tableCells.get(0), "No", colorStr);
            this.setTableTitle(tableCells.get(1), "Headline", colorStr);
            this.setTableTitle(tableCells.get(2), "Media", colorStr);
            this.setTableTitle(tableCells.get(3), "Publish Date", colorStr);

            //将数据放入表格
            this.createTableBody(tableRows, dataList);
            //设置每个表格的宽度
            Map<Integer, Double> columnWidthMap = new HashMap<>();
            columnWidthMap.put(0, 0.95D);
            columnWidthMap.put(1, 10.6D);
            columnWidthMap.put(2, 3.69D);
            columnWidthMap.put(3, 3.17D);
            columnWidthMap.forEach((column, columnWidth) -> {
                for (XWPFTableRow tableRow : table.getRows()) {
                    tableRow.getCell(column).setWidth(Integer.toString((int) Util.getPixel(columnWidth, Util.LengthUnit.centimeter)));
                }
            });
            super.createParagraph();
        });
        //目录表格写入完毕，写入全部数据
        Map<String, Integer> headerMap = super.monitor.getHeaders();
        super.monitor.getDataList().forEach(dataColumn -> {
            Bookmark bookmark = (Bookmark) dataColumn.get(dataColumn.size() - 1);
            XWPFParagraph paragraph = super.createParagraph();
            paragraph.createRun().addBreak(BreakType.PAGE);
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            CTPPr ctpPr = Objects.isNull(paragraph.getCTP().getPPr()) ? paragraph.getCTP().addNewPPr() : paragraph.getCTP().getPPr();
            CTInd ctInd = Objects.isNull(ctpPr.getInd()) ? ctpPr.addNewInd() : ctpPr.getInd();
            ctInd.setFirstLineChars(new BigInteger("200"));
            super.createBookmark(paragraph, bookmark.getTarget());
            this.createHyperlink(paragraph, "返回首页",
                    bookmark.getReturnMark(), "微软雅黑", ContentFont.小五);

            //写出Headline:
            String title = "Headline:  ";
            paragraph = super.createParagraph();
            XWPFRun xwpfRun = paragraph.createRun();
            xwpfRun.setBold(true);
            xwpfRun.setFontFamily("Arial");
            xwpfRun.setFontSize(ContentFont.小五.getPoundValue());
            xwpfRun.setText(title);
            String text = (String) dataColumn.get(headerMap.get(ExcelTitles.韩文标题.name()));
            if (Objects.nonNull(text)) {
                xwpfRun = paragraph.createRun();
                xwpfRun.setFontFamily("BatangChe");
                xwpfRun.setFontSize(ContentFont.小五.getPoundValue());
                xwpfRun.setText(text);
                //写出中文标题
                text = (String) dataColumn.get(headerMap.get(ExcelTitles.标题.name()));
                if (Objects.nonNull(text)) {
                    paragraph = super.createParagraph();
                    xwpfRun = paragraph.createRun();
                    StringBuilder string = new StringBuilder();
                    int length = 11;//增加空格保持文本对齐
                    for (int i = 0; i < length; i++) {
                        string.append(" ");
                    }
                    string.append(text);
                    xwpfRun.setFontFamily("微软雅黑");
                    xwpfRun.setFontSize(ContentFont.小五.getPoundValue());
                    xwpfRun.setText(string.toString());
                }
            } else {
                //没有韩文输出中文
                text = (String) dataColumn.get(headerMap.get(ExcelTitles.标题.name()));
                if (Objects.nonNull(text)) {
                    xwpfRun = paragraph.createRun();
                    xwpfRun.setFontFamily("微软雅黑");
                    xwpfRun.setFontSize(ContentFont.小五.getPoundValue());
                    xwpfRun.setText(text);
                }
            }

            title = "Publication: ";
            paragraph = super.createParagraph();
            xwpfRun = paragraph.createRun();
            xwpfRun.setFontSize(ContentFont.小五.getPoundValue());
            xwpfRun.setFontFamily("Arial");
            xwpfRun.setBold(true);
            xwpfRun.setText(title);
            xwpfRun = paragraph.createRun();
            xwpfRun.setFontFamily("Arial");
            xwpfRun.setFontSize(ContentFont.小五.getPoundValue());
            xwpfRun.setText(dataColumn.get(headerMap.get(ExcelTitles.渠道域名.name())) + " ");
            xwpfRun = paragraph.createRun();
            xwpfRun.setFontFamily("微软雅黑");
            xwpfRun.setFontSize(ContentFont.小五.getPoundValue());
            xwpfRun.setText("/ " + dataColumn.get(headerMap.get(ExcelTitles.渠道.name())));

            title = "Paper Date: ";
            paragraph = super.createParagraph();
            xwpfRun = paragraph.createRun();
            xwpfRun.setFontSize(ContentFont.小五.getPoundValue());
            xwpfRun.setFontFamily("Arial");
            xwpfRun.setBold(true);
            xwpfRun.setText(title);
            xwpfRun = paragraph.createRun();
            xwpfRun.setFontFamily("Arial");
            xwpfRun.setFontSize(ContentFont.小五.getPoundValue());
            xwpfRun.setText(((String) dataColumn.get(headerMap.get(ExcelTitles.时间.name()))).replaceAll("/", "."));

            title = "Summary: ";
            paragraph = super.createParagraph();
            xwpfRun = paragraph.createRun();
            xwpfRun.setFontSize(ContentFont.小五.getPoundValue());
            xwpfRun.setFontFamily("Arial");
            xwpfRun.setBold(true);
            xwpfRun.setText(title);

            //韩文
            text = (String) dataColumn.get(headerMap.get(ExcelTitles.韩文摘要.name()));
            if (Objects.nonNull(text)) {
                String[] texts = text.split("\n");
                for (String sentence : texts) {
                    paragraph = super.createParagraph();
                    xwpfRun = paragraph.createRun();
                    xwpfRun.setFontFamily("BatangChe");
                    xwpfRun.setFontSize(ContentFont.小五.getPoundValue());
                    xwpfRun.setText(sentence);
                }

                super.createParagraph();
                super.createParagraph();//中间换行

                //中文
                //标题
                paragraph = super.createParagraph();
                xwpfRun = paragraph.createRun();
                xwpfRun.setFontSize(ContentFont.小五.getPoundValue());
                xwpfRun.setFontFamily("微软雅黑");
                text = (String) dataColumn.get(headerMap.get(ExcelTitles.标题.name()));
                text = Objects.isNull(text) ? "" : text;
                xwpfRun.setText(text);

                //日期和来源
                paragraph = super.createParagraph();
                xwpfRun = paragraph.createRun();
                xwpfRun.setFontFamily("微软雅黑");
                xwpfRun.setFontSize(ContentFont.小五.getPoundValue());
                text = (String) dataColumn.get(headerMap.get(ExcelTitles.时间.name()));
                text = text.replaceAll("/", "-");
                text = text + " 来源于：" + dataColumn.get(headerMap.get(ExcelTitles.渠道.name()));
                xwpfRun.setText(text);

                //正文
                text = (String) dataColumn.get(headerMap.get(ExcelTitles.全文.name()));
                text = Objects.isNull(text) ? "" : text;
                texts = text.split("\n");
                for (String sentence : texts) {
                    paragraph = super.createParagraph();
                    paragraph.setAlignment(ParagraphAlignment.BOTH);
                    //设置首行缩进2个字符
                    ctpPr = Objects.isNull(paragraph.getCTP().getPPr()) ? paragraph.getCTP().addNewPPr() : paragraph.getCTP().getPPr();
                    ctInd = Objects.isNull(ctpPr.getInd()) ? ctpPr.addNewInd() : ctpPr.getInd();
                    ctInd.setFirstLineChars(new BigInteger("200"));
                    xwpfRun = paragraph.createRun();
                    xwpfRun.setFontFamily("微软雅黑");
                    xwpfRun.setFontSize(ContentFont.小五.getPoundValue());
                    xwpfRun.setText(sentence);
                }
            } else {
                //正文
                text = (String) dataColumn.get(headerMap.get(ExcelTitles.全文.name()));
                text = Objects.isNull(text) ? "" : text;
                String[] texts = text.split("\n");
                for (String sentence : texts) {
                    paragraph = super.createParagraph();
                    paragraph.setAlignment(ParagraphAlignment.BOTH);
                    //设置首行缩进2个字符
                    ctpPr = Objects.isNull(paragraph.getCTP().getPPr()) ? paragraph.getCTP().addNewPPr() : paragraph.getCTP().getPPr();
                    ctInd = Objects.isNull(ctpPr.getInd()) ? ctpPr.addNewInd() : ctpPr.getInd();
                    ctInd.setFirstLineChars(new BigInteger("200"));
                    xwpfRun = paragraph.createRun();
                    xwpfRun.setFontFamily("微软雅黑");
                    xwpfRun.setFontSize(ContentFont.小五.getPoundValue());
                    xwpfRun.setText(sentence);
                }
            }
        });
    }

    /**
     * 将数据放入表格
     *
     * @param tableRows 表格行
     * @param dataList  数据列表
     */
    private void createTableBody(@NotNull List<XWPFTableRow> tableRows, @NotNull List<List<Object>> dataList) {
        Map<String, Integer> headerMap = super.monitor.getHeaders();
        for (int i = 1; i < tableRows.size(); i++) {
            XWPFTableRow tableRow = tableRows.get(i);
            tableRow.setHeight((int) Util.getPixel(1.11D, Util.LengthUnit.centimeter));
            List<XWPFTableCell> tableCells = tableRow.getTableCells();//获得行中的单元格
            List<Object> dataColumn = dataList.get(i - 1);//获得对应的一行数据
            try {
                int index = 0;
                //设置No列
                XWPFTableCell tableCell = tableCells.get(index);
                XWPFParagraph paragraph = tableCell.getParagraphArray(0);
                XWPFRun xwpfRun = paragraph.createRun();
                xwpfRun.setFontFamily("Arial");
                xwpfRun.setFontSize(ContentFont.小五.getPoundValue());
                xwpfRun.setText((String) dataColumn.get(headerMap.get(ExcelTitles.序号.name())));

                //设置Headline列
                index++;
                String cellContent = (String) dataColumn.get(headerMap.get(ExcelTitles.韩文标题.name()));//韩文标题
                tableCell = tableCells.get(index);
                paragraph = tableCell.getParagraphArray(0);
                Bookmark bookmark = (Bookmark) dataColumn.get(dataColumn.size() - 1);//获得书签信息
                //韩文标题超链接
                this.createHyperlink(paragraph, cellContent, bookmark.getTarget(), "BatangChe", ContentFont.小五);
                paragraph.setAlignment(ParagraphAlignment.LEFT);//设置左对齐，防止换行出现混乱
                paragraph.getRuns().get(0).addBreak();//换行
                //中文标题超链接
                cellContent = (String) dataColumn.get(headerMap.get(ExcelTitles.标题.name()));//中文标题
                this.createHyperlink(paragraph, cellContent, bookmark.getTarget(), "微软雅黑", ContentFont.小五);
                super.createBookmark(paragraph, bookmark.getReturnMark());

                //设置Media Name列
                index++;
                tableCell = tableCells.get(index);
                cellContent = (String) dataColumn.get(headerMap.get(ExcelTitles.渠道域名.name()));
                paragraph = tableCell.getParagraphArray(0);
                xwpfRun = paragraph.createRun();
                xwpfRun.setText(cellContent);
                xwpfRun.setFontFamily("Arial");
                xwpfRun.setFontSize(ContentFont.小五.getPoundValue());

                paragraph = tableCell.addParagraph();
                xwpfRun = paragraph.createRun();
                cellContent = (String) dataColumn.get(headerMap.get(ExcelTitles.渠道.name()));
                xwpfRun.setText(cellContent);
                xwpfRun.setFontFamily("微软雅黑");
                xwpfRun.setFontSize(ContentFont.小五.getPoundValue());

                //设置Publish Date
                index++;
                tableCell = tableCells.get(index);
                cellContent = (String) dataColumn.get(headerMap.get(ExcelTitles.时间.name()));
                cellContent = cellContent.replaceAll("/", ".");
                paragraph = tableCell.getParagraphArray(0);
                xwpfRun = paragraph.createRun();
                xwpfRun.setText(cellContent);
                xwpfRun.setFontFamily("Arial");
                xwpfRun.setFontSize(ContentFont.小五.getPoundValue());
            } catch (Exception e) {
                e.printStackTrace();
                System.exit(500);
            }
        }
    }

    /**
     * 设置表格的标题
     *
     * @param tableTitle 表格标题单元格
     * @param content    标题
     */
    private void setTableTitle(@NotNull XWPFTableCell tableTitle, String content, String colorStr) {
        tableTitle.setColor(colorStr);
        //表格内容居中
        CTTc cttc = tableTitle.getCTTc();
        CTTcPr ctPr = cttc.addNewTcPr();
        ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
        cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);

        XWPFParagraph paragraph = tableTitle.getParagraphArray(0);//获得表格的段落
        XWPFRun xwpfRun = paragraph.createRun();
        xwpfRun.setFontFamily("Arial");
        xwpfRun.setFontSize(ContentFont.小五.getPoundValue());
        xwpfRun.setBold(true);
        xwpfRun.setText(content);
        xwpfRun.setColor("FFFFFF");
    }

    /**
     * 插入一个链接到本文档某个位置的超链接
     *
     * @param paragraph 超链接段落
     * @param linkText  超链接显示的文本
     * @param action    锚点名称（书签名称）
     */
    public void createHyperlink(@NotNull XWPFParagraph paragraph, String linkText, String action, String fontName, @NotNull ContentFont contentFont) {
        XWPFParagraphWrapper wrapper = new XWPFParagraphWrapper(paragraph);
        XWPFRun hyperRun = wrapper.insertNewHyperLinkRun(paragraph.createRun(), "#" + action);
        hyperRun.setText(linkText);
        hyperRun.setFontFamily(fontName);
        hyperRun.setFontSize(contentFont.getPoundValue());
        hyperRun.setColor(Util.getColorString(new Color(5, 99, 193)));
        hyperRun.setUnderline(UnderlinePatterns.SINGLE);
    }

    @NotNull
    public static SCSWord init() throws IOException {
        return new SCSWord();
    }
}
