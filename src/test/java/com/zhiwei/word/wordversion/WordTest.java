package com.zhiwei.word.wordversion;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.zhiwei.word.BaseWord;
import com.zhiwei.word.ExcelEntity;
import com.zhiwei.word.WordTemplateVersion;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Before;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @author 朽木不可雕也
 * @date 2020/9/20 16:49
 * @week 星期日
 */
@Log4j2
public class WordTest {
    public static class Monitor extends AnalysisEventListener<ExcelEntity> {
        private final List<ExcelEntity> dataList = new ArrayList<>();

        @Override
        public void invoke(ExcelEntity excelEntity, AnalysisContext analysisContext) {
            dataList.add(excelEntity);
            log.info(excelEntity);
        }

        @Override
        public void doAfterAllAnalysed(AnalysisContext analysisContext) {
            log.info("读取完毕");
        }
    }

    private List<ExcelEntity> dataList;

    @Before
    public void before() {
        Monitor monitor = new Monitor();
        EasyExcel.read("testFiles/数据上传模板.xlsx", ExcelEntity.class, monitor).sheet(1).doRead();
        this.dataList = monitor.dataList;
    }

    @Test
    public void sasWord() throws IOException {
        //传入数据
        BaseWord.start(
                //需要生成的模板版本，为枚举类：com.zhiwei.word.WordTemplateVersion
                //目前支持SAS和SCS
                WordTemplateVersion.SAS,
                //excel数据
                this.dataList,
                //输出流
                new FileOutputStream("testWrite/write-sas.docx")
        );
        //传入数据
        BaseWord.start(
                //需要生成的模板版本，为枚举类：com.zhiwei.word.WordTemplateVersion
                //目前支持SAS和SCS
                WordTemplateVersion.SCS,
                //excel数据
                this.dataList,
                //输出流
                new FileOutputStream("testWrite/write-scs.docx")
        );
    }

    @Test
    public void scsWord() throws IOException {
        //传入数据
        BaseWord.start(
                //需要生成的模板版本，为枚举类：com.zhiwei.word.WordTemplateVersion
                //目前支持SAS和SCS
                WordTemplateVersion.SCS,
                //excel数据
                this.dataList,
                //输出流
                new FileOutputStream("testWrite/write-scs.docx")
        );
    }
}
