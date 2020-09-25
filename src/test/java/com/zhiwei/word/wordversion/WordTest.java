package com.zhiwei.word.wordversion;

import com.zhiwei.word.BaseWord;
import com.zhiwei.word.WordTemplateVersion;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author 朽木不可雕也
 * @date 2020/9/20 16:49
 * @week 星期日
 */
public class WordTest {
    @Test
    public void sasWord() throws IOException, BaseWord.ExcelTitleIncompleteException {
        BaseWord.start(
                WordTemplateVersion.SAS,
                new FileInputStream("testFiles/数据上传模板.xlsx"),
                new FileOutputStream("testWrite/write-sas.docx")
        );
    }

    @Test
    public void scsWord() throws IOException, BaseWord.ExcelTitleIncompleteException {
        BaseWord.start(
                WordTemplateVersion.SCS,
                new FileInputStream("testFiles/数据上传模板.xlsx"),
                new FileOutputStream("testWrite/write-scs.docx")
        );
    }
}
