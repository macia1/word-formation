package com.zhiwei.word.wordversion;

import com.zhiwei.word.BaseWord;
import com.zhiwei.word.WordTemplateVersion;

import java.io.IOException;

/**
 * @author 朽木不可雕也
 * @date 2020/9/17 16:11
 * @week 星期四
 */
public class SCSImageWord extends BaseWord {
    private SCSImageWord() throws IOException {
        super(WordTemplateVersion.SCS_IMAGE.getTemplateFile());
    }

    @Override
    public void start() {

    }

    public static SCSImageWord init() throws IOException {
        return new SCSImageWord();
    }
}
