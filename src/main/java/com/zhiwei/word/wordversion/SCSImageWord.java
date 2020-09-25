package com.zhiwei.word.wordversion;

import com.zhiwei.util.Util;
import com.zhiwei.word.BaseWord;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.jetbrains.annotations.Contract;
import org.jetbrains.annotations.NotNull;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * @author 朽木不可雕也
 * @date 2020/9/17 16:11
 * @week 星期四
 */
public class SCSImageWord extends BaseWord {
    private SCSImageWord() throws IOException {
        super(null);
    }

    @Override
    public void start() {
        try {
            XWPFParagraph paragraph = super.createParagraph();
            XWPFRun xwpfRun = paragraph.createRun();
            String path = "resources/scs_image/scs_daily_news.png";
            int width = (int) Util.getPixel(19D, Util.LengthUnit.centimeter);
            int height = (int) Util.getPixel(3.18D, Util.LengthUnit.centimeter);
            xwpfRun.addPicture(new FileInputStream(path), XWPFDocument.PICTURE_TYPE_PNG, "scs_daily_news.png", Units.toEMU(256), Units.toEMU(256));
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    @NotNull
    @Contract(" -> new")
    public static SCSImageWord init() throws IOException {
        return new SCSImageWord();
    }
}
