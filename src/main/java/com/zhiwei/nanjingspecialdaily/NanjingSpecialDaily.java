package com.zhiwei.nanjingspecialdaily;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.event.SyncReadListener;
import com.zhiwei.bossdirectrecruitmentsummaryautomation.RecruitmentSummaryException;
import com.zhiwei.config.ContentFont;
import com.zhiwei.util.UrlUtil;
import lombok.extern.log4j.Log4j2;
import org.apache.commons.lang3.time.FastDateFormat;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.Collections;
import java.util.List;
import java.util.Locale;
import java.util.stream.Collectors;

import static java.util.Objects.isNull;
import static org.apache.commons.lang3.StringUtils.isBlank;
import static org.apache.commons.lang3.StringUtils.isNoneBlank;

/**
 * 南京专项日报
 *
 * @author aszswaz
 * @date 2021/5/28 14:45:04
 * @IDE IntelliJ IDEA
 */
@Log4j2
@SuppressWarnings("JavaDoc")
public class NanjingSpecialDaily extends XWPFDocument {
    private final List<NanjingSpecialDailyEntity> nanjingSpecialDailyEntities;
    private final String writePath;
    private static final FastDateFormat FAST_DATE_FORMAT = FastDateFormat.getInstance("yyy-MM-dd");

    private NanjingSpecialDaily(List<NanjingSpecialDailyEntity> nanjingSpecialDailyEntities, String writePath) {
        this.nanjingSpecialDailyEntities = Collections.unmodifiableList(nanjingSpecialDailyEntities);
        this.writePath = writePath;
    }

    public static void start(String sourcePath, String writePath) throws NanjingSpecialDailyException, IOException {
        List<NanjingSpecialDailyEntity> nanjingSpecialDailyEntities = EasyExcel.read(sourcePath, NanjingSpecialDailyEntity.class, new SyncReadListener()).sheet(0).doReadSync();
        if (isNull(nanjingSpecialDailyEntities) || nanjingSpecialDailyEntities.isEmpty()) {
            throw new NanjingSpecialDailyException(sourcePath + ": 这是一个空文件！");
        }

        // 过滤空白行
        nanjingSpecialDailyEntities = nanjingSpecialDailyEntities.stream().filter(nanjingSpecialDailyEntity ->
                isNoneBlank(nanjingSpecialDailyEntity.getTitle())
                        && isNoneBlank(nanjingSpecialDailyEntity.getUrl())
                        && isNoneBlank(nanjingSpecialDailyEntity.getWord())
                        && isNoneBlank(nanjingSpecialDailyEntity.getText())
        ).collect(Collectors.toList());

        checked(nanjingSpecialDailyEntities);

        // 过滤不需要的url
        nanjingSpecialDailyEntities = nanjingSpecialDailyEntities.stream().filter(nanjingSpecialDailyEntity -> {
            for (String filterUrl : UrlConfig.filterUrls) {
                if (nanjingSpecialDailyEntity.getUrl().toLowerCase(Locale.ROOT).contains(filterUrl.toLowerCase(Locale.ROOT))) {
                    return false;
                }
            }
            return true;
        }).collect(Collectors.toList());

        NanjingSpecialDaily nanjingSpecialDaily = new NanjingSpecialDaily(nanjingSpecialDailyEntities, writePath);
        nanjingSpecialDaily.generateWord();
        nanjingSpecialDaily.write();
    }

    /**
     * 数据检查
     */
    private static void checked(List<NanjingSpecialDailyEntity> nanjingSpecialDailyEntities) throws NanjingSpecialDailyException {
        for (NanjingSpecialDailyEntity nanjingSpecialDailyEntity : nanjingSpecialDailyEntities) {
            if (isBlank(nanjingSpecialDailyEntity.getTitle())) {
                throw new NanjingSpecialDailyException("聚合标题不能为空");
            }
            if (isBlank(nanjingSpecialDailyEntity.getText())) {
                nanjingSpecialDailyEntity.setText("");
            }
            if (isBlank(nanjingSpecialDailyEntity.getUrl())) {
                throw new NanjingSpecialDailyException("地址不能为空");
            }
            if (isBlank(nanjingSpecialDailyEntity.getWord())) {
                throw new NanjingSpecialDailyException("命中词不能为空");
            }
        }
    }

    /**
     * 生成word
     */
    private void generateWord() throws NanjingSpecialDailyException, UnsupportedEncodingException {
        int number = 1;
        for (NanjingSpecialDailyEntity nanjingSpecialDailyEntity : this.nanjingSpecialDailyEntities) {
            this.title(nanjingSpecialDailyEntity, number++);
            this.summary(nanjingSpecialDailyEntity);
            this.address(nanjingSpecialDailyEntity);
        }
    }

    /**
     * 标题
     */
    private void title(NanjingSpecialDailyEntity nanjingSpecialDailyEntity, int number) {
        String titleTemplate = "%d、【%s】%s";

        String title = nanjingSpecialDailyEntity.getTitle();
        String channel = nanjingSpecialDailyEntity.getChannel();

        XWPFParagraph paragraph = super.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun xwpfRun = paragraph.createRun();
        xwpfRun.setText(String.format(titleTemplate, number, channel, title));
        xwpfRun.setFontFamily("微软雅黑");
        xwpfRun.setFontSize(ContentFont.小四.getPoundValue());
        xwpfRun.setBold(true);
    }

    /**
     * 摘要
     */
    private void summary(NanjingSpecialDailyEntity nanjingSpecialDailyEntity) {
        final String word = nanjingSpecialDailyEntity.getWord();
        final String text = nanjingSpecialDailyEntity.getText();
        final StringBuilder summary = new StringBuilder();

        int wordIndex = text.indexOf(word);
        // 文本截取范围
        if (wordIndex > 0) {
            int startIndex, endIndex;
            startIndex = Math.max(wordIndex - 20, 0);
            // 寻找关键词之后的第二个句号
            endIndex = text.indexOf('。', wordIndex);
            if (endIndex > 0) {
                endIndex = text.indexOf('。', endIndex + 1);
            }
            // 摘要最长只能有100个字
            if (endIndex <= 0 || endIndex - startIndex > 100) {
                endIndex = Math.min(startIndex + 100, text.length() - 1);
            }
            summary.append(text, startIndex, endIndex + 1);
        }

        summary.append("（").append(FAST_DATE_FORMAT.format(nanjingSpecialDailyEntity.getTime())).append("）");

        XWPFParagraph paragraph = super.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("摘要：");
        run.setFontSize(ContentFont.小四.getPoundValue());
        run.setBold(true);
        run.setFontFamily("微软雅黑");

        run = paragraph.createRun();
        run.setText(summary.toString().replaceAll("\u3000", ""));
        run.setFontSize(ContentFont.小四.getPoundValue());
        run.setFontFamily("微软雅黑");
    }

    /**
     * 地址
     */
    private void address(NanjingSpecialDailyEntity nanjingSpecialDailyEntity) throws UnsupportedEncodingException, NanjingSpecialDailyException {
        try {
            String address = nanjingSpecialDailyEntity.getUrl();
            address = UrlUtil.escapeUrl(address);
            XWPFRun xwpfRun = super.createParagraph().createHyperlinkRun(address);
            xwpfRun.setText(address);
            xwpfRun.setFontFamily("微软雅黑");
            xwpfRun.setFontSize(12);
            xwpfRun.setUnderline(UnderlinePatterns.SINGLE);
        } catch (RecruitmentSummaryException recruitmentSummaryException) {
            throw new NanjingSpecialDailyException(recruitmentSummaryException.getMessage());
        }
    }

    private void write() throws IOException {
        try (OutputStream outputStream = new FileOutputStream(this.writePath)) {
            super.write(outputStream);
        }
    }
}
