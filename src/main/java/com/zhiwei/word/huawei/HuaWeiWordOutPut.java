package com.zhiwei.word.huawei;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.zhiwei.util.Util;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;

/**
 * @author aszswaz
 * @date 2021/4/2 11:20:41
 */
@Log4j2
public class HuaWeiWordOutPut {
    private Map<String, Integer> titleMap;

    private XWPFDocument doc;

    private long topId;

    /**
     * 文件输出
     *
     * @param dataJson  转换数据固定格式
     * @param writePath 输出路径
     */
    public void wordWrite(JSONObject dataJson, String writePath) throws Exception {
        if (!writePath.endsWith(".docx")) throw new Exception("文件格式错误");
        if (Objects.isNull(dataJson)) throw new NullPointerException("传入数据不能为空");
        String style = dataJson.getString("style");
        String match = HuaWeiStyleEnum.match(style);
        JSONArray data = dataJson.getJSONArray("data");
        if (Objects.nonNull(match)) {
            doc = new XWPFDocument();
            // 进行数据格式转换
            dataPut(data);
            // 目录生成
            tableOfContent();
            switch (match) {
                case "格式一":
                    styleOneDecoration(data);
                    break;
                case "格式二":
                    styleTwoDecoration(data);
                    break;
                default:
                    throw new NullPointerException("传输文件格式不存在" + style);
            }
        } else {
            throw new NullPointerException("传输文件格式不存在" + style);
        }

        // 文件写出
        try (FileOutputStream out = new FileOutputStream(writePath)) {
            doc.write(out);
        } catch (IOException e) {
            log.info("文件写出失败....");
            e.printStackTrace();
        } finally {
            doc.close();
        }
    }

    /**
     * 格式一 输出
     */
    private void styleOneDecoration(JSONArray data) {
        // 正文内容生成
        contentStyleOneDecoration(data);
    }

    /**
     * 格式二 输出
     */
    private void styleTwoDecoration(JSONArray data) {
        // 正文内容生成
        contentStyleTwoDecoration(data);
    }

    /**
     * 数据转换
     */
    private void dataPut(JSONArray data) {
        Map<String, Integer> linkedHashMap = new LinkedHashMap<>();
        for (int i = 0; i < data.size(); i++) {
            JSONObject jsonObject = data.getJSONObject(i);
            String source = jsonObject.getString("source");
            Integer count = jsonObject.getInteger("count");
            linkedHashMap.put(source, count);
            if (i == 0) {
                topId = getBookId(source) / 1000;
            }
        }
        titleMap = linkedHashMap;
    }


    /**
     * 生成bookId
     */
    private long getBookId(String key) {
        long bookId = key.hashCode();
        if (bookId < 0) {
            bookId = bookId * (-10);
        }
        return bookId / 100;
    }

    /**
     * 目录生成
     */
    private void tableOfContent() {
        boolean first = true;
        for (Map.Entry<String, Integer> entry : titleMap.entrySet()) {
            // 标题生成
            String key = entry.getKey();
            XWPFParagraph paragraph = doc.createParagraph();
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            XWPFHyperlinkRun hyperlinkRun = paragraph.createHyperlinkRun("#" + getBookId(key));
            hyperlinkRun.setTextPosition(20);
            hyperlinkRun.setColor(Util.getColorString(new Color(5, 99, 193)));
            hyperlinkRun.setFontFamily("Microsoft YaHei");
            hyperlinkRun.setFontSize(9);
            hyperlinkRun.setBold(true);
            hyperlinkRun.setText(key + "（共计" + entry.getValue() + "条）");
            if (first) {
                // 添加书签
                addBookmark(paragraph, String.valueOf(topId));
                first = false;
            }
        }
        doc.createParagraph();
        doc.createParagraph();
    }

    /**
     * 添加书签
     */
    private void addBookmark(XWPFParagraph paragraph, String bookId) {
        // 添加书签
        CTBookmark ctBookmark = paragraph.getCTP().addNewBookmarkStart();
        ctBookmark.setName(bookId);
        ctBookmark.setId(BigInteger.valueOf(0));
        paragraph.getCTP().addNewBookmarkEnd().setId(BigInteger.valueOf(0));
    }

    /**
     * 通用段落设置
     *
     * @param paragraph 段落
     * @return run
     */
    private XWPFRun commonDecoration(XWPFParagraph paragraph, String text, boolean isBold) {
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun run = paragraph.createRun();
        run.setFontFamily("Microsoft YaHei");
        run.setFontSize(9);
        run.setText(text);
        if (isBold) {
            run.setBold(true);
        }
        return run;
    }

    /**
     * 超链接通用配置
     */
    private void hyperlinkDecoration(XWPFParagraph paragraph, String url) {
        if (url.startsWith("http")) {
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            try {
                if (url.contains("\"")) {
                    url = url.replaceAll("\"", "%22");
                }
                XWPFHyperlinkRun hyperlinkRun = paragraph.createHyperlinkRun(url);
                hyperlinkRun.setFontFamily("Microsoft YaHei");
                hyperlinkRun.setFontSize(9);
                hyperlinkRun.setColor(Util.getColorString(new Color(5, 99, 193)));
                hyperlinkRun.setUnderline(UnderlinePatterns.SINGLE);
                hyperlinkRun.setText(url);
            } catch (Exception e) {
                commonUrlDecoration(paragraph, url);
                log.error("链接创建失败：{}", url);
                e.printStackTrace();
            }
        } else {
            commonUrlDecoration(paragraph, url);
        }
    }


    private void commonUrlDecoration(XWPFParagraph paragraph, String url) {
        XWPFRun run = commonDecoration(paragraph, url, false);
        run.setColor(Util.getColorString(new Color(5, 99, 193)));
        run.setUnderline(UnderlinePatterns.SINGLE);
    }


    /**
     * 进行内容生成
     */
    private void contentStyleOneDecoration(JSONArray dataArray) {
        for (int i = 0; i < dataArray.size(); i++) {
            JSONObject data = dataArray.getJSONObject(i);
            XWPFParagraph titleParagraph = doc.createParagraph();
            String source = data.getString("source");
            addBookmark(titleParagraph, String.valueOf(getBookId(source)));
            commonDecoration(titleParagraph, source + "（共计" + data.getInteger("count") + "条）", true);
            JSONArray urls = data.getJSONArray("urls");
            for (int j = 0; j < urls.size(); j++) {
                JSONObject urlJson = urls.getJSONObject(j);
                // rootSource
                XWPFParagraph rootSourceParagraph = doc.createParagraph();
                String rootSource = urlJson.getString("rootSource");
                commonDecoration(rootSourceParagraph, rootSource, false);
                // url
                String url = urlJson.getString("url");
                XWPFParagraph urlParagraph = doc.createParagraph();
                commonDecoration(urlParagraph, "链接：", false);
                hyperlinkDecoration(urlParagraph, url);
                if (j != urls.size() - 1) {
                    doc.createParagraph();
                }
            }
            // 回到顶部
            topForward();
            if (i != dataArray.size() - 1) {
                doc.createParagraph();
                doc.createParagraph();
            }
        }
    }

    /**
     * 进行内容生成
     */
    private void contentStyleTwoDecoration(JSONArray dataArray) {
        for (int i = 0; i < dataArray.size(); i++) {
            JSONObject data = dataArray.getJSONObject(i);
            XWPFParagraph titleParagraph = doc.createParagraph();
            String source = data.getString("source");
            commonDecoration(titleParagraph, source, true);
            addBookmark(titleParagraph, String.valueOf(getBookId(source)));
            doc.createParagraph();
            JSONArray urls = data.getJSONArray("urls");
            for (int j = 0; j < urls.size(); j++) {
                JSONObject urlJson = urls.getJSONObject(j);
                // rootSource
                XWPFParagraph rootSourceParagraph = doc.createParagraph();
                String rootSource = urlJson.getString("rootSource");
                commonDecoration(rootSourceParagraph, rootSource + "：", false);
                // url
                String url = urlJson.getString("url");
                hyperlinkDecoration(rootSourceParagraph, url);
            }
            // 回到顶部
            topForward();
            if (i != dataArray.size() - 1) {
                doc.createParagraph();
                doc.createParagraph();
            }
        }
    }

    /**
     * 顶部跳转
     */
    private void topForward() {
        XWPFParagraph paragraph = doc.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.RIGHT);
        XWPFHyperlinkRun hyperlinkRun = paragraph.createHyperlinkRun("#" + topId);
        hyperlinkRun.setFontFamily("Microsoft YaHei");
        hyperlinkRun.setFontSize(9);
        hyperlinkRun.setColor(Util.getColorString(new Color(165, 165, 165)));
        hyperlinkRun.setText("回到顶部");
    }
}
