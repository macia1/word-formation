package com.zhiwei.bossdirectrecruitmentsummaryautomation;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.event.SyncReadListener;
import lombok.extern.log4j.Log4j2;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;

import static java.util.Objects.isNull;

/**
 * Boss直聘汇总自动化输出程序
 *
 * @author aszswaz
 * @date 2021/5/11 12:14:50
 */
@Log4j2
public class RecruitmentSummary {
    private final String dataSourceFile;
    /**
     * 文件的输出路径
     */
    private final String outputFile;
    /**
     * 文件的保存路径
     */
    private final XWPFDocument document;
    /**
     * 多叉树
     */
    private final Polytree polytree = new Polytree();

    public RecruitmentSummary(String dataSourceFile, String outputFile) {
        this.dataSourceFile = dataSourceFile;
        this.outputFile = outputFile;
        this.document = new XWPFDocument();
    }

    /**
     * 程序入口
     */
    public void start() throws RecruitmentSummaryException {
        try {
            List<ExcelReadEntity> readEntities = EasyExcel.read(this.dataSourceFile, ExcelReadEntity.class, new SyncReadListener()).sheet(0).doReadSync();
            readEntities = this.pretreatment(readEntities);
            readEntities.forEach(log::debug);
            this.writeToWord(readEntities);
        } catch (RecruitmentSummaryException e) {
            throw e;
        } catch (Exception e) {
            throw new RecruitmentSummaryException(e.getMessage(), e);
        }
    }

    /**
     * 数据预处理
     */
    @SuppressWarnings("CodeBlock2Expr")
    public List<ExcelReadEntity> pretreatment(List<ExcelReadEntity> readEntities) throws RecruitmentSummaryException {
        if (isNull(readEntities) || readEntities.isEmpty()) {
            throw new RecruitmentSummaryException("数据不能为空");
        }

        readEntities = readEntities.stream().filter(excelReadEntity -> {
            return StringUtils.isNotBlank(excelReadEntity.getMediaTage()) && !"/".equals(excelReadEntity.getMediaTage());
        }).collect(Collectors.toList());

        for (ExcelReadEntity readEntity : readEntities) {
            if (StringUtils.isBlank(readEntity.getSource())) {
                throw new RecruitmentSummaryException("来源不能为空");
            } else if (StringUtils.isBlank(readEntity.getChannel())) {
                throw new RecruitmentSummaryException("渠道不能为空");
            } else if (StringUtils.isBlank(readEntity.getTitle())) {
                throw new RecruitmentSummaryException("标题不能为空");
            } else if (StringUtils.isBlank(readEntity.getUrl())) {
                throw new RecruitmentSummaryException("地址不能为空");
            }
        }

        if (readEntities.isEmpty()) {
            throw new RecruitmentSummaryException("不存在媒介");
        }

        // 纠正错误的来源
        this.correctSource(readEntities);

        return readEntities;
    }

    /**
     * 纠正错误的来源
     */
    private void correctSource(List<ExcelReadEntity> readEntities) {
        readEntities.forEach(readEntity -> {
            String source = this.matchSourceByUrl(readEntity.getUrl());
            if (StringUtils.isNotBlank(source)) {
                readEntity.setSource(source);
            }
        });
    }

    /**
     * 根据url匹配平台
     */
    @SuppressWarnings("SpellCheckingInspection")
    private String matchSourceByUrl(String url) {
        if (url.contains("weibo.com") || url.contains("weibo.cn")) {
            return "微博";
        } else if (url.contains("weixin.qq.com")) {
            return "微信公众号";
        } else if (url.contains("www.toutiao.com") || url.contains("m.toutiao.com")) {
            return "今日头条";
        } else if (url.contains("zhihu.com")) {
            return "知乎专栏";
        } else if (url.contains("36kr.com")) {
            return "36kr";
        } else if (url.contains("sike.news.cn")) {
            return "思客";
        } else if (url.contains("jianshu.com")) {
            return "简书";
        } else if (url.contains("huxiu.com")) {
            return "虎嗅";
        } else if (url.contains("douban.com")) {
            return "豆瓣";
        } else if (url.contains("wap.peopleapp.com/article/rmh") || url.contains("rmh.pdnews.cn")) {
            return "人民号";
        } else if (url.contains("new.qq.com") || url.contains("xw.qq.com") || url.contains("page.om.qq")
                || url.contains("kuaibao.qq.com")) {
            return "企鹅号";
        } else if (url.contains("feng.ifeng.com") || url.contains("ishare.ifeng.com")
                || url.contains("fashion.ifeng.com")) {
            return "大风号";
        } else if (url.contains("a.mp.uc.cn") || url.contains("mparticle.uc.cn") || url.contains("m.uczzd.cn")
                || url.contains("iflow.uc.cn/webview")) {
            return "大鱼号";
        } else if (url.contains("xiaohongshu.com")) {
            return "小红书";
        } else if (url.contains("360kuai.com")) {
            return "快咨询";
        } else if (url.contains("dcdapp.com") || url.contains("dcd.zjbyte.cn")) {
            return "懂车帝";
        } else if (url.contains("sohu.com")) {
            return "搜狐号";
        } else if (url.contains("thepaper.cn/newsDetail_forward")) {
            return "澎湃号";
        } else if (url.contains("baijiahao.baidu.com") || url.contains("mbd.baidu.com")
                || url.contains("cpu.baidu.com")) {
            return "百家号";
        } else if (url.contains("dy.163.com") || url.contains("3g.163.com/dy/") || url.contains("c.m.163.com")) {
            return "网易号";
        } else if (url.contains("tmtpost.com")) {
            return "钛媒体";
        } else if (url.contains("yidianzixun.com")) {
            return "一点咨询";
        } else if (url.contains("weibo.com/ttarticle") || url.contains("weibo.com/article/")) {
            return "新浪专栏";
        } else if (url.contains("jiemian.com")) {
            return "界面新闻";
        } else if (url.contains("xueqiu.com")) {
            return "雪球";
        } else if (url.contains("3w.huanqiu.com")) {
            return "环球号";
        } else if (url.contains("bilibili.com")) {
            return "bilibili";
        } else if (url.contains("iesdouyin.com") || url.contains("aweme.snssdk.com")) {
            return "抖音";
        } else if (url.contains("haokan.baidu.com")) {
            return "好看视频";
        } else if (url.contains("kuaishou.com") || url.contains("m.gifshow.com")) {
            return "快手";
        } else if (url.contains("tieba.baidu.com") || url.contains("bbs.")) {
            return "贴吧论坛";
        }
        return null;
    }

    /**
     * 输出数据到word文件
     */
    private void writeToWord(List<ExcelReadEntity> readEntities) throws IOException {
        for (ExcelReadEntity readEntity : readEntities) {
            this.polytree.add(readEntity);
        }

        final PolytreeNode rootNode = this.polytree.getRoot();
        this.summary(rootNode);
        this.getParagraph();
        this.mediaTage(rootNode);
        this.document.write(new FileOutputStream(this.outputFile));
    }

    /**
     * 内容摘要部分
     */
    private void summary(PolytreeNode rootNode) {
        final List<PolytreeNode> mediaTagNodes = rootNode.getPolytreeNodes();

        String[] lines = new String[2 + mediaTagNodes.size()];
        int index = 0;
        lines[index++] = "重点媒体相关舆情数量：" + mediaTagNodes.size() + "篇";
        lines[index++] = "涉及的媒介及媒体：";

        for (PolytreeNode mediaTagNode : mediaTagNodes) {
            if (!mediaTagNode.getPolytreeNodes().isEmpty()) {
                lines[index++] = mediaTagNode.getNodeName() + ": " + mediaTagNode.getPolytreeNodes().get(0).getNodeName();
            } else {
                lines[index++] = mediaTagNode.getNodeName() + ": ";
            }
        }

        for (String line : lines) {
            XWPFRun xwpfRun = this.getRun(this.getParagraph());
            xwpfRun.setText(line);
        }
    }

    /**
     * 媒介
     */
    private void mediaTage(PolytreeNode rootNode) {
        // 媒介
        final Iterator<PolytreeNode> mediaTageIterator = rootNode.getPolytreeNodes().iterator();
        while (mediaTageIterator.hasNext()) {
            final PolytreeNode mediaTageNode = mediaTageIterator.next();
            XWPFRun xwpfRun = this.getRun(this.getParagraph());
            xwpfRun.setText("@" + mediaTageNode.getNodeName());

            this.mediaType(mediaTageNode);

            if (mediaTageIterator.hasNext()) {
                this.document.createParagraph();
            }
        }
    }

    /**
     * 媒体名称
     */
    private void mediaType(PolytreeNode mediaTageNode) {
        XWPFRun xwpfRun;
        // 媒体名称
        final List<PolytreeNode> mediaTypeNodes = mediaTageNode.getPolytreeNodes();
        final Iterator<PolytreeNode> iterator = mediaTypeNodes.iterator();
        while (iterator.hasNext()) {
            PolytreeNode mediaTypeNode = iterator.next();
            xwpfRun = this.getRun(this.getParagraph());
            xwpfRun.setText("媒体名称： " + mediaTypeNode.getNodeName());

            xwpfRun = this.getRun(this.getParagraph());
            xwpfRun.setText("相关内容链接如下：");

            this.sourceNode(mediaTypeNode);

            if (iterator.hasNext()) {
                this.getParagraph();
            }
        }
    }

    /**
     * 来源
     */
    private void sourceNode(PolytreeNode mediaTypeNode) {
        XWPFRun xwpfRun;
        // 来源
        int serialNumber = 0;
        final List<PolytreeNode> sourceNodes = mediaTypeNode.getPolytreeNodes();
        final Iterator<PolytreeNode> sourceIterator = sourceNodes.iterator();
        while (sourceIterator.hasNext()) {
            final PolytreeNode sourceNode = sourceIterator.next();
            xwpfRun = this.getRun(this.getParagraph());
            xwpfRun.setText(++serialNumber + ".【" + sourceNode.getNodeName() + "】");

            this.dataNode(sourceNode);

            if (sourceIterator.hasNext()) {
                this.document.createParagraph();
            }
        }
    }

    /**
     * 数据
     */
    private void dataNode(PolytreeNode sourceNode) {
        XWPFRun xwpfRun;
        // 数据节点
        final List<PolytreeNode> dataNodes = sourceNode.getPolytreeNodes();
        final Iterator<PolytreeNode> dataIterator = dataNodes.iterator();
        while (dataIterator.hasNext()) {
            ExcelReadEntity readEntity = (ExcelReadEntity) dataIterator.next();
            xwpfRun = this.getRun(this.getParagraph());
            xwpfRun.setText(readEntity.getTitle());

            xwpfRun = this.document.createParagraph().createHyperlinkRun(readEntity.getUrl());
            xwpfRun.setText(readEntity.getUrl());
            xwpfRun.setColor("0000FF");
            xwpfRun.setUnderline(UnderlinePatterns.SINGLE);
            if (dataIterator.hasNext()) {
                this.document.createParagraph();
            }
        }
    }

    private XWPFRun getRun(XWPFParagraph paragraph) {
        XWPFRun xwpfRun = paragraph.createRun();
        xwpfRun.setFontFamily("微软雅黑");
        xwpfRun.setFontSize(10);
        return xwpfRun;
    }

    private XWPFParagraph getParagraph() {
        return this.document.createParagraph();
    }
}
