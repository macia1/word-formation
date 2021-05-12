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

        return readEntities;
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
        for (PolytreeNode mediaTypeNode : mediaTypeNodes) {
            xwpfRun = this.getRun(this.getParagraph());
            xwpfRun.setText("媒体名称： " + mediaTypeNode.getNodeName());

            xwpfRun = this.getRun(this.getParagraph());
            xwpfRun.setText("相关内容链接如下：");

            this.sourceNode(mediaTypeNode);
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
