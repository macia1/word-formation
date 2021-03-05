package com.zhiwei.bossdirecthireautomation.wordutil;

import com.zhiwei.bossdirecthireautomation.tree.EventExcelEntity;
import com.zhiwei.bossdirecthireautomation.tree.PolyTreeNode;
import com.zhiwei.bossdirecthireautomation.tree.PolyTreeUtil;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * word生成
 *
 * @author fuck-aszswaz
 * @date 2021-02-20 16:32:54
 */
@Log4j2
public class WorldUtil extends XWPFDocument {
    /**
     * 多叉树
     */
    private final PolyTreeUtil polyTreeUtil;
    /**
     * 顶级排序
     */
    private final AtomicInteger topRanking = new AtomicInteger(1);

    public WorldUtil(PolyTreeUtil polyTreeUtil) {
        super();
        // 标题
        XWPFParagraph titleParagraph = super.createParagraph();// 创建段落
        XWPFRun titleRun = titleParagraph.createRun();// 创建文本操作对象
        titleRun.setText("Boss直聘品牌&竞品舆情周报");
        titleRun.setFontFamily("微软雅黑");
        titleRun.setFontSize(ContentFont.三号.getPoundValue());
        titleRun.setBold(true);
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);
        titleParagraph.setSpacingBefore(12 * 10 * 2);// 设置段前间距12磅
        titleParagraph.setSpacingAfter(3 * 10 * 2);// 设置段后间距3磅
        this.polyTreeUtil = polyTreeUtil;
    }

    /**
     * world内容生成
     */
    public void generate() {
        /*
        生成以下七大部分
         1. Boss直聘正面事件
         2. Boss直聘负面+敏感事件
         3. Boss直聘个人投诉类舆情分析
         4. 竞品正面事件
         5. 竞品负面事件
         6. 竞品个人投诉类舆情分析
         7. 行业负面事件
         */
        final PolyTreeNode rootNode = this.polyTreeUtil.getRootNode();
        PolyTreeNode labelNode = rootNode.getChildNode("BOSS直聘");
        this.bossPositive(labelNode);// Boss直聘正面
        this.bossNegative(labelNode);// Boss负面+敏感
        labelNode = rootNode.getChildNode("竞品");
        this.otherBrand("正面", labelNode);// 竞品正面
        this.otherBrand("负面", labelNode);// 竞品负面
        labelNode = rootNode.getChildNode("行业");
        this.industry(labelNode);// 行业
    }

    /**
     * Boss直聘正面事件
     */
    private void bossPositive(PolyTreeNode labelNode) {
        // 获取Boss直聘节点
        final PolyTreeNode bossBrandNode = labelNode.getChildNode("Boss直聘");
        // 获得正面情感节点
        final PolyTreeNode positiveNode = bossBrandNode.getChildNode("正面");
        final XWPFParagraph mainParagraph = super.createParagraph();
        this.writeTopTitle("Boss直聘正面事件", mainParagraph);// 生成标题

        mainParagraph.createRun().addCarriageReturn();

        this.event(positiveNode, mainParagraph);// 处理事件
    }

    /**
     * Boss直聘负面+敏感
     */
    private void bossNegative(PolyTreeNode labelNode) {
        // 获得Boss直聘节点
        final PolyTreeNode bossBrand = labelNode.getChildNode("Boss直聘");
        // 负面情感节点
        final PolyTreeNode negative = bossBrand.getChildNode("负面+敏感");
        // 创建专属段落
        final XWPFParagraph paragraph = super.createParagraph();
        // 标题
        this.writeTopTitle("品牌负面+敏感事件", paragraph);

        paragraph.createRun().addCarriageReturn();

        this.event(negative, paragraph);
    }

    /**
     * 行业部分
     */
    private void industry(PolyTreeNode labelNode) {
        // 行业部分按照事件传播量排序
        final List<PolyTreeNode> events = new ArrayList<>();
        labelNode.getNodes().forEach(brand -> brand.getNodes().forEach(emotion -> events.addAll(emotion.getNodes())));
        if (events.isEmpty()) return;
        // 按照传播量降序
        events.sort(Collections.reverseOrder(Comparator.comparingInt(PolyTreeNode::getDataNodeSize)));

        final XWPFParagraph paragraph = super.createParagraph();
        this.writeTopTitle("行业负面事件", paragraph);
        paragraph.createRun().addCarriageReturn();

        // 输出所有事件
        AtomicInteger sort = new AtomicInteger(1);// 排序
        events.forEach(event -> {
            XWPFRun run = paragraph.createRun();
            run.setText(sort.getAndIncrement() + "、");// 排序
            run.setFontFamily(BaseConfig.DEFAULT_FONT);
            run.setFontSize(BaseConfig.DEFAULT_FONT_SIZE);
            run.setBold(true);
            List<EventExcelEntity> selects = new ArrayList<>();// 候选
            for (PolyTreeNode node : event.getNodes()) {
                List<EventExcelEntity> conversions = new ArrayList<>(node.getNodes().size());
                node.getNodes().forEach(conversion -> conversions.add((EventExcelEntity) conversion));// 转换类型

                selects.add(this.selectByInfluence(conversions));// 选出稿件的第一权重，加入时间候选列表
            }
            // 从候选列表获得权重最高的
            EventExcelEntity eventExcelEntity = this.selectByInfluence(selects);
            // 输出
            this.createLink(event.getNodeName(), eventExcelEntity.getUrl(), paragraph);
            run = this.writeDefaultString("（传播量：" + event.getDataNodeSize() + "）", paragraph);// 输出传播量
            run.setBold(true);
            run.addCarriageReturn();

            // 聚合事件下的所有数据
            final PolyTreeNode countNode = new PolyTreeNode();
            event.getNodes().forEach(manuscript -> manuscript.getNodes().forEach(countNode::add));
            // 从数据中选择n个渠道
            List<EventExcelEntity> channels = this.chooseChannel(countNode);
            Iterator<EventExcelEntity> iterator = channels.iterator();
            StringBuilder builder = new StringBuilder("高频/重点参与渠道：");
            while (iterator.hasNext()) {
                builder.append(iterator.next().getChannel());
                if (iterator.hasNext()) builder.append("、");
            }
            if (event.getDataNodeSize() > channels.size()) builder.append("等");
            this.writeDefaultString(builder.toString(), paragraph).addCarriageReturn();
        });
    }

    /**
     * 竞品
     */
    private void otherBrand(String emotion, PolyTreeNode labelNode) {
        // 输出其他品牌的正面情感消息
        // 除Boss直聘以外的其他品牌
        final List<PolyTreeNode> brands = labelNode.getNodes();
        final XWPFParagraph paragraph = super.createParagraph();
        this.writeTopTitle("竞品" + emotion + "事件", paragraph);
        paragraph.createRun().addCarriageReturn();

        // 输出品牌
        NumberUtil numberUtil = new NumberUtil();
        int index = 0;
        for (PolyTreeNode brand : brands) {
            // 检查是否存在符合输出条件的事件
            PolyTreeNode emotionNode = brand.getChildNode(emotion);
            if (!this.checkBrandEmotion(emotionNode)) continue;

            {
                // 输出品牌名称
                numberUtil.setValue(++index);
                XWPFRun run = this.writeDefaultString("（" + numberUtil.toChineseNumber() + "）" + brand.getNodeName(), paragraph);
                run.setBold(true);
                run.addCarriageReturn();
            }
            this.event(emotionNode, paragraph);
        }
    }

    /**
     * 检查品牌的某一情感符合条件的事件是否为空
     *
     * @param brandNode 品牌节点
     */
    private boolean checkBrandEmotion(PolyTreeNode brandNode) {
        if (Objects.isNull(brandNode)) return false;
        List<PolyTreeNode> events = brandNode.getNodes();
        for (PolyTreeNode event : events) {
            if (this.checkEvent(event)) return true;
        }
        return false;
    }

    /**
     * 输出顶级标题
     */
    @SuppressWarnings("SameParameterValue")
    private void writeTopTitle(String title, XWPFParagraph paragraph) {
        NumberUtil numberUtil = new NumberUtil(this.topRanking.getAndIncrement());
        String chineseNumber = numberUtil.toChineseNumber();// 获得对应的中文数字
        title = chineseNumber + "、" + title;

        XWPFRun run = paragraph.createRun();
        run.setText(title);
        run.setBold(true);
        run.setFontFamily(BaseConfig.DEFAULT_FONT);
        run.setFontSize(BaseConfig.DEFAULT_FONT_SIZE);
        run.addCarriageReturn();// 换行
    }

    /**
     * 事件输出
     *
     * @param emotionNode 情感节点
     * @param paragraph   一个主要部分的段落
     */
    private void event(PolyTreeNode emotionNode, XWPFParagraph paragraph) {
        final List<PolyTreeNode> events = emotionNode.getNodes();
        int serialNumber = 0;// 事件序号
        for (PolyTreeNode event : events) {

            // 所有稿件的传播量小于5，不输出
            if (!this.checkEvent(event)) continue;

            String title = ++serialNumber + "、" + event.getNodeName() + "(传播量: " + event.getDataNodeSize() + ")";
            XWPFRun eventRun = paragraph.createRun();
            eventRun.setText(title);
            eventRun.setFontFamily(BaseConfig.DEFAULT_FONT);
            eventRun.setFontSize(BaseConfig.DEFAULT_FONT_SIZE);
            eventRun.addCarriageReturn();
            eventRun.setBold(true);

            this.manuscript(event, paragraph);// 处理稿件
        }
    }

    /**
     * 检查事件是否符合输出条件
     */
    private boolean checkEvent(PolyTreeNode event) {
        for (PolyTreeNode node : event.getNodes()) {
            if (node.getDataNodeSize() >= 5) {
                return true;
            }
        }
        return false;
    }

    /**
     * 稿件处理
     *
     * @param eventNode 事件节点
     * @param paragraph 段落
     */
    private void manuscript(PolyTreeNode eventNode, XWPFParagraph paragraph) {
        List<PolyTreeNode> manuscripts = this.selectManuscript(eventNode);// 选择需要输出到world的稿件
        int index = 1;
        for (PolyTreeNode manuscript : manuscripts) {
            if (manuscript.getNodeName().equals("春华资本联合管理层战略投资智联招聘 深耕中国人力资源服务市场")) {
                new Object();
            }
            /*
             1.输出稿件标题
             2.输出重点媒体
             3.输出原发、权重媒体、高频参与渠道, 原发1个，权重列3个，高频列3个
             */
            EventExcelEntity aData = this.selectData(manuscript);// 选择一条数据作为稿件的展示数据
            String title;
            if ("原发".equals(aData.getWhether())) {
                title = index++ + ")【原发：" + aData.getChannel() + "】";
            } else {
                title = index++ + ")【原发：未知】";
            }
            this.writeDefaultString(title, paragraph);
            // 写出超链接
            this.createLink(manuscript.getNodeName(), aData.getUrl(), paragraph);
            // 写出传播量
            title = "(传播量：" + manuscript.getDataNodeSize() + ")";
            this.writeDefaultString(title, paragraph).addCarriageReturn();// 输出并换行
            // 原发、权重媒体
            List<EventExcelEntity> channels = this.chooseChannel(manuscript);

            StringBuilder builder = new StringBuilder("高频/重点参与渠道：");
            Iterator<EventExcelEntity> iterator = channels.iterator();
            int size = 0;
            while (iterator.hasNext()) {
                builder.append(iterator.next().getChannel());
                if (iterator.hasNext() && ++size < 5) {
                    builder.append("、");
                } else {
                    break;
                }
            }
            if (manuscript.getDataNodeSize() > channels.size()) builder.append("等");
            this.writeDefaultString(builder.toString(), paragraph).addCarriageReturn();
        }
    }

    /**
     * 选择需要被列举的渠道
     *
     * @param manuscript 稿件节点
     */
    List<EventExcelEntity> chooseChannel(PolyTreeNode manuscript) {
        /*
        选举需要被列举的渠道
         原发1个，权重列3个，高频列3个
         */
        DataNodeList dataNodeList = new DataNodeList();

        manuscript.getNodes().forEach(node -> dataNodeList.add((EventExcelEntity) node));

        dataNodeList.sort(Collections.reverseOrder(Comparator.comparingDouble(EventExcelEntity::getInfluence)));// 按照影响力降序

        DataNodeList results = new DataNodeList(); // 结果集
        // 获得原发
        for (int i = 0; i < dataNodeList.size(); i++) {
            if ("原发".equals(dataNodeList.get(i).getWhether())) {
                results.add(dataNodeList.remove(i));
                break;
            }
        }
        // 获得三个权重（影响力最高的三个渠道）
        List<EventExcelEntity> subList = dataNodeList.subList(0, Math.min(dataNodeList.size(), 3));
        results.addAll(subList);
        dataNodeList.removeAll(subList);

        // 按照渠道的出现频次排序
        Map<String, Integer> countMap = new HashMap<>();// 统计容器
        dataNodeList.forEach(dataNode -> {
            Integer count = countMap.getOrDefault(dataNode.getChannel(), 0);// 获取统计记录
            countMap.put(dataNode.getChannel(), ++count);// 更新统计结果
        });

        // 根据统计结果排序
        List<EventExcelEntity> sortList = new ArrayList<>();
        dataNodeList.forEach(dataNode -> {
            if (sortList.isEmpty()) {
                sortList.add(dataNode);
                return;
            }

            int count01 = countMap.get(dataNode.getChannel());
            // 排序
            boolean isAdd = true;// 是否添加
            for (int i = 0; i < sortList.size(); i++) {
                int count02 = countMap.get(sortList.get(i).getChannel());
                if (count01 > count02) {
                    sortList.add(i, dataNode);
                    isAdd = false;
                    break;
                }
            }
            if (isAdd) sortList.add(dataNode);
        });
        subList = sortList.subList(0, Math.min(sortList.size(), 3));
        results.addAll(subList);
        dataNodeList.removeAll(subList);

        return results;
    }

    /**
     * 选择需要输出到world的稿件
     *
     * @param eventNode 事件节点
     */
    private List<PolyTreeNode> selectManuscript(PolyTreeNode eventNode) {
        List<PolyTreeNode> manuscripts = eventNode.getNodes();// 稿件节点
        List<PolyTreeNode> result = new ArrayList<>();
        for (PolyTreeNode manuscript : manuscripts) {
            if (manuscript.getDataNodeSize() >= 5) {
                result.add(manuscript);
            }
        }

        if (result.isEmpty()) result.addAll(manuscripts);// 如果不存在传播量>=5的稿件，则输出全部稿件
        return result;
    }

    /**
     * 从稿件节点中，选择数据
     */
    private EventExcelEntity selectData(PolyTreeNode manuscriptNode) {
        /*
        从稿件节点中选择数据
         1.原发优先
         2.如果存在多条原发，网媒、微信优先
         3.如果没有原发，按照影响力最高的优先
         */
        final List<PolyTreeNode> multipleData = manuscriptNode.getNodes();
        // 对原发进行筛选
        List<EventExcelEntity> results = new ArrayList<>();
        for (PolyTreeNode node : multipleData) {
            EventExcelEntity data = (EventExcelEntity) node;
            if ("原发".equals(data.getWhether())) {
                results.add(data);
            }
        }
        // 只有一条原发
        if (results.size() == 1) {
            return results.get(0);
        } else if (results.size() > 1) {
            // 存在多条原发， 删除渠道包含“原创列表的”的数据
            results.removeIf(eventExcelEntity -> eventExcelEntity.getChannel().contains("原创列表"));
            // 根据影响力选择一条
            return this.selectByInfluence(results);
        } else {
            // 不存在原发，从转载当中选择一条
            List<EventExcelEntity> data = new ArrayList<>();
            multipleData.forEach(polyTreeNode -> data.add((EventExcelEntity) polyTreeNode));
            return this.selectByInfluence(data);
        }
    }

    /**
     * 根据影响力选择数据
     */
    private EventExcelEntity selectByInfluence(List<EventExcelEntity> data) {
        /*
        根据影响力选择数据
         1.以网媒、微信优先
         2.根据影响力进行排名
         3.根据影响力排名结果，选择影响力最大的一条
         */
        // 判断来源
        for (EventExcelEntity datum : data) {
            if ("网媒".equals(datum.getSource())) {
                return datum;
            }
        }
        for (EventExcelEntity datum : data) {
            if ("微信".equals(datum.getSource())) return datum;
        }
        // 根据影响力进行排名
        for (int i = 0; i < data.size(); i++) {
            double currentInfluence = data.get(i).getInfluence();
            for (int i1 = 0; i1 < data.size(); i1++) {
                EventExcelEntity aData = data.get(i1);
                double aInfluence = aData.getInfluence();
                if (aInfluence > currentInfluence) {// 降序排序
                    data.set(i1, data.set(i, aData));// 交换位置
                }
            }
        }
        return data.get(0);
    }

    /**
     * 写出默认格式的文本
     *
     * @param text 文本
     */
    private XWPFRun writeDefaultString(String text, XWPFParagraph paragraph) {
        XWPFRun run = paragraph.createRun();
        run.setText(text);
        run.setFontFamily(BaseConfig.DEFAULT_FONT);
        run.setFontSize(BaseConfig.DEFAULT_FONT_SIZE);
        return run;
    }

    /**
     * 创建超链接
     */
    private void createLink(String text, String url, XWPFParagraph paragraph) {
        CTHyperlink hyperlink = paragraph.getCTP().addNewHyperlink();
        PackageRelationship relationship = super.getPackagePart().addExternalRelationship(url, XWPFRelation.HYPERLINK.getRelation());// 添加外部链接映射
        hyperlink.setId(relationship.getId());// 设置外部映射ID
        hyperlink.addNewR();
        XWPFHyperlinkRun hyperlinkRun = new XWPFHyperlinkRun(hyperlink, hyperlink.getRArray(0), paragraph);
        hyperlinkRun.setFontSize(BaseConfig.DEFAULT_FONT_SIZE);
        hyperlinkRun.setFontFamily(BaseConfig.DEFAULT_FONT);
        hyperlinkRun.setColor(this.getColorString(5, 99, 193));
        hyperlinkRun.setUnderline(UnderlinePatterns.SINGLE);// 添加下划线
        hyperlinkRun.setText(text);
    }

    /**
     * 获得颜色的十六进制
     */
    @SuppressWarnings({"SameParameterValue", "DuplicatedCode"})
    private String getColorString(int red, int green, int blue) {
        int[] numbers = {red, green, blue};
        StringBuilder builder = new StringBuilder();
        for (Integer number : numbers) {
            String numberStr = Integer.toHexString(number);
            //字符串长度不足2位，追加“0”
            if (numberStr.length() < 2) {
                builder.append("0");
            }
            builder.append(numberStr);
        }
        return builder.toString().toUpperCase();
    }

    /**
     * 写出word文档
     */
    public void write(String writePath) throws IOException {
        OutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(writePath);
            super.write(outputStream);
        } finally {
            try {
                if (Objects.nonNull(outputStream)) outputStream.close();
            } catch (IOException e) {
                log.error(e.getMessage(), e);
            }
        }
    }
}
