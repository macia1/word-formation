package com.zhiwei.word.wordversion;

import com.zhiwei.word.BaseWord;
import com.zhiwei.word.ExcelEntity;
import com.zhiwei.word.WordTemplateVersion;
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
public class WordTest {
    @Test
    public void sasWord() throws IOException {
        List<ExcelEntity> dataList = new ArrayList<>();

        ExcelEntity dataColumn = new ExcelEntity();
        dataColumn.setSerialNumber("1");
        dataColumn.setTitle("三星计划量产3纳米GAA尖端芯片，SK海力士推进EUV DRAM生产时程");
        dataColumn.setFullText("集微网消息（文/小山），据韩媒etnews报道，在“Tech Week 2020 LIVE”活动上，三星电子和SK海力士宣布了各自关于下一代半导体技术的发展战略。其中，三星电子计划量产业界首批采用3纳米环绕式栈极（gate-all-around，简称GAA）工艺制造的尖端芯片；SK海力士则正准备生产基于极紫外光刻 (EUV)技术的DRAM。\n" +
                "三星电子表示，该公司计划通过IBM和英伟达的下一代CPU和GPU的订单，在全球代工市场上利用GAA技术开拓下一代产品市场。\n" +
                "三星代工的执行董事Kang Moon-soo称，“我们计划大规模生产行业第一批基于GAA技术的半导体”。\n" +
                "韩媒指出，截至目前，三星电子和台积电是业界仅有的开始开发GAA工艺的两家公司。\n" +
                "这也意味着，如果三星电子能够在量产时程上超越台积电，三星将能够抓住机遇，甚至在全球代工市场上领先台积电。\n" +
                "而SK海力士也将很快量产基于EUV工艺的DRAM。SK海力士未来技术研究所负责人Lim Chang-moon表示，“我们计划从第4代10nm (1a) DRAM开始应用EUV工艺。同时我们计划明年初开始大规模生产”。\n" +
                "数据显示，全球DRAM市场约94%的份额分别由三星、SK海力士和美光垄断，两家韩国公司约占74%的市场份额。\n" +
                "因而，若GAA和EUV DRAM成功落地商业化，韩国半导体产业的地位将进一步提升。同时，报道指出，EUV DRAM将能进一步扩大韩国半导体企业与中国半导体企业之间的差距。鉴于EUV设备的昂贵成本以及有限的供应量，韩国企业将更有别于中国企业。");
        dataColumn.setKoreanTitle("삼성 3나노 GAA 첨단 칩 양산 계획, SK하이닉스 EUV로 D램 생산 일정에 박차");
        dataColumn.setKoreanAbstract("한국 언론 etnews 보도에 의하면, 삼성전자와 SK하이닉스는 '테크위크(tech week) 2020 LIVE'에서 각자 차세대 반도체 기술 발전 전략을 밝혔다고 한다. 삼성전자는 업계 최초로 3나노 게이트 올 어라운드(gate-all-around, GAA로 약칭) 공정을 적용한 첨단 칩을 양산할 계획이다. 삼성전자는 IBM과 엔비디아의 차세대 CPU와 GPU 수주를 통해 글로벌 파운드리 시장에서 GAA 기술로 차세대 시장을 개척할 계획이라고 밝혔다.");
        dataColumn.setChannel("集微网");
        dataColumn.setChannelDomainName("laoyaoba.com");
        dataColumn.setTime("2020/9/18");
        dataColumn.setChineseBrandClassification("三星关联");
        dataColumn.setKoreanBrandClassification("삼성");
        dataList.add(dataColumn);

        //传入数据
        BaseWord.start(
                //需要生成的模板版本，为枚举类：com.zhiwei.word.WordTemplateVersion
                //目前支持SAS和SCS
                WordTemplateVersion.SAS,
                //excel数据
                dataList,
                //输出流
                new FileOutputStream("testWrite/write-sas.docx")
        );
    }

    @Test
    public void scsWord() throws IOException {
        List<ExcelEntity> dataList = new ArrayList<>();

        ExcelEntity dataColumn = new ExcelEntity();
        dataColumn.setSerialNumber("1");
        dataColumn.setTitle("三星计划量产3纳米GAA尖端芯片，SK海力士推进EUV DRAM生产时程");
        dataColumn.setFullText("集微网消息（文/小山），据韩媒etnews报道，在“Tech Week 2020 LIVE”活动上，三星电子和SK海力士宣布了各自关于下一代半导体技术的发展战略。其中，三星电子计划量产业界首批采用3纳米环绕式栈极（gate-all-around，简称GAA）工艺制造的尖端芯片；SK海力士则正准备生产基于极紫外光刻 (EUV)技术的DRAM。\n" +
                "三星电子表示，该公司计划通过IBM和英伟达的下一代CPU和GPU的订单，在全球代工市场上利用GAA技术开拓下一代产品市场。\n" +
                "三星代工的执行董事Kang Moon-soo称，“我们计划大规模生产行业第一批基于GAA技术的半导体”。\n" +
                "韩媒指出，截至目前，三星电子和台积电是业界仅有的开始开发GAA工艺的两家公司。\n" +
                "这也意味着，如果三星电子能够在量产时程上超越台积电，三星将能够抓住机遇，甚至在全球代工市场上领先台积电。\n" +
                "而SK海力士也将很快量产基于EUV工艺的DRAM。SK海力士未来技术研究所负责人Lim Chang-moon表示，“我们计划从第4代10nm (1a) DRAM开始应用EUV工艺。同时我们计划明年初开始大规模生产”。\n" +
                "数据显示，全球DRAM市场约94%的份额分别由三星、SK海力士和美光垄断，两家韩国公司约占74%的市场份额。\n" +
                "因而，若GAA和EUV DRAM成功落地商业化，韩国半导体产业的地位将进一步提升。同时，报道指出，EUV DRAM将能进一步扩大韩国半导体企业与中国半导体企业之间的差距。鉴于EUV设备的昂贵成本以及有限的供应量，韩国企业将更有别于中国企业。");
        dataColumn.setKoreanTitle("삼성 3나노 GAA 첨단 칩 양산 계획, SK하이닉스 EUV로 D램 생산 일정에 박차");
        dataColumn.setKoreanAbstract("한국 언론 etnews 보도에 의하면, 삼성전자와 SK하이닉스는 '테크위크(tech week) 2020 LIVE'에서 각자 차세대 반도체 기술 발전 전략을 밝혔다고 한다. 삼성전자는 업계 최초로 3나노 게이트 올 어라운드(gate-all-around, GAA로 약칭) 공정을 적용한 첨단 칩을 양산할 계획이다. 삼성전자는 IBM과 엔비디아의 차세대 CPU와 GPU 수주를 통해 글로벌 파운드리 시장에서 GAA 기술로 차세대 시장을 개척할 계획이라고 밝혔다.");
        dataColumn.setChannel("集微网");
        dataColumn.setChannel("laoyaoba.com");
        dataColumn.setTime("2020/9/18");
        dataColumn.setChineseBrandClassification("三星关联");
        dataColumn.setKoreanBrandClassification("삼성");
        dataList.add(dataColumn);

        //传入数据
        BaseWord.start(
                //需要生成的模板版本，为枚举类：com.zhiwei.word.WordTemplateVersion
                //目前支持SAS和SCS
                WordTemplateVersion.SCS,
                //excel数据
                dataList,
                //输出流
                new FileOutputStream("testWrite/write-sas.docx")
        );
    }
}
