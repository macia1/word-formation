package com.zhiwei;

import com.zhiwei.bossdirectrecruitmentsummaryautomation.RecruitmentSummaryException;
import com.zhiwei.util.UrlUtil;
import org.junit.jupiter.api.Test;

import java.io.UnsupportedEncodingException;

/**
 * @author aszswaz
 * @date 2021/5/25 17:22:10
 * @IDE IntelliJ IDEA
 */
@SuppressWarnings("JavaDoc")
class UrlUtilTest {

    @Test
    void escapeUrl() throws UnsupportedEncodingException, RecruitmentSummaryException {
        @SuppressWarnings("SpellCheckingInspection")
        String url = "https://weixin.sogou.com/link?url=dn9a_\"-gY295K0Rci_xozVXfdMkSQTLW6cwJThYulHEtVjXrGTiVgS7BMwA5SXGqN3xnrimnD14BKY7z0S8wFCVqXa8Fplpd9hpApiVwUXKN6LnkT80nb3yvCNHjH5rKEvHDO29Eh49fie1kWhTbW_GMzwGW8XnGxaqVoHnAhgVG3uwPws7w0Ls2imaxMiGqAEeB05tBpDqVftczFRVSP85SGmlQCvoYNy4oHumClRB6OqGBPp4AZF7Vjx9pqreSNVnRvGqAVb35Ae0f8bRARvQ..&type=2&query=%E6%AF%9B%E8%80%80%E6%A3%AE %E8%BA%AB%E4%BB%BD%E8%B0%83%E6%9F%A5 %E4%B8%96%E5%A5%A2%E4%BC%9A&token=112D7520291525B0C2C60113254F8D32C35CD76760A74EA1";
        System.out.println(UrlUtil.escapeUrl(url));
    }
}