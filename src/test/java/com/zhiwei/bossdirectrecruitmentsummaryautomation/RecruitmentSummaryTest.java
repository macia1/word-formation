package com.zhiwei.bossdirectrecruitmentsummaryautomation;

import lombok.extern.log4j.Log4j2;
import org.junit.jupiter.api.Test;

import java.io.IOException;

/**
 * @author aszswaz
 * @date 2021/5/11 16:42:36
 */
@SuppressWarnings("TryWithIdenticalCatches")
@Log4j2
class RecruitmentSummaryTest {

    @Test
    void start() {
        String out = "target/" + RecruitmentSummaryTest.class.getSimpleName() + "-04.docx";
        String source = "source/Boss直聘汇总输出格式自动化/222222.xlsx";
        try {
            RecruitmentSummary summary = new RecruitmentSummary(
                    source, out
            );
            summary.start();
            // 调用系统的默认程序打开文件
            // Desktop.getDesktop().open(new File(out));
        } catch (RecruitmentSummaryException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void demo() {
        System.out.println((int) ' ');
    }
}