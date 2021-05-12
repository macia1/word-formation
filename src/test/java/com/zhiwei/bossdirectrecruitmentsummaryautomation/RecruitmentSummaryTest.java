package com.zhiwei.bossdirectrecruitmentsummaryautomation;

import lombok.extern.log4j.Log4j2;
import org.junit.jupiter.api.Test;

import java.awt.*;
import java.io.File;
import java.io.IOException;

/**
 * @author aszswaz
 * @date 2021/5/11 16:42:36
 */
@Log4j2
class RecruitmentSummaryTest {

    @Test
    void start() throws IOException {
        String out = "source/Boss直聘汇总输出格式自动化/" + RecruitmentSummaryTest.class.getSimpleName() + ".docx";
        String source = "source/Boss直聘汇总输出格式自动化/Boss直聘负面舆情汇总（0506-0507）.xlsx";
        try {
            RecruitmentSummary summary = new RecruitmentSummary(
                    source, out
            );
            summary.start();
            // 调用系统的默认程序打开文件
            Desktop.getDesktop().open(new File(out));
        } catch (RecruitmentSummaryException e) {
            log.error(e.getMessage(), e);
            Desktop.getDesktop().open(new File(source));
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        }
    }
}