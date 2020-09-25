package util;

import com.zhiwei.util.Util;
import org.junit.Test;

/**
 * @author 朽木不可雕也
 * @date 2020/9/17 12:01
 * @week 星期四
 */
public class UtilTest {
    @Test
    public void translationTest() {
        try {
            String text = "If you want to add custom margins to your document. use this code.";
            /*Map<String, String> textMap = Util.translation(text, Util.Language.English_Chinese);
            textMap.forEach((content, translate) -> {
                System.out.println(content + ": " + translate);
            });*/
            System.out.println(Util.translationToString(text, Util.Language.English_Chinese));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
