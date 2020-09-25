package util;

import com.zhiwei.util.Util;
import org.junit.Test;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.io.RandomAccessFile;

/**
 * @author 朽木不可雕也
 * @date 2020/9/17 12:01
 * @week 星期四
 */
public class UtilTest {
    @Test
    public void getImageRgb() {
        int[] rgb = new int[3];

        File file = new File("src/main/resources/sas/tableTitle.jpg");
        BufferedImage bi = null;
        try {
            bi = ImageIO.read(file);
        } catch (Exception e) {
            e.printStackTrace();
        }
        int width = bi.getWidth();
        int height = bi.getHeight();
        int minx = bi.getMinX();
        int miny = bi.getMinY();
        System.out.println("width=" + width + ",height=" + height + ".");
        System.out.println("minx=" + minx + ",miniy=" + miny + ".");
        for (int i = minx; i < width; i++) {
            for (int j = miny; j < height; j++) {
                int pixel = bi.getRGB(i, j);
                rgb[0] = (pixel & 0xff0000) >> 16;
                rgb[1] = (pixel & 0xff00) >> 8;
                rgb[2] = (pixel & 0xff);
                System.out.println("i=" + i + ",j=" + j + ":(" + rgb[0] + "," + rgb[1] + "," + rgb[2] + ")");
            }
        }
    }

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

    @Test
    public void setImageSize() {
        try {
            BufferedImage image = Util.setImageSize("E:\\图片\\来自QQ\\-5aeb2eca88659f4e.png", 100, 100);
            ImageIO.write(image, "PNG", new File("test.png"));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void vbEquals() {
        try {
            RandomAccessFile accessFile1 = new RandomAccessFile("test1.txt", "r");
            RandomAccessFile accessFile2 = new RandomAccessFile("test2.txt", "r");
            while (true) {
                int i1 = accessFile1.read();
                int i2 = accessFile2.read();
                if (i1 == -1 || i2 == -1) {
                    break;
                }
                if (i1 != i2) {
                    System.out.println(false);
                    return;
                }
                System.out.print((char) i1);
            }
            System.out.println(true);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
