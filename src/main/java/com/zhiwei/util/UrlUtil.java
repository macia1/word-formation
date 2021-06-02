package com.zhiwei.util;

import com.zhiwei.bossdirectrecruitmentsummaryautomation.RecruitmentSummaryException;

import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

/**
 * url转换
 *
 * @author aszswaz
 * @date 2021/5/25 15:59:37
 * @IDE IntelliJ IDEA
 */
@SuppressWarnings("JavaDoc")
public class UrlUtil {
    /**
     * 数字的unicode范围
     */
    private static final int[] NUMBERS = {0x30, 0x39};
    /**
     * 小写字母unicode范围
     */
    private static final int[] LOWER_CASE_LETTERS = {0x61, 0x7a};
    /**
     * 大写字母unicode范围
     */
    private static final int[] UPPERCASE_LETTER = {0x41, 0x5a};
    /**
     * HTTP中的保留字符，本程序不需要转义的字符
     */
    private static final char[] HTTP_CHARS = {
            '_', '-', ':', '/', '.', '=', '%', '&'
    };
    /**
     * 在“W3C标准”和“RFC 2396”两种URL规范中存在争议的字符统一进行替换
     */
    private static final Map<Character, String> SPECIAL_CHARS;

    static {
        Map<Character, String> specialChars = new HashMap<>();
        specialChars.put(' ', "%20");
        specialChars.put('+', "%2B");
        SPECIAL_CHARS = Collections.unmodifiableMap(specialChars);
    }

    /**
     * 转义url
     */
    public static String escapeUrl(String url) throws UnsupportedEncodingException, RecruitmentSummaryException {
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append(escapePath(url));
        if (url.contains("?")) {
            stringBuilder.append("?");
            stringBuilder.append(escapeParameter(url));
        }
        return stringBuilder.toString();
    }

    /**
     * 转义url的路径
     */
    private static String escapePath(String url) throws RecruitmentSummaryException, UnsupportedEncodingException {
        if (!url.matches("http(s)?://.*")) {
            throw new RecruitmentSummaryException("url格式不正确：" + url);
        }
        int index = url.indexOf('?');
        if (index > 0) {
            url = url.substring(0, index);
        }
        char[] chars = url.toCharArray();
        StringBuilder stringBuilder = new StringBuilder();
        for (char aChar : chars) {
            toIso(aChar, stringBuilder);
        }
        return stringBuilder.toString();
    }

    /**
     * 转义参数
     */
    private static String escapeParameter(String url) throws UnsupportedEncodingException {
        String parameter = url.substring(url.indexOf('?') + 1);
        char[] chars = parameter.toCharArray();
        StringBuilder stringBuilder = new StringBuilder();
        for (char aChar : chars) {
            toIso(aChar, stringBuilder);
        }
        return stringBuilder.toString();
    }

    private static void toIso(char aChar, StringBuilder stringBuilder) throws UnsupportedEncodingException {
        if (isNumber(aChar)) {
            stringBuilder.append(aChar);
        } else if (isLetter(aChar)) {
            stringBuilder.append(aChar);
        } else if (isSpecialChar(aChar)) {
            stringBuilder.append(aChar);
        } else if (SPECIAL_CHARS.containsKey(aChar)) {
            stringBuilder.append(SPECIAL_CHARS.get(aChar));
        } else {
            stringBuilder.append(URLEncoder.encode(Character.toString(aChar), "UTF-8"));
        }
    }

    /**
     * 是数字
     */
    private static boolean isNumber(char aChar) {
        return aChar >= NUMBERS[0] && aChar <= NUMBERS[1];
    }

    /**
     * 是字母
     *
     * @param aChar
     * @return
     */
    private static boolean isLetter(char aChar) {
        if (aChar >= LOWER_CASE_LETTERS[0] && aChar <= LOWER_CASE_LETTERS[1]) {
            return true;
        } else {
            return aChar >= UPPERCASE_LETTER[0] && aChar <= UPPERCASE_LETTER[1];
        }
    }

    /**
     * 判断是否特殊字符，不需要转义的字符
     */
    private static boolean isSpecialChar(char aChar) {
        for (char specialChar : HTTP_CHARS) {
            if (specialChar == aChar) {
                return true;
            }
        }
        return false;
    }
}
