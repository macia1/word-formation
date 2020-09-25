package com.zhiwei.util;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.sun.istack.internal.NotNull;
import com.zhiwei.word.Bookmark;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.extern.log4j.Log4j2;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * excel读取监听器
 *
 * @author 朽木不可雕也
 * @date 2020/9/18 11:15
 * @week 星期五
 */
@EqualsAndHashCode(callSuper = true)
@Log4j2
@Data
public class ExcelMonitor extends AnalysisEventListener<Map<Integer, String>> {
    /**
     * 文件的标题
     */
    private final Map<String, Integer> headers = new HashMap<>();
    private final List<List<Object>> dataList = new ArrayList<>();
    private int dataArrayLength;

    @Override
    public void invokeHeadMap(@NotNull Map<Integer, String> headMap, AnalysisContext context) {
        headMap.forEach((column, title) -> {
            headers.put(title, column);
        });
        this.dataArrayLength = headers.size() + 1;//多出一个位置用于存储书签信息
    }

    @Override
    public void invoke(@NotNull Map<Integer, String> integerStringMap, AnalysisContext analysisContext) {
        //根据列的行数转换一条数据，防止easyExcel自动过滤空单元格
        Object[] data = new Object[this.dataArrayLength];
        //将元素有序插入数组
        integerStringMap.forEach((column, value) -> {
            data[column] = value;
        });
        data[data.length - 1] = new Bookmark("t" + this.dataList.size(), "r" + this.dataList.size());
        log.info(Arrays.toString(data));
        this.dataList.add(Arrays.asList(data));
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        log.info("数据读取完毕，总计：{}条", this.dataList.size());
    }
}
