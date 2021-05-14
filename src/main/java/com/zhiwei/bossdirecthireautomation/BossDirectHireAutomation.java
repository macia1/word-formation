package com.zhiwei.bossdirecthireautomation;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.event.SyncReadListener;
import com.zhiwei.bossdirecthireautomation.exceptions.BossDirectHireAutomationException;
import com.zhiwei.bossdirecthireautomation.exceptions.DataSourceAbnormalException;
import com.zhiwei.bossdirecthireautomation.tree.EventExcelEntity;
import com.zhiwei.bossdirecthireautomation.tree.PolyTreeUtil;
import com.zhiwei.bossdirecthireautomation.wordutil.WorldUtil;
import lombok.Getter;
import lombok.extern.log4j.Log4j2;
import org.apache.commons.lang3.StringUtils;

import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;

/**
 * Boss直聘自动生成程序
 *
 * @author fuck-aszswaz
 */
@Log4j2
public class BossDirectHireAutomation {
    /**
     * 品牌数据源
     */
    @Getter
    private final List<EventExcelEntity> brandData;

    private BossDirectHireAutomation(List<EventExcelEntity> brandData) {
        this.brandData = brandData;
    }

    /**
     * 校验数据
     *
     * @param eventExcelEntities 数据源
     */
    private static void checkData(@org.jetbrains.annotations.NotNull List<EventExcelEntity> eventExcelEntities) throws Exception {
        log.info("开始检查数据源的正确性");
        if (eventExcelEntities.isEmpty())
            throw new BossDirectHireAutomationException("源数据列表或平台影响力列表不能为空！");
        ClassUtil<EventExcelEntity> classUtil = new ClassUtil<>(EventExcelEntity.class);
        Map<Field, Method> fieldMethodMap = classUtil.getGetMethod();
        Set<Map.Entry<Field, Method>> entries = fieldMethodMap.entrySet();
        for (EventExcelEntity eventExcelEntity : eventExcelEntities) {
            for (Map.Entry<Field, Method> entry : entries) {
                Method method = entry.getValue();
                Object object = method.invoke(eventExcelEntity);
                if (Objects.isNull(object)) {
                    Field field = entry.getKey();
                    if (field.isAnnotationPresent(NotNull.class) && field.isAnnotationPresent(ExcelProperty.class)) {
                        ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
                        String[] value = excelProperty.value();
                        throw new DataSourceAbnormalException("列： " + value[0] + " 不能为空");
                    }
                }
            }
        }
        log.info("数据源检查完毕，无异常数据");
    }

    /**
     * 生成world文档
     *
     * @param writePath 文档写出路径
     */
    public void generateWord(String writePath) throws IOException {
        // 生成world文档
        log.info("开始生成world文档");

        PolyTreeUtil treeUtil = new PolyTreeUtil();
        // 将各个数据封装到对应的节点
        this.brandData.forEach(treeUtil::add);
        // 排序
        treeUtil.sort();

        // 开始生成word
        WorldUtil worldUtil = new WorldUtil(treeUtil);
        worldUtil.generate();
        worldUtil.write(writePath);

        log.info("world文档生成完毕！");
    }

    /**
     * 读取数据文件，构造执行对象
     *
     * @param sourcePath 数据源文件爱你
     */
    @SuppressWarnings("unused")
    public static BossDirectHireAutomation build(String sourcePath) throws BossDirectHireAutomationException {
        return build(sourcePath, null);
    }

    /**
     * 读取文件并构造执行对象
     *
     * @param sourcePath  数据源文件
     * @param channelPath 渠道影响力文件
     * @return 执行对象
     */
    public static BossDirectHireAutomation build(String sourcePath, String channelPath) throws BossDirectHireAutomationException {
        // 读取sheet 1
        List<EventExcelEntity> brandData = EasyExcel.read(sourcePath, EventExcelEntity.class, new SyncReadListener()).sheet(0).doReadSync();

        // 判断是否需要读取渠道影响力文件
        if (StringUtils.isNoneBlank(channelPath)) {
            List<EventExcelEntity.Influence> influences = EasyExcel.read(channelPath, EventExcelEntity.Influence.class, new SyncReadListener()).sheet(0).doReadSync();
            BossDirectHireAutomation.matchChannelInfluence(brandData, influences);// 匹配渠道影响力
        }

        // 分别读取两批数据
        return BossDirectHireAutomation.build(brandData);
    }

    /**
     * 构造执行对象
     *
     * @param brandData 品牌数据源
     */
    public static BossDirectHireAutomation build(List<EventExcelEntity> brandData) throws BossDirectHireAutomationException {
        try {
            BossDirectHireAutomation.checkData(brandData);

            return new BossDirectHireAutomation(brandData);
        } catch (BossDirectHireAutomationException e) {
            throw e;
        } catch (Exception e) {
            throw new BossDirectHireAutomationException(e.getMessage(), e);
        }
    }

    /**
     * 匹配渠道影响力
     *
     * @param brandData  品牌数据
     * @param influences 渠道影响力数据
     */
    private static void matchChannelInfluence(List<EventExcelEntity> brandData, List<EventExcelEntity.Influence> influences) {
        Map<String, Double> influencesMap = new HashMap<>();
        influences.forEach(influence -> influencesMap.put(influence.getPlatform() + "-" + influence.getChannel(), influence.getInfluence()));
        brandData.forEach(eventExcelEntity -> {
            Double influence = influencesMap.get(eventExcelEntity.getPlatform() + "-" + eventExcelEntity.getChannel());
            influence = Objects.isNull(influence) ? 0D : influence;
            eventExcelEntity.setInfluence(influence);
        });
    }
}
