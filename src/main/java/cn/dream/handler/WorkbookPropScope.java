package cn.dream.handler;

import cn.dream.util.anno.Feature.RequireCopy;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;

/**
 * 声明仅在Workbook对象生命周期内的对象，仅一份实例
 */
public abstract class WorkbookPropScope {

    /* ===========                  需要此次WordBook共享的对象                      =========================  */

    protected static final String[] IGNORE_PROP = new String[]{"globalCellStyle","cacheCellStyleMap"};

    /**
     * 单元格样式缓存Map,避免创建过多的样式对象
     */
    @RequireCopy
    protected Map<Integer, CellStyle> cacheCellStyleMap;

    @RequireCopy
    protected Workbook workbook;

}