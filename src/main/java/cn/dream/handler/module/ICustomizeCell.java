package cn.dream.handler.module;

import cn.dream.handler.module.helper.CellHelper;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.text.ParseException;

@FunctionalInterface
public interface ICustomizeCell {

    /**
     * 自定义处理Sheet单元格对象
     *
     * 如果要使用 {@code workbook.createCellStyle()} 之类创建对象，请尽可能留存实例，以进行复用，可通过cacheStyle进行缓存并返回cloneFrom你创建的CellStyle可用的CellStyle对象
     * @param workbook 工作簿对象
     * @param sheet Sheet对象，后续的操作都是基于此Sheet对象上进行操作的
     * @param cacheStyle 缓存工作簿创建的CellStyle对象，可以重用之前通过Workbook创建的CellStyle对象
     * @param cellHelper 基于Sheet实例提供的操作单元格的帮助工具
     */
    void customize(Workbook workbook, Sheet sheet, WorkbookCacheStyle cacheStyle, CellHelper cellHelper) throws ParseException;


    @FunctionalInterface
    interface WorkbookCacheStyle {

        CellStyle cache(CellStyle cellStyle);

    }

}
