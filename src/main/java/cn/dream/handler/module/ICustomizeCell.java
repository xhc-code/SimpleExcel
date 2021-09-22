package cn.dream.handler.module;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.text.ParseException;
import java.util.function.Function;

@FunctionalInterface
public interface ICustomizeCell {

    /**
     * 使用Sheet定义单元格和合并单元格及样式
     * @param workbook
     * @param sheet
     * @param putCellStyle 给Cell设置CellStyle对象的时候，请使用这个Lambda返回的CellStyle(已进行全局缓存样式)
     * @param setMergeCell 设置合并单元格的时候的便携方法
     */
    void customize(Workbook workbook, Sheet sheet, Function<CellStyle,CellStyle> putCellStyle,IAddMergeRegionCell setMergeCell) throws ParseException;


    @FunctionalInterface
    interface IAddMergeRegionCell {

        boolean setMergeCell(Sheet sheet, CellRangeAddress cellRangeAddress, Object value) throws ParseException;

    }

}
