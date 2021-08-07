package cn.dream.handler.module;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.function.Function;

@FunctionalInterface
public interface ICustomizeCell {

    /**
     * 使用Sheet定义单元格和合并单元格及样式
     * @param workbook
     * @param sheet
     * @param putCellStyle 给Cell设置CellStyle对象的时候，请使用这个Lambda返回的CellStyle(已进行全局缓存样式)
     */
    void customize(Workbook workbook, Sheet sheet, Function<CellStyle,CellStyle> putCellStyle);

}
