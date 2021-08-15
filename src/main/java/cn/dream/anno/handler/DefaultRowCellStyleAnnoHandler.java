package cn.dream.anno.handler;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * 设置行样式
 */
public class DefaultRowCellStyleAnnoHandler {

    /**
     * 设置行样式
     * @param rowCellStyle 行单元格样式对象
     * @param value
     * @param rowIndex 行所索引
     */
    public void setRowStyle(CellStyle rowCellStyle,Object value,int rowIndex){

        if(rowIndex % 2 == 0){
            rowCellStyle.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
            rowCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

    }

}
