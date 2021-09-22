package cn.dream.anno.handler;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 设置行样式
 */
public class DefaultRowCellStyleAnnoHandler {

    /**
     * 设置行样式
     *  例: if(rowIndex % 2 == 0){
     *             rowCellStyle.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
     *             rowCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
     *         }
     * @param rowCellStyle 行单元格样式对象
     * @param value 数据项的值对象
     * @param rowIndex 行所索引
     */
    public void setRowStyle(CellStyle rowCellStyle,Object value,int rowIndex){

    }

}
