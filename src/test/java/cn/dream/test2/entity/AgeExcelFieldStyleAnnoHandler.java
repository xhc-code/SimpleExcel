package cn.dream.test2.entity;

import cn.dream.anno.handler.excelfield.DefaultExcelFieldStyleAnnoHandler;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * 年龄端分颜色
 */
public class AgeExcelFieldStyleAnnoHandler extends DefaultExcelFieldStyleAnnoHandler {

    @Override
    protected void setBodyCellStyle(CellStyle target, Object value) {
        super.setBodyCellStyle(target, value);

        if(value instanceof Integer){
            int v = (Integer) value;
            if(v < 24){
                target.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            }else if(v >= 24){
                target.setFillForegroundColor(IndexedColors.PINK.index);
            }


        }

    }
}
