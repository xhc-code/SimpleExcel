package cn.dream.test;

import cn.dream.anno.handler.excelfield.DefaultExcelFieldStyleAnnoHandler;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

public class DefaultExcelFieldStyleAnnoHandler02 extends DefaultExcelFieldStyleAnnoHandler {

    @Override
    protected void setHeaderCellStyle(CellStyle target) {
        super.setHeaderCellStyle(target);
        target.setFillForegroundColor(IndexedColors.YELLOW1.getIndex());
        target.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }
}
