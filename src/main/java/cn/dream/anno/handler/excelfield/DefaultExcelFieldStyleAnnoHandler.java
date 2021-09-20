package cn.dream.anno.handler.excelfield;

import cn.dream.enu.HandlerTypeEnum;
import cn.dream.excep.InvalidArgumentException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

public class DefaultExcelFieldStyleAnnoHandler {

	/**
	 * 设置单元格样式
	 * @param target 目标样式
	 * @param value 准备要写入的值
	 * @param handlerTypeEnum 设置样式单元格的类型
	 */
	public void cellStyle(CellStyle target,Object value, HandlerTypeEnum handlerTypeEnum) {
		if(handlerTypeEnum == HandlerTypeEnum.HEADER){
			setHeaderCellStyle(target);
		}else if(handlerTypeEnum == HandlerTypeEnum.BODY){
			setBodyCellStyle(target,value);
		}else{
			throw new InvalidArgumentException("无效设置单元格样式的Type类型");
		}
	}

	protected void setHeaderCellStyle(CellStyle target){
		target.setFillForegroundColor(IndexedColors.BLUE.getIndex());
		target.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	}

	protected void setBodyCellStyle(CellStyle target,Object value) {
		target.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
		target.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	}

}
