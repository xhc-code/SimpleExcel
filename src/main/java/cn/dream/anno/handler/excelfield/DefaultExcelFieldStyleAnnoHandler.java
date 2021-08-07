package cn.dream.anno.handler.excelfield;

import cn.dream.enu.HandlerTypeEnum;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

public class DefaultExcelFieldStyleAnnoHandler {

	public void headerCellStyle(CellStyle origin,CellStyle target){
		cellStyle(origin,target,HandlerTypeEnum.HEADER);
	}


	public void bodyCellStyle(CellStyle origin,CellStyle target){
		cellStyle(origin,target,HandlerTypeEnum.BODY);
	}

	/**
	 * 设置单元格样式
	 * @param origin model样式
	 * @param target 目标样式
	 * @param handlerTypeEnum 设置样式单元格的类型
	 */
	public void cellStyle(CellStyle origin, CellStyle target, HandlerTypeEnum handlerTypeEnum) {
		cloneStyleFrom(origin,target);
		if(handlerTypeEnum == HandlerTypeEnum.HEADER){
			setHeaderCellStyle(target);
		}else if(handlerTypeEnum == HandlerTypeEnum.BODY){
			setBodyCellStyle(target);
		}else{
			throw new RuntimeException("无效设置单元格样式的Type类型");
		}
	}

	protected void setHeaderCellStyle(CellStyle target){
		target.setFillForegroundColor(IndexedColors.BLUE.getIndex());
		target.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	}

	protected void setBodyCellStyle(CellStyle target) {
		target.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
		target.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	}

	private void cloneStyleFrom(CellStyle origin,CellStyle target) {
		if(origin != null){
			target.cloneStyleFrom(origin);
		}
	}

}
