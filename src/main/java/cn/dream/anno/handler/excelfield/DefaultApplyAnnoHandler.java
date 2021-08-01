package cn.dream.anno.handler.excelfield;

import cn.dream.anno.ExcelField;
import org.springframework.stereotype.Component;

public class DefaultApplyAnnoHandler {

	/**
	 * 导出的字段过滤；指导出的字段列表是否包含此字段
	 * @return
	 */
	public boolean apply(ExcelField excelField,Type type) {
		return true;
	}

	public enum Type {
		IMPORT,EXPORT;
	}

}
