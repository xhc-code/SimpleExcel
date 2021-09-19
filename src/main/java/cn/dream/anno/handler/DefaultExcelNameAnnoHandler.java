package cn.dream.anno.handler;

import cn.dream.util.DateUtils;

import java.util.Date;

public class DefaultExcelNameAnnoHandler {

	/**
	 * 年月日_{AM|PM}_时分秒
	 */
	private static final String DATE_TIME_FORMAT = "yyyyMMdd_a_HHmmss";

	/**
	 * 可根据自定义规则生成导出时的Excel名称
	 * @param name
	 * @return
	 */
	public String getName(String name) {
		String formatDate = DateUtils.formatDate(new Date(), DATE_TIME_FORMAT);
		return formatDate + "_" + name;
	}
		
}
