package cn.dream.handler;

public interface IExcelNameHandler {

	/**
	 * 可根据自定义规则生成导出时的Excel名称
	 * @param name
	 * @return
	 */
	String getName(String name);
		
}
