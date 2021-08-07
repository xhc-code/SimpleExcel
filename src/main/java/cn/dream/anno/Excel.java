package cn.dream.anno;

import java.lang.annotation.*;

import cn.dream.anno.handler.DefaultExcelNameAnnoHandler;

/**
 * Excel定义实体的全局规则
 * @author Dream
 */
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Target(ElementType.TYPE)
public @interface Excel {

	/**
	 * 导出的Excel名称
	 * @return
	 */
	String name();
	
	/**
	 * 导出时Excel生成的名称
	 * @return
	 */
	Class<? extends DefaultExcelNameAnnoHandler> handlerName() default DefaultExcelNameAnnoHandler.class;

	/**
	 * <span style='color:red'>仅导出时生效</span><br />
	 * 生成的起始行索引，从0开始，-1为默认自动;全局行索引位置
	 * @return
	 */
	int rowIndex() default -1;

	/**
	 * <span style='color:red'>仅导出时生效</span><br />
	 * 生成的起始列所以，从0开始,-1为默认自动；全局列索引位置
	 * @return
	 */
	int columnIndex() default -1;

	/**
	 * <span style='color:red'>仅导入时生效</span><br />
	 * Header表头的首行索引位置；不设置默认从0开始，也就是Excel第一行开始
	 * @return
	 */
	int[] headerRowRangeIndex() default {};

	/**
	 * <span style='color:red'>仅导入时生效</span><br />
	 * Body数据的首行索引位置；不设置默认从0开始，也就是Excel第二行开始
	 * @return
	 */
	int dataFirstRowIndex() default -1;


	/**
	 * 是否根据 HeaderName进行填充值
	 * @return
	 */
	boolean byHeaderName() default false;
			// TODO 应该会有个问题，手动添加合并单元格，设置的样式好像会有问题，无法全部生效

	/**
	 * 必选的
	 * @return
	 */
	boolean required() default true;

}
