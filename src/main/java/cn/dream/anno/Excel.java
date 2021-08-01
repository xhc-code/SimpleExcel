package cn.dream.anno;

import static java.lang.annotation.ElementType.FIELD;

import java.lang.annotation.*;
import java.util.function.Function;

import cn.dream.anno.handler.DefaultExcelNameAnnoHandler;
import cn.dream.handler.IExcelNameHandler;

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
	 * 生成的起始行索引，从0开始，-1为默认自动
	 * @return
	 */
	int rowIndex() default -1;

	/**
	 * 生成的起始列所以，从0开始,-1为默认自动
	 * @return
	 */
	int columnIndex() default -1;
	
}
