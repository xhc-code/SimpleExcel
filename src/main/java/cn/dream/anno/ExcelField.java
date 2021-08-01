package cn.dream.anno;

import static java.lang.annotation.ElementType.FIELD;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.util.function.Function;

import cn.dream.anno.handler.excelfield.*;

@Retention(RetentionPolicy.RUNTIME)
@Documented
@Target(FIELD)
public @interface ExcelField {
	
	/**
	 * 标题头
	 * @return
	 */
	String name();

	/**
	 * 字段的值类型,以注解的type值优先，没有设置则为属性的类型
	 * @return
	 */
	Class<?> type() default Class.class;

	/**
	 * 验证标题头是否名称一致，这是为了防止导入错误的Excel数据
	 * @return true标识需要验证标题头，false不验证
	 */
	boolean validateHeader() default false ;
	
	/**
	 * Excel可选择的值列表,多个值以逗号分割
	 * @return
	 */
	String selectValues() default "";
	
	/**
	 * 返回可选择值的列表
	 * @return
	 */
	Class<? extends DefaultSelectValueListAnnoHandler> selectValueListCls() default DefaultSelectValueListAnnoHandler.class;
	
	/**
	 * 处理并设置单元格的样式
	 * @return
	 */
	Class<? extends DefaultExcelFieldStyleAnnoHandler> cellStyleCls() default DefaultExcelFieldStyleAnnoHandler.class;
	
	/**
	 * 格式化值的对象
	 * @return
	 */
	Class<? extends DefaultFormatValueAnnoHandler> formatValueCls() default DefaultFormatValueAnnoHandler.class;
	
	/**
	 * 读取内容表达式；示例：1=男,2=女,3=未知; 导入时会将文字值转换成对应的数值，适用于枚举形式
	 * @return
	 */
	String converterValueExpression() default "";
	
	/**
	 * 读取内容表达式,可以从Bean容器中获取
	 * @return
	 */
 	Class<? extends DefaultConverterValueAnnoHandler> converterValueCls() default DefaultConverterValueAnnoHandler.class;
	
 	/**
 	 * 是否开启自动列宽
 	 * @return
 	 */
 	boolean autoSizeColumn() default false;
 	
 	/**
 	 * 当字段为空时的默认值
 	 * @return
 	 */
 	String defaultValue() default "";

	/**
	 * 是否应用此字段
	 * @return
	 */
 	boolean apply() default true;

 	/**
 	 * 使用进行映射时，是否包含此字段的条件，默认是 true
 	 * @return
 	 */
 	Class<? extends DefaultApplyAnnoHandler> applyCls() default DefaultApplyAnnoHandler.class;

	/**
	 * 修改值和类型的一个阶段；处于 值获取后，但处于写入Excel字段值之前，此before处理完之后，将其中的值和类型作为最终结果写入到Excel中
	 * @return
	 */
	Class<? extends DefaultWriteValueAnnoHandler> handlerWriteValue() default DefaultWriteValueAnnoHandler.class;

	/**
	 * 是否是合并单元格
	 * @return
	 */
	boolean mergeCell() default false;

	/**
	 * 当 mergeRow = true时，此属性指 合并的组键使用多个值进行唯一性，这里的数组指示了组Key的组成部分
	 * @return
	 */
	MergeField[] mergeFields() default {};

}


/**
 *
 * 导入时
 * 		validateHeader
 *
 * 值写入顺序
 *
 *   apply
 *   	autoColumnHeight
 *   	cellStyleCls
 *   	handlerWriteValue ↘
 *   		formatValueCls
 *   			if val==null defaultValue
 *   			converterValueExpression | converterValueCls
 *   			selectValues | selectValueListCls
 *
 *
 *
 *
 *
 *
 *
 * */
