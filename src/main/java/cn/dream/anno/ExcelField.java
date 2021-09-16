package cn.dream.anno;

import cn.dream.anno.handler.excelfield.*;
import cn.dream.anno.mark.FutureUse;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.FIELD;

@Retention(RetentionPolicy.RUNTIME)
@Documented
@Target(FIELD)
public @interface ExcelField {
	
	/**[导出时生效][导入时验证标头生效]
	 * 标题头
	 * @return
	 */
	String name();

	/**[导入时验证标头生效]
	 * 验证的HeaderName名称,如果validateHeader为true,并且 validateHeaderName非空，从此名称进行获取并验证HeaderName,否则 值来源与name
	 * @return
	 */
	String validateHeaderName() default "";

	/**[导入时验证标头生效]
	 * 验证标题头是否名称一致，这是为了防止导入错误的Excel数据;仅在根据索引位置填充数据时有效
	 * @return true标识需要验证标题头，false不验证
	 */
	boolean validateHeader() default false ;
	
	/**[导出和导入时生效]
	 * Excel可选择的值列表,多个值以逗号分割，Excel中只能单值验证
	 * @return
	 */
	String selectValues() default "";
	
	/**[导出时生效]
	 * 返回可选择值的列表
	 * @return
	 */
	Class<? extends DefaultSelectValueListAnnoHandler> selectValueListCls() default DefaultSelectValueListAnnoHandler.class;
	
	/**[导出时生效]
	 * 处理并设置单元格的样式
	 * @return
	 */
	Class<? extends DefaultExcelFieldStyleAnnoHandler> cellStyleCls() default DefaultExcelFieldStyleAnnoHandler.class;
	
	/**[导出时生效]
	 * 只有设置Date和Calendar字段时，才会生效; 默认： yyyy-MM-dd HH:mm:ss  ==  2021-08-12 09:31:33
	 *  为空字符串或字段类型为Date时才会调用这个属性
	 * @return
	 */
	String dateFormat() default "yyyy-MM-dd HH:mm:ss";
	
	/**[导出和导入时生效]
	 * 读取内容表达式；示例：1=男,2=女,3=未知; 导入时会将文字值转换成对应的数值，适用于枚举形式
	 * @return
	 */
	String converterValueExpression() default "";
	
	/**[导出和导入时生效]
	 * 读取内容表达式,可以从Bean容器中获取
	 * @return
	 */
 	Class<? extends DefaultConverterValueAnnoHandler> converterValueCls() default DefaultConverterValueAnnoHandler.class;

	/**[导出和导入时生效]
	 * 启用转换器多值处理;true代表启用,false代表不启用多值匹配
	 * @return
	 */
	boolean enableConverterMultiValue() default false;

 	/**[导出时生效]
 	 * 是否开启自动列宽
 	 * @return
 	 */
 	boolean autoSizeColumn() default false;
 	
 	/**[导出时生效]
 	 * 当字段为空时的默认值
 	 * @return
 	 */
 	String defaultValue() default "";

	/**[导出和导入时生效]
	 * 是否应用此字段
	 * @return
	 */
 	boolean apply() default true;

	/**【写入Excel有作用，预留功能】
	 * 修改值和类型的一个阶段；处于 值获取后，但处于写入Excel字段值之前，此before处理完之后，将其中的值和类型作为最终结果写入到Excel中
	 * @return
	 */
	Class<? extends DefaultWriteValueAnnoHandler> handlerWriteValue() default DefaultWriteValueAnnoHandler.class;

	/**[导出时生效]
	 * 是否是合并单元格
	 * @return
	 */
	boolean mergeCell() default false;

	/**[导出时生效]
	 * 当 mergeRow = true时，此属性指 合并的组键使用多个值进行唯一性，这里的数组指示了组Key的组成部分;默认不包含当前字段的值
	 * @return
	 */
	MergeField[] mergeFields() default {};

	/**
	 * 用户可自行传递的JSON数据,此数据会传递到每个阶段
	 * @return
	 */
	@FutureUse("功能未开发")
	String dataJson() default "";

}

