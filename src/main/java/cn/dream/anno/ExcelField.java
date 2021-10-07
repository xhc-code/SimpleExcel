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

	/**[导入时生效]
	 * 验证Header头
	 * @return
	 */
	FieldValidateHeaderConf validateHeaderConf() default @FieldValidateHeaderConf();

	/**[导出时生效]
	 * 单元格样式配置
	 * @return
	 */
	FieldCellStyleConf cellStyleConf() default @FieldCellStyleConf();
	
	/**[导出时生效]
	 * 只有设置Date和Calendar字段时，才会生效; 默认： yyyy-MM-dd HH:mm:ss  ==  2021-08-12 09:31:33
	 *  为空字符串或字段类型为Date时才会调用这个属性
	 * @return
	 */
	String dateFormat() default "yyyy-MM-dd HH:mm:ss";

	/**[导出时生效]
	 * 转换值表达式配置
	 * @return
	 */
	FieldConverterValueConf converterValueConf() default @FieldConverterValueConf();

	/**[导出时生效]
	 * 列宽大小；-1为不进行设置
	 * @return
	 */
	int columnWidth() default -1;

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

	/**【导出时生效】
	 * 处理写出的值及类型；可以处理最终写入到Excel中的值和类型
	 * @return
	 */
	Class<? extends DefaultWriteValueAnnoHandler> handlerWriteValue() default DefaultWriteValueAnnoHandler.class;

	/**[导出时生效]
	 * 合并单元格配置
	 * @return
	 */
	FieldMergeConf mergeConf() default @FieldMergeConf();

	/**[导出时生效]
	 * 下拉选择值配置
	 * @return
	 */
	FieldSelectValueConf selectValueConf() default @FieldSelectValueConf();

	/**
	 * 用户可自行传递的JSON数据,此数据会传递到每个阶段
	 * @return
	 */
	@FutureUse("功能未开发")
	String data() default "";

}

