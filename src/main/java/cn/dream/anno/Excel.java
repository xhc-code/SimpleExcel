package cn.dream.anno;

import cn.dream.anno.handler.DefaultExcelNameAnnoHandler;
import cn.dream.anno.handler.DefaultRowCellStyleAnnoHandler;
import cn.dream.enu.WorkBookTypeEnum;

import java.lang.annotation.*;

/**
 * Excel定义实体的全局规则
 * @author Dream
 */
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Target(ElementType.TYPE)
public @interface Excel {

	/**[导出时生效]
	 * 导出的Excel名称
	 * @return
	 */
	String name();


	/**[导出时生效]
	 * 行高大小；-1为不进行设置
	 * @return
	 */
	short defaultRowHeight() default -1;
	
	/**[导出时生效]
	 * 导出时Excel生成的名称
	 * @return
	 */
	Class<? extends DefaultExcelNameAnnoHandler> handlerName() default DefaultExcelNameAnnoHandler.class;

	/**
	 * 导出Excel文件扩展名的类型
	 * @return
	 */
	WorkBookTypeEnum extendFileType() default WorkBookTypeEnum.XLSX;

	/**[导出时生效]
	 * 处理设置行样式
	 * @return
	 */
	Class<? extends DefaultRowCellStyleAnnoHandler> handleRowStyle() default DefaultRowCellStyleAnnoHandler.class;

	/**[导出时生效]
	 * <span style='color:red'>仅导出时生效</span><br />
	 * 生成的起始行索引，从0开始，-1为默认自动;全局行索引位置
	 * @return
	 */
	int rowIndex() default -1;

	/**[{导出}时生效]
	 * <span style='color:red'>仅导出时生效</span><br />
	 * 生成的起始列索引，从0开始,-1为默认自动；全局列索引位置
	 * @return
	 */
	int columnIndex() default -1;

	/**[{读取}时生效]
	 * <span style='color:red'>仅导入时生效</span><br />
	 * Header表头的首行索引位置；不设置默认从0开始，也就是Excel第一行开始
	 * @return
	 */
	int[] headerRowRangeIndex() default {};

	/**[{读取}时生效]
	 * <span style='color:red'>仅导入时生效</span><br />
	 * Body数据的首行索引位置；不设置默认从1开始，也就是Excel第二行开始
	 * @return
	 */
	int dataFirstRowIndex() default -1;


	/**[读取时生效]
	 * 是否根据 HeaderName进行填充值；true是根据headerName填充,false是根据列索引填充值
	 * @return
	 */
	boolean byHeaderName() default false;

}
