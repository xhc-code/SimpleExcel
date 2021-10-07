package cn.dream.anno;

import cn.dream.anno.handler.excelfield.DefaultSelectValueListAnnoHandler;

import java.lang.annotation.*;

/**
 * 选择下拉值的字段
 * @author xiaohuichao
 * @createdDate 2021/10/5 13:56
 */
@Target(ElementType.ANNOTATION_TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface FieldSelectValueConf {


    /**
     * Excel可选择的值列表,多个值以逗号分割，Excel中只能单值验证
     * @return
     */
    String selectValues() default "";

    /**
     * 是否解析 {@link ExcelField#converterValueExpression()} 的表达式作为下拉框的值
     * @return
     */
    boolean buildFromValueExpression() default true;

    /**
     * 选择值的列表
     * @return
     */
    Class<? extends DefaultSelectValueListAnnoHandler> selectValueListCls() default DefaultSelectValueListAnnoHandler.class;

}
