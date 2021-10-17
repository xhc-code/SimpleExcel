package cn.dream.anno;

import cn.dream.anno.handler.excelfield.DefaultConverterValueAnnoHandler;

/**
 * 字典转换值
 * @author xiaohuichao
 * @createdDate 2021/10/5 16:12
 */
public @interface FieldConverterValueConf {

    /**[导出和导入时生效]
     * 读取内容表达式；示例：1=男,2=女,3=未知; 导入时会将文字值转换成对应的数值，适用于枚举形式
     * @return
     */
    String valueExpression() default "";

    /**[导出和导入时生效]
     * 读取内容表达式,可以从Bean容器中获取
     * @return
     */
    Class<? extends DefaultConverterValueAnnoHandler> valueCls() default DefaultConverterValueAnnoHandler.class;

    /**[导出和导入时生效]
     * 启用转换器多值处理;true代表启用,false代表不启用多值匹配
     * @return
     */
    boolean enableMultiValue() default false;

    /**
     * 多值匹配转换的分隔符
     * @return
     */
    String delimiter() default ",";

}
