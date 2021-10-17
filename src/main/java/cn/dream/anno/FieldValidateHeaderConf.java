package cn.dream.anno;

/**
 * 验证Header名称的注解
 * @author xiaohuichao
 * @createdDate 2021/10/5 16:07
 */
public @interface FieldValidateHeaderConf {

    /**
     * 验证的HeaderName名称,如果validateHeader为true,并且 validateHeaderName非空，从此名称进行获取并验证HeaderName,否则 值来源与name
     * @return
     */
    String headerName() default "";

    /**
     * 验证标题头是否名称一致，这是为了防止导入错误的Excel数据;仅在根据索引位置填充数据时有效
     * @return true标识需要验证标题头，false不验证
     */
    boolean validation() default false;

}
