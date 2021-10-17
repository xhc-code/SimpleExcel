package cn.dream.anno;

/**
 * 合并单元格配置
 * @author xiaohuichao
 * @createdDate 2021/10/5 16:14
 */
public @interface FieldMergeConf {

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

}
