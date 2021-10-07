package cn.dream.anno;

import cn.dream.anno.handler.excelfield.DefaultExcelFieldStyleAnnoHandler;

/**
 * 单元格样式配置
 * @author xiaohuichao
 * @createdDate 2021/10/5 16:16
 */
public @interface FieldCellStyleConf {


    /**[导出时生效]
     * 处理并设置单元格的样式
     * @return
     */
    Class<? extends DefaultExcelFieldStyleAnnoHandler> cellStyleCls() default DefaultExcelFieldStyleAnnoHandler.class;

}
