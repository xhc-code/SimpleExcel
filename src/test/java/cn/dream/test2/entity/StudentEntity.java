package cn.dream.test2.entity;

import cn.dream.anno.*;
import lombok.*;

import java.util.Date;

@Getter
@Setter
@ToString
@Builder
@AllArgsConstructor
@NoArgsConstructor
@Excel(name = "学生文件")
public class StudentEntity {

    @ExcelField(name = "学生编号")
    private String uid;

    @ExcelField(name = "学生姓名")
    private String name;

    /**
     * 年份
     */
    @ExcelField(name = "年龄",cellStyleConf = @FieldCellStyleConf(cellStyleCls = AgeExcelFieldStyleAnnoHandler.class))
    private Integer age;

    /**
     * 创建毫秒值
     */
    @ExcelField(name = "创建的毫秒值")
    private Long createMillisecond;

    @ExcelField(name = "生日日期")
    private Date birthdayDate;

    @ExcelField(name = "是否公开",converterValueConf = @FieldConverterValueConf(valueExpression = "1=是,0=否"))
    private Integer isPublic;

    @ExcelField(name = "省份名称",mergeConf = @FieldMergeConf(mergeCell = true))
    private String provinceName;

    @ExcelField(name = "市区名称",mergeConf = @FieldMergeConf(mergeCell = true))
    private String cityName;

    @ExcelField(name = "区县名称")
    private String distinctName;

    @ExcelField(name = "是否成功")
    private Boolean success;

    @ExcelField(name = "生日")
    private Date birthday;


}
