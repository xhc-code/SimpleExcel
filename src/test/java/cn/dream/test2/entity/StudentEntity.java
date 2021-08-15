package cn.dream.test2.entity;

import cn.dream.anno.Excel;
import cn.dream.anno.ExcelField;
import cn.dream.anno.MergeField;
import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.util.Date;

@Getter
@Setter
@ToString
@Builder
@Excel(name = "学生文件")
public class StudentEntity {

    @ExcelField(name = "学生编号")
    private String uid;

    @ExcelField(name = "学生姓名")
    private String name;

    /**
     * 年份
     */
    @ExcelField(name = "年龄")
    private Integer age;

    /**
     * 创建毫秒值
     */
    @ExcelField(name = "创建的毫秒值")
    private Long createMillisecond;

    @ExcelField(name = "生日日期")
    private Date birthdayDate;

    @ExcelField(name = "是否公开",selectValues = "是,否",converterValueExpression = "1=是,0=否")
    private Byte isPublic;

    @ExcelField(name = "省份名称",mergeCell = true)
    private String provinceName;

    @ExcelField(name = "市区名称",mergeCell = true)
    private String cityName;

    @ExcelField(name = "区县名称")
    private String distinctName;

    @ExcelField(name = "p1",mergeCell = true)
    private String p1;
    @ExcelField(name = "p2",mergeCell = true,mergeFields = { @MergeField(fieldName = "p1")})
    private String p2;
    @ExcelField(name = "p3",mergeCell = true)
    private String p3;


}
