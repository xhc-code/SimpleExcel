package cn.dream.test;


import cn.dream.anno.Excel;
import cn.dream.anno.ExcelField;
import cn.dream.util.anno.Feature.RequireCopy;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.util.Date;

@Getter
@Setter
@ToString
@Excel(name = "学生Excel")
public class StudentTestEntity {

    @ExcelField(name = "UID",mergeCell = true,autoSizeColumn = true)
    private String uid;

    @RequireCopy
    @ExcelField(name = "学生名1称",cellStyleCls = DefaultExcelFieldStyleAnnoHandler02.class,autoSizeColumn = true)
    private String name;

    @RequireCopy
    @ExcelField(name = "学生年龄",autoSizeColumn = true)
    private Integer age;

    @ExcelField(name = "是否开放",converterValueExpression = "1=是,0=否",selectValues = "是,否")
    private Integer isPublic;


    @ExcelField(name = "生日",dateFormat = "yyyy/MM-dd HH:mm-ss")
    private Date birthday;

    @ExcelField(name = "成功了吗?")
    private Boolean success;


    @ExcelField(name = "成功了吗?")
    private Boolean success1;

}
