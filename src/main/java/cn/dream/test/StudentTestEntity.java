package cn.dream.test;


import cn.dream.anno.Excel;
import cn.dream.anno.ExcelField;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Getter
@Setter
@ToString
@Excel(name = "学生Excel")
public class StudentTestEntity {

    @ExcelField(name = "UID",mergeCell = true,autoSizeColumn = true)
    private String uid;

    @ExcelField(name = "学生名1称",cellStyleCls = DefaultExcelFieldStyleAnnoHandler02.class,autoSizeColumn = true)
    private String name;

    @ExcelField(name = "学生年龄",autoSizeColumn = true)
    private Integer age;

}
