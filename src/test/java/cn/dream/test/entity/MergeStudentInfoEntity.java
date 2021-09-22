package cn.dream.test.entity;

import cn.dream.anno.Excel;
import cn.dream.anno.ExcelField;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.util.Date;

/**
 * @author xiaohuichao
 * @createdDate 2021/9/21 13:01
 */
@Getter
@Setter
@Excel(name = "test",byHeaderName = true,headerRowRangeIndex = {1,2})
@ToString
public class MergeStudentInfoEntity {

    @ExcelField(name = "用户UID")
    private String uid;

    @ExcelField(name = "用户名称")
    private String userName;

    @ExcelField(name = "用户年龄")
    private Integer age;

    @ExcelField(name = "用户性别")
    private Character sex;

    @ExcelField(name = "生日日期")
    private Date birthday;

    @ExcelField(name = "记录日期")
    private String recordDate;

    @ExcelField(name = "创建ID")
    private Integer createBy;

    @ExcelField(name = "创建名称",mergeCell = true)
    private String createName;

    @ExcelField(name = "审核状态",converterValueExpression = "1=编辑,2=审核中,3=审核成功,4=审核失败")
    private Integer auditStatus;

    @ExcelField(name = "是否公开",converterValueExpression = "1=公开,2=未公开")
    private Integer isPublic;

}
