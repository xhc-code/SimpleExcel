package cn.dream.handler.module.entity;

import cn.dream.anno.ExcelField;
import cn.dream.anno.FieldMergeConf;
import lombok.Builder;
import lombok.Data;

import java.util.Date;

/**
 * @author xiaohuichao
 * @createdDate 2021/10/4 14:19
 */
@Builder
@Data
public class StudentInfoEntity {

    /**
     * 主键ID
     */
    @ExcelField(name = "ID主键")
    private Long id;

    /**
     * 用户ID
     */
    @ExcelField(name = "用户ID",mergeConf = @FieldMergeConf(mergeCell = true))
    private Long userId;

    /**
     * 会员ID
     */
    @ExcelField(name = "会员ID")
    private Long memberId;

    /**
     * 用户名称
     */
    @ExcelField(name = "用户名称")
    private String userName;

    /**
     * 性别
     */
    @ExcelField(name = "性别")
    private Character sex;

    /**
     * 年龄
     */
    @ExcelField(name = "年龄")
    private Short age;

    /**
     * 生日
     */
    @ExcelField(name = "生日",columnWidth = 25*256)
    private Date birthday;

    /**
     * 创建人ID
     */
    @ExcelField(name = "创建人ID")
    private Long createBy;

    /**
     * 创建人名称
     */
    @ExcelField(name = "创建人名称")
    private String createName;

    /**
     * 创建日期时间
     */
    @ExcelField(name = "创建日期时间")
    private Date createDate;

}
