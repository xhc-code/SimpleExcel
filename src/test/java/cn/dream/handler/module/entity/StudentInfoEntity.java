package cn.dream.handler.module.entity;

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
    private Long id;

    /**
     * 用户ID
     */
    private Long userId;

    /**
     * 会员ID
     */
    private Long memberId;

    /**
     * 用户名称
     */
    private String userName;

    /**
     * 性别
     */
    private Character sex;

    /**
     * 年龄
     */
    private Short age;

    /**
     * 生日
     */
    private Date birthday;

    /**
     * 创建人ID
     */
    private Long createBy;

    /**
     * 创建人名称
     */
    private String createName;

    /**
     * 创建日期时间
     */
    private Date createDate;

}
