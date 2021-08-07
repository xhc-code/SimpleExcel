package cn.dream.handler.bo;

import cn.dream.enu.HandlerTypeEnum;
import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.lang.reflect.Field;

/**
 * 记录数据验证的信息，用于当push完毕之后，使用此对象创建相应的验证
 */
@Getter
@Setter
@ToString
@Builder
public class RecordDataValidator {

    /**
     * 范围地址
     */
    private CellAddressRange cellAddressRange;

    /**
     * 选择项数组
     */
    private String[] selectedItems;

    private HandlerTypeEnum handlerTypeEnum;

    /**
     * 数据Cls
     */
    private Class<?> dataCls;

    /**
     * 字段对象
     */
    private Field field;

    /**
     * 对象信息
     */
    private Object o;


}
