package cn.dream.bo;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.lang.reflect.Field;
import java.util.List;

/**
 * 数据指针位置
 */
@Getter
@Setter
@ToString
public class DataPointBO {

    /**
     * 数据集合中的当前索引位置
     */
    private int index;

    /**
     * 当前处理的字段对象
     */
    private Field currentField;

    /**
     * 当前列值
     */
    private Object cellValue;

    /**
     * 当前行对象
     */
    private Object rowValue;

    /**
     * 数据列表
     */
    private List<Object> dataList;

}
