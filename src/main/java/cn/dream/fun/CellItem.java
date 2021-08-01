package cn.dream.fun;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import org.apache.commons.lang3.Validate;

@Getter
@Setter
@ToString
public class CellItem {

    /**
     * 对象值
     */
    private Object value;

    /**
     * 值类型
     */
    private Class<?> valueType;

    /**
     * 行索引，从0开始
     */
    private int rowIndex;

    /**
     * 列索引，从0开始
     */
    private int columnIndex;

    /**
     * 跨越行的数量
     */
    private int spanRowNum;

    /**
     * 跨越列的数量
     */
    private int spanColumnNum;

    /**
     * 表示此对象的单元格类型；true合并单元格，false普通单元格
     */
    private boolean mergeCell;

    private CellItem(){}

    /**
     * 验证对象的值是否在允许范围之内
     * @return
     */
    private void valid(){
        Validate.isTrue(rowIndex > 0);
        Validate.isTrue(columnIndex > 0);
        Validate.isTrue(spanRowNum > 0);
        Validate.isTrue(spanColumnNum > 0);
    }

    /**
     * 创建一个合并单元格
     * @return
     */
    public static CellItem newMergeCell(int spanRowIndex, int spanColumnIndex, Object value){
        CellItem cellItem = new CellItem();
        cellItem.setMergeCell(true);
        cellItem.setSpanRowNum(spanRowIndex);
        cellItem.setSpanColumnNum(spanColumnIndex);
        cellItem.setValue(value);
        cellItem.valid();
        return cellItem;
    }

    /**
     * 创建一个普通的单元格
     * @return
     */
    public static CellItem newCell(int spanRowIndex, int spanColumnIndex, Object value){
        CellItem cellItem = new CellItem();
        cellItem.setSpanRowNum(spanRowIndex);
        cellItem.setSpanColumnNum(spanColumnIndex);
        cellItem.setValue(value);
        cellItem.valid();
        return cellItem;
    }

    private boolean autoSizeColumn = false;

}
