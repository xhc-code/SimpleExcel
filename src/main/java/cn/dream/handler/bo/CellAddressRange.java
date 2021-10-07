package cn.dream.handler.bo;

import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

/**
 * 模式位置
 *  属性的值都是以0起始，与Sheet表示的行和列一致
 */
@Getter
@Setter
@ToString
@Builder
public class CellAddressRange {

    /**
     * 首行
     */
    private int firstRow;

    /**
     * 尾行
     */
    private int lastRow;

    /**
     * 首列
     */
    private int firstCol;

    /**
     * 尾列
     */
    private int lastCol;

}
