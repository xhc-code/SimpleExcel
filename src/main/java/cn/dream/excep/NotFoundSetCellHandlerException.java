package cn.dream.excep;

import cn.dream.excep.base.BaseExcelException;

/**
 * 没有找到设置单元格处理器异常
 * @author xiaohuichao
 * @createdDate 2021/10/2 17:12
 */
public class NotFoundSetCellHandlerException extends BaseExcelException {

    public NotFoundSetCellHandlerException() {
    }

    public NotFoundSetCellHandlerException(String message) {
        super(message);
    }
}
