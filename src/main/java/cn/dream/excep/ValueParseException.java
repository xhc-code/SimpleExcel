package cn.dream.excep;

/**
 * 值转换解析异常
 * @author xiaohuichao
 * @createdDate 2021/10/1 11:30
 */
public class ValueParseException extends ExcelRuntimeException{
    public ValueParseException(String msg) {
        super(msg);
    }
}
