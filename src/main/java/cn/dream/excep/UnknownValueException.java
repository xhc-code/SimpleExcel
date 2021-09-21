package cn.dream.excep;

import cn.dream.excep.base.BaseExcelException;
import lombok.NoArgsConstructor;

/**
 * @author xiaohuichao
 * @createdDate 2021/9/21 10:35
 */
@NoArgsConstructor
public class UnknownValueException extends BaseExcelException {

    public UnknownValueException(String msg){
        super(msg);
    }


}
