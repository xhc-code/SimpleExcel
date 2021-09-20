package cn.dream.excep;

import cn.dream.excep.base.BaseExcelException;
import lombok.NoArgsConstructor;

/**
 * @author xiaohuichao
 * @createdDate 2021/9/21 7:37
 */
@NoArgsConstructor
public class InvalidArgumentException extends BaseExcelException {

    public InvalidArgumentException(String msg){
        super(msg);
    }


}


