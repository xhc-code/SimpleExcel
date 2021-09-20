package cn.dream.excep;

import cn.dream.excep.base.BaseExcelException;
import lombok.NoArgsConstructor;

/**
 * @author xiaohuichao
 * @createdDate 2021/9/20 17:06
 */
@NoArgsConstructor
public class ActionException extends BaseExcelException {

    public ActionException(String msg) {
        super(msg);
    }

}
