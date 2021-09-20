package cn.dream.excep;

import lombok.NoArgsConstructor;

/**
 * @author xiaohuichao
 * @createdDate 2021/9/20 17:06
 */
@NoArgsConstructor
public class ActionNotSupportedException extends ActionException{

    public ActionNotSupportedException(String msg) {
        super(msg);
    }

}
