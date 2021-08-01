package cn.dream.enu;

import cn.dream.anno.mark.FutureUse;

public enum HandlerTypeEnum {
    /**
     * 处理Header表头
     */
    HEADER,
    /**
     * 处理Body主体数据
     */
    BODY,
    /**
     * 自定义类型，未来使用
     */
    @FutureUse
    CUSTOM;

}
