package cn.dream.util;

import cn.dream.excep.InvalidArgumentException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationContext;
import org.springframework.core.convert.ConversionService;
import org.springframework.stereotype.Component;

import javax.annotation.PostConstruct;

/**
 * 值类型转换工具
 */
@Component
public class ValueTypeUtils {

    private static ConversionService conversionService;

    @Autowired
    private ApplicationContext applicationContext;

    @PostConstruct
    public void init(){
        ValueTypeUtils.conversionService = applicationContext.getBean(ConversionService.class);
    }


    /**
     * 转换类型
     * @param value 值对象
     * @param targetType 目标类型
     * @return
     */
    public static Object convertValueType(Object value,Class<?> targetType){
        if(conversionService.canConvert(value.getClass(),targetType)){
            return conversionService.convert(value,targetType);
        }
        throw new InvalidArgumentException(String.format("不支持的类型转换: %s -> %s",value.getClass(),targetType));
    }

}
