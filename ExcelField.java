package com.bgyfw.comprehensive.util;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;


@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface  ExcelField {
    String name();

    /**
     * 验证方法参数 T t,List<T> success,List<T> error,后两者可省
     * @return
     */
    String[] validateMethod() default {};

    int order() default 0;


    boolean isExport() default true;

    String[] option() default {};

}

