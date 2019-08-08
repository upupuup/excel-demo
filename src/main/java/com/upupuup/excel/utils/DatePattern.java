package com.upupuup.excel.utils;

import java.lang.annotation.*;

@Documented
@Target(value={ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface DatePattern {

	String pattern() default "yyyy-MM-dd HH:mm:ss";
}
