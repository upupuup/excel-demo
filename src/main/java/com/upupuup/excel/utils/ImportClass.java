package com.upupuup.excel.utils;

import java.lang.annotation.*;

/**
 * @Author: jiangzhihong
 * @CreateDate: 2019/8/6 9:14
 * @Version: 1.0
 * @Description: [java类作用描述]
 */
@Documented
@Target(value={ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ImportClass {

	/**
	 * ExcelModel 中第几列
	 * @return
	 */
	int excelColumn();
	/**
	 * 数据类型
	 * @return
	 */
	DataType valueType() default DataType.STRING;
	
	
}
