package com.upupuup.excel;

import com.upupuup.excel.utils.DataType;
import com.upupuup.excel.utils.ImportClass;
import lombok.Data;

import java.io.Serializable;

/**
 * @Author: jiangzhihong
 * @CreateDate: 2019/8/6 8:58
 * @Version: 1.0
 * @Description: [java类作用描述]
 */
@Data
public class ExcelModel implements Serializable {
	private static final long serialVersionUID = 1L;

	/**
	 * 商品数量
	 */
	@ImportClass(excelColumn = 2, valueType = DataType.INT)
	private Integer num;

	/**
	 * 商品名称
	 */
	@ImportClass(excelColumn = 0)
	private String productName;

	/**
	 * 商品编号
	 */
	@ImportClass(excelColumn = 1, valueType = DataType.LONG)
	private long productNo;
}
