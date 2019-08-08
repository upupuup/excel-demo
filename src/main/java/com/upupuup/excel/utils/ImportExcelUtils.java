package com.upupuup.excel.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @Author: jiangzhihong
 * @CreateDate: 2019/8/6 8:59
 * @Version: 1.0
 * @Description: [java类作用描述]
 */
public class ImportExcelUtils {

	/**
	 * 将excel中的数据转换成list
	 * @param inputStream   输入流
	 * @param excelFormat   excel格式 也可以是excel的文件名带格式
	 * @param targetTemplete 封装对象模板
	 * @param startRow  开始行
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> excelToList(InputStream inputStream, String excelFormat, Class<T> targetTemplete,int startRow) throws Exception {
		if (inputStream == null || targetTemplete == null || StringUtils.isEmpty(excelFormat)) {
			return null;
		}
		Workbook hssfWorkbook = null;
		if (excelFormat.endsWith(Constant.ExcelType.XLSX)) {
			// Excel 2007
			hssfWorkbook = new XSSFWorkbook(inputStream);
		} else if (excelFormat.endsWith(Constant.ExcelType.XLS)) {
			// Excel 2003
			hssfWorkbook = new HSSFWorkbook(inputStream);
		} else {
			return null;
		}
		return generateResultList(hssfWorkbook, targetTemplete, startRow);
	}


	/**
	 *
	 * @param excel  文件
	 * @param targetTemplete  封装对象
	 * @param startRow  开始行
	 * @return
	 * @throws Exception
	 */
	@SuppressWarnings("resource")
	public static <T> List<T> excelToList(File excel, Class<T> targetTemplete,int startRow) throws Exception {

		if (excel == null || targetTemplete == null) {
			return null;
		}
		InputStream is = new FileInputStream(excel);
		Workbook hssfWorkbook = null;
		if (excel.getName().endsWith(Constant.ExcelType.XLSX)) {
			// Excel 2007
			hssfWorkbook = new XSSFWorkbook(is);
		} else if (excel.getName().endsWith(Constant.ExcelType.XLS)) {
			// Excel 2003
			hssfWorkbook = new HSSFWorkbook(is);
		} else {
			return null;
		}
		return generateResultList(hssfWorkbook, targetTemplete, startRow);
	}


	private static <T> List<T> generateResultList(Workbook hssfWorkbook,Class<T> targetTemplete ,int startRow) throws Exception {

		// 获取所有的fields
		Field[] fields = targetTemplete.getDeclaredFields();
		if (fields == null || fields.length <= 0 || hssfWorkbook.getNumberOfSheets() == 0) {
			return null;
		}
		List<T> resultList = new ArrayList<T>();
		// 循环工作表Sheet
		for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
			Sheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
			if (hssfSheet == null) {
				continue;
			}
			System.out.println(hssfSheet);
			for (int rowNum = startRow<0 ? 0:startRow; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {

				T t = targetTemplete.newInstance();

				Row hssfRow = hssfSheet.getRow(rowNum);
				if(hssfRow==null) {
					break;
				}
				for (Field filed : fields) {
					filed.setAccessible(true);

					ImportClass importCalss = filed.getAnnotation(ImportClass.class);
					if (null == importCalss) {
						continue;
					} else {
						Cell cell = hssfRow.getCell(importCalss.excelColumn(), Row.RETURN_NULL_AND_BLANK);
						if (cell == null) {
							continue;
						}
						switch (importCalss.valueType()) {
							case BYTE:
								byte value = new BigDecimal(cell.getNumericCellValue()).byteValue();
								filed.set(t, value);
								break;
							case CHAR:
								char charValue = (char) new BigDecimal(cell.getNumericCellValue()).intValue();
								filed.set(t, charValue);
								break;
							case SHORT:
								short shortValue = new BigDecimal(cell.getNumericCellValue()).shortValue();
								filed.set(t, shortValue);
								break;
							case INT:
								int intValue = new BigDecimal(cell.getNumericCellValue()).intValue();
								filed.set(t, intValue);
								break;
							case FLOAT:
								float floatValue = new BigDecimal(cell.getNumericCellValue()).floatValue();
								filed.set(t, floatValue);
								break;
							case DOUBLE:
								double doubleValue = new BigDecimal(cell.getNumericCellValue()).doubleValue();
								filed.set(t, doubleValue);
								break;
							case LONG:
								long longValue = new BigDecimal(cell.getNumericCellValue()).longValue();
								filed.set(t, longValue);
								break;
							case BOOLEAN:
								boolean booleanValue = cell.getBooleanCellValue();
								filed.setBoolean(t, booleanValue);
								break;
							case STRING:
								String str = cell.toString().trim();
								if(str.endsWith(".0")) {
									str = str.substring(0, str.length()-2);
								}
								filed.set(t, str);
								break;
							case DATE:
								Date dateValue = cell.getDateCellValue();
								filed.set(t, dateValue);
								break;
							case NUM_STRING:
								NumberFormat nf = NumberFormat.getInstance();
								String value1 = nf.format(cell.getNumericCellValue());
								if (value1.indexOf(",") >= 0) {
									value1 = value1.replace(",", "");
								}
								filed.set(t, value1);
								break;
							default:
								break;
						}
					}
				}
				resultList.add(t);
			}
		}
		return resultList;
		
	}
	
	

}
