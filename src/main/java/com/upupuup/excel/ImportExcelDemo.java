package com.upupuup.excel;

import com.upupuup.excel.utils.ImportExcelUtils;
import org.springframework.util.CollectionUtils;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

/**
 * @Author: jiangzhihong
 * @CreateDate: 2019/8/6 9:14
 * @Version: 1.0
 * @Description: [java类作用描述]
 */
@RestController
@RequestMapping("/excel")
public class ImportExcelDemo {
	/**
	 * 导入excel
	 * @param file
	 */
	@PostMapping("/importExcel")
	public List<ExcelModel> ImportExcel(MultipartFile file) throws Exception {
		List<ExcelModel> excelModels = ImportExcelUtils.excelToList(file.getInputStream(), file.getOriginalFilename(), ExcelModel.class, 2);

		if (CollectionUtils.isEmpty(excelModels)) {
			throw new Exception("没有数据");
		}

		for (ExcelModel excelModel : excelModels) {
			System.out.println(excelModel.toString());
		}

		return excelModels;
	}
}
