package ar.com.bancopatagonia.main;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Properties;

import ar.com.bancopatagonia.mapper.ExcelMapper;
import ar.com.bancopatagonia.util.Utilities;

public class Execute {

	public static void main(String[] args) {
		Properties prop = new Properties();
		InputStream input = null;
		
		try {
			input = new FileInputStream( args[0]);
			prop.load(input);
		} catch (Exception e) {
			
			System.err.println(e.getMessage());
			System.exit(1);
		}
		
		ExcelMapper excelMapper = new ExcelMapper();
		String templatePath = prop.getProperty("templateFile");
		String csvFilePath = prop.getProperty("csvFilePath");
		String resultExcelPath = prop.getProperty("resultExcelPath");
		String resultExcelFile = prop.getProperty("resultExcelFile");
		String fileType = prop.getProperty("fileType");
		String sheetName = prop.getProperty("sheetName");
		
		String datestring = Utilities.dataString();
		String resultExcelFullName = resultExcelPath + resultExcelFile + datestring + "." +fileType;
		excelMapper.excecuteMapping(templatePath, csvFilePath, resultExcelFullName, sheetName);
		
	}

}
