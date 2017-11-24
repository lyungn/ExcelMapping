package ar.com.bancopatagonia.mapper;

import java.io.FileReader;
import java.util.Date;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.supercsv.io.CsvMapReader;
import org.supercsv.io.ICsvMapReader;
import org.supercsv.prefs.CsvPreference;

import ar.com.bancopatagonia.util.Utilities;

public class ExcelMapper {

	public void excecuteMapping(String templatePath, String csvFilePath, String resultExcelPath, String sheetName) throws Exception {
		/*Setear la coordenada inicial para completar el excel*/
		Integer startRow = 8;
		Integer startCell = 3;
		Integer actualRow = 0;
		Integer actualCell = 0;
		
		//Utilities.fileCopyer(templatePath, resultExcelPath);		
		// InputStream imputExcelToRead = new FileInputStream(resultExcelPath);
		
		// System.out.println("archivo copiado: " + excelOut.getName());
		ICsvMapReader csvMapReader = new CsvMapReader(new FileReader(csvFilePath), CsvPreference.EXCEL_PREFERENCE);
		String[] columns = csvMapReader.getHeader(true);
		Map<String, String> featMap = null;
		
		Workbook wb = Utilities.getWorkBook(templatePath);
		Sheet sheet = Utilities.getSheetToMap(wb, sheetName);
		
		System.out.println("valor del A(7): "+ sheet.getRow(7).getCell(0));
		
		Row row = sheet.createRow(startRow + actualRow);
		while ((featMap = csvMapReader.read(columns)) != null) {
			
			
			Cell cell = row.createCell(startCell + actualCell);
			//Cell cell = sheet.getRow(startRow + actualRow).getCell(startCell + actualCell);
			cell.setCellValue(featMap.get("Session ID"));
			actualCell++;
			
			cell = row.createCell(startCell + actualCell);
			cell.setCellValue(featMap.get("Person ID"));
			actualCell++;
			
			cell = row.createCell(startCell + actualCell);
			cell.setCellValue(featMap.get("Operator"));
			actualCell++;
			
			cell = row.createCell(startCell + actualCell);
			cell.setCellValue(featMap.get("Period Number"));
			actualCell++;
			
			cell = row.createCell(startCell + actualCell);
			cell.setCellValue(featMap.get("Working hours"));
			actualCell++;
			
			cell = row.createCell(startCell + actualCell);
			cell.setCellValue(featMap.get("Non working hours"));
			actualCell++;
			
			cell = row.createCell(startCell + actualCell);
			cell.setCellValue(featMap.get("Enrollment date"));
			actualCell++;
			
			cell = row.createCell(startCell + actualCell);
			cell.setCellValue(featMap.get("Training deduction ID"));
			actualCell++;
			
			cell = row.createCell(startCell + actualCell);
			cell.setCellValue(featMap.get("Cancellation date"));
			actualCell++;
			
			cell = row.createCell(startCell + actualCell);
			cell.setCellValue(featMap.get("Cancellation Reason ID"));
			actualCell++;
			
			cell = row.createCell(startCell + actualCell);
			cell.setCellValue(featMap.get("Priority group ID"));
			actualCell++;
			
			
			cell = row.createCell(startCell + actualCell);
			cell.setCellValue(featMap.get("Subsidied"));
			actualCell++;
			
			cell = row.createCell(startCell + actualCell);
			cell.setCellValue(featMap.get("Finance Method ID"));
			actualCell++;
			
			cell = row.createCell(startCell + actualCell);
			cell.setCellValue(featMap.get("Training Location Type ID"));
			actualCell++;
			
			cell = row.createCell(startCell + actualCell);
			cell.setCellValue(featMap.get("Training Level ID"));
			actualCell++;
			
			cell = row.createCell(startCell + actualCell);
			cell.setCellValue(featMap.get("Training Objective ID"));
			actualCell++;
			
			cell = row.createCell(startCell + actualCell);
			cell.setCellValue(featMap.get("Comment"));
			
			
			
			System.out.println("Session ID: " + featMap.get("Session ID") );
			System.out.println("Person ID: " + featMap.get("Person ID") );
			System.out.println("Operator: " + featMap.get("Operator") );
			System.out.println("Period Number: " + featMap.get("Period Number") );
			System.out.println("Working hours: " + featMap.get("Working hours") );
			System.out.println("Non working hours: " + featMap.get("Non working hours") );
			System.out.println("Enrollment date: " + featMap.get("Enrollment date") );
			System.out.println("Training deduction ID: " + featMap.get("Training deduction ID") );
			System.out.println("Cancellation date: " + featMap.get("Cancellation date") );
			System.out.println("Cancellation Reason ID: " + featMap.get("Cancellation Reason ID") );
			System.out.println("Priority group ID: " + featMap.get("Priority group ID") );
			System.out.println("Subsidied: " + featMap.get("Subsidied") );
			System.out.println("Finance Method ID: " + featMap.get("Finance Method ID") );
			System.out.println("Training Location Type ID: " + featMap.get("Training Location Type ID") );
			System.out.println("Training Level ID: " + featMap.get("Training Level ID") );
			System.out.println("Training Objective ID: " + featMap.get("Training Objective ID") );
			System.out.println("Comment: " + featMap.get("Comment") );
			
			/* mover a siguiente fila*/
			actualRow++;
			/* reseteo la columna*/
			actualCell= 0; 
			
			row = row = sheet.createRow(startRow + actualRow);
			/* Fin de cada linea del archivo csv*/
			System.out.println("-- Fin de cada linea del archivo csv --" );
		}
		System.out.println("Fin de Mapping del CSV a EXCEL - " + new Date() );
		csvMapReader.close();
		Utilities.createOutExcel(wb, resultExcelPath);
	}
	
}
