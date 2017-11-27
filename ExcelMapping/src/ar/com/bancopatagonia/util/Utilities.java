package ar.com.bancopatagonia.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Utilities {
	
	public static String dataString() {
		DateFormat dateFormat = new SimpleDateFormat("yyyyMMdd");
		Date date = new Date();
		String datestring = dateFormat.format(date);
		return datestring;
	}
	
	public static String dataTimeString() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd-HH:mm:ss");
		Date date = new Date();
		String datestring = dateFormat.format(date);
		return datestring;
	}

	public static void fileCopyer(String fileasked, String fName) {

		String origenStringFile = fileasked;

		String destStringFile = fName;
		File orignFile = new File(origenStringFile);
		File destFile = new File(destStringFile);
		FileInputStream archivoSelected = null;
		FileOutputStream archivoSalida = null;
		try {

			archivoSelected = new FileInputStream(orignFile);
			archivoSalida = new FileOutputStream(destFile);
			byte[] byteBuf = new byte[16384];
			int numBytesRead;
			while ((numBytesRead = archivoSelected.read(byteBuf)) != -1) {
				archivoSalida.write(byteBuf, 0, numBytesRead);
			}
			archivoSalida.close();
			archivoSelected.close();

		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
	
	public static Sheet getSheetToMap (Workbook wb, String sheetName) {
		Sheet sheet = null;
		// System.out.println("cantidad de hojas: " +wb.getNumberOfSheets());
		
		for (int i = 0; i < wb.getNumberOfSheets(); i++) {
			// System.out.println("Nombre Hoja: " +
			// wb.getSheetAt(i).getSheetName());
			if (wb.getSheetAt(i).getSheetName().equals(sheetName)) {
				sheet = wb.getSheetAt(i);
			}
		}
		System.out.println("sheetName: " + sheet.getSheetName());
		return sheet;
		
	}
	
	public static Workbook getWorkBook (String excelPath) throws Exception{
		File excelOut = new File(excelPath);
		// System.out.println("Work book" + wb);
		return WorkbookFactory.create(excelOut);
	}
	
	public static void createOutExcel(Workbook wb, String destStringFile) throws Exception {
		FileOutputStream archivoSalida = new FileOutputStream(destStringFile);
		wb.write(archivoSalida);
		
	}
}
