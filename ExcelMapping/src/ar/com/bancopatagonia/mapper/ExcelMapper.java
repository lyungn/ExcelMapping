package ar.com.bancopatagonia.mapper;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.supercsv.io.CsvMapReader;
import org.supercsv.io.ICsvMapReader;
import org.supercsv.prefs.CsvPreference;

public class ExcelMapper {

	public void excecuteMapping(String templatePath, String csvFilePath, String resultExcelPath, String sheetName) {

		try {
			fileCopyer(templatePath,resultExcelPath);
			//InputStream imputExcelToRead = new FileInputStream(resultExcelPath);
			File excelOut = new File(resultExcelPath);
			//System.out.println("archivo copiado: " + excelOut.getName());
			ICsvMapReader csvMapReader = new CsvMapReader(new FileReader(csvFilePath), CsvPreference.EXCEL_PREFERENCE);
			Workbook wb = WorkbookFactory.create(excelOut);
			//System.out.println("Work book" + wb);
			Sheet sheet= null;
			//System.out.println("cantidad de hojas: " +wb.getNumberOfSheets());
			for (int i = 0; i < wb.getNumberOfSheets(); i++) {
			//	System.out.println("Nombre Hoja: " + wb.getSheetAt(i).getSheetName());
				if (wb.getSheetAt(i).getSheetName().equals(sheetName)) {
					
					sheet = wb.getSheetAt(i);
				}
			}
			
			System.out.println("sheetName: " + sheet.getSheetName());
			
			
			int rowNo = 0;
			

		} catch (Exception e) {
			e.printStackTrace();
		}

	}
	
    private void fileCopyer(String fileasked, String fName) {

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

}
