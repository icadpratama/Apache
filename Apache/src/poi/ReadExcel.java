package poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	static XSSFRow row;
	
	public static void main(String[] args){
		FileInputStream fis;
		try {
			fis = new FileInputStream(new File("New Upload File (with EQ columns)20170120.xlsx"));
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			int sumOfSheets = workbook.getNumberOfSheets();
			for (int i = 0; i < sumOfSheets; i++) {
				XSSFSheet spreadsheet = workbook.getSheetAt(i);
				String sheetName = workbook.getSheetName(i);
				System.out.println("Sheet name: " + sheetName);
				for (int j = 1; j < spreadsheet.getLastRowNum(); j++) {
					for (int j2 = 0; j2 < spreadsheet.getRow(0).getLastCellNum(); j2++) {
						Cell wanted;
						wanted = spreadsheet.getRow(j).getCell(j2);
						String wantedRef = (new CellReference(wanted)).formatAsString();
						System.out.println(wantedRef);
						switch(wanted.getCellType()){
							case Cell.CELL_TYPE_NUMERIC:
								System.out.println("The value is " + workbook.getSheetAt(i).getRow(j).getCell(j2).getNumericCellValue());
								break;
							case Cell.CELL_TYPE_STRING:
								System.out.println("The value is " + workbook.getSheetAt(i).getRow(j).getCell(j2).getStringCellValue());
								break;
						}
						
					}
					
				}
			}
			
			fis.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
