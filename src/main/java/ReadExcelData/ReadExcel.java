package ReadExcelData;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadExcel {

	public static void main(String[] args) {
		
		String data[][]= null;
		
		try {
			FileInputStream path = new FileInputStream("C:\\Users\\telma\\eclipse-workspace\\MukeshExcelPOI\\src\\main\\java\\ReadExcelData\\MukeshExcelData.xlsx");
			
			Workbook book = WorkbookFactory.create(path);
			
			Sheet sheet = book.getSheet("contacts");
			
			data = new String [sheet.getLastRowNum()][sheet.getRow(0).getLastCellNum()];
			
//			String data0 = sheet.getRow(0).getCell(0).getStringCellValue();
			
//			System.out.println("Data from excel is : " + data0);
			
			for(int i =0; i<sheet.getRow(0).getLastCellNum();i++) {
				for(int m = 0; m<sheet.getLastRowNum();i++) {
					data [i][m]=  sheet.getRow(i+1).getCell(m).toString();
					

					
				}
			}
			
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			
			e.printStackTrace();
		} catch (IOException e) {
			
			e.printStackTrace();
		}
		
		
		
	}

}
