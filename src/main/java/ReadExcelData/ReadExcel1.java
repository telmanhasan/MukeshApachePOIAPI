package ReadExcelData;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadExcel1 {

	public static void main(String[] args) {
		
		Object data[][]= null;
		
		try {
			FileInputStream path = new FileInputStream("C:\\Users\\telma\\eclipse-workspace\\MukeshExcelPOI\\src\\main\\java\\ReadExcelData\\MukeshExcelData.xlsx");
			
			Workbook book = WorkbookFactory.create(path);
			
			Sheet sheet = book.getSheet("contacts");
			
			
			
			
			 int z = sheet.getRow(0).getFirstCellNum();
			System.out.println(z);
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			
			e.printStackTrace();
		} catch (IOException e) {
			
			e.printStackTrace();
		}
		
		
		
	}

}
