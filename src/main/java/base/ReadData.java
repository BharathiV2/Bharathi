package base;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadData {
	
	public String readData(String colName) throws IOException
	{
//		String path=System.getProperty("C:\\ProgramData\\Eclipse\\eclipse-workspace\\eclipse-workspace\\HotelBooking\\src\\test\\resources\\testData\\BookingDetails.xlsx");
		FileInputStream fis = new FileInputStream(new File("C:\\ProgramData\\Eclipse\\eclipse-workspace\\eclipse-workspace\\HotelBooking\\src\\test\\resources\\testData\\BookingDetails.xlsx"));
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet bookingDetails = workbook.getSheetAt(0);
		String colValue= "";
		for(int i=1;i<bookingDetails.getLastRowNum();i++)
		{
			XSSFRow row = bookingDetails.getRow(i);
					
			if(row.getCell(0).getStringCellValue().equalsIgnoreCase(colName)) {
				colValue = row.getCell(1).getStringCellValue();
			}
		}
		workbook.close();
		
		return colValue;
		
	}

}
