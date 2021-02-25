import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XSSFExample {

	public static void main(String[] args) throws IOException {
		
		XSSFExample ex = new XSSFExample();
		ex.writeToExcel();
		
		//ex.readFromExcel();
	}
	
	public void writeToExcel() //
	{
		// create an object of workbook
		
		XSSFWorkbook wb = new XSSFWorkbook();			
		
		//create a sheet
		
		XSSFSheet ws = wb.createSheet("myFirstSheet");
		
		// Create a row
		
		XSSFRow row = ws.createRow(0);
		
		// create a column
		
		XSSFCell xCell = row.createCell(0);
		
		// add value to the created cell and colu,m
		
		xCell.setCellValue("Stepham Tester");
		
		// Write to an output file
		
		try {
			wb.write(new FileOutputStream("SimpleXSSF.xslx"));
			wb.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}
	
	public void readFromExcel() throws IOException {
		
		// Create an object of workbook
		XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream("SimpleXSSF.xlsx"));
		
		// Create a sheet
		XSSFSheet wSheet = wb.getSheetAt(0);
		String strData = wSheet.getRow(0).getCell(0).getStringCellValue();
		System.out.println("Data value is: "+ strData);
				
		wb.close();
		
		
	}

}
