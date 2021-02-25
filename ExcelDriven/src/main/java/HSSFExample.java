import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class HSSFExample {

	public static void main(String[] args) throws IOException {
		
		HSSFExample objHSSF = new HSSFExample();
		
		objHSSF.readFromExcel();
//		try {
//			objHSSF.writeToExcel();
//			
//		} catch (FileNotFoundException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}

	}
	
	// How to write to a .xls file
	
	public void writeToExcel() throws FileNotFoundException, IOException {
		
		// Create a workbook object
		
		HSSFWorkbook wb = new HSSFWorkbook();
		
		// create a worksheet
		
		HSSFSheet wSheet = wb.createSheet("FirstHSSF");
		
		// create a row
		
		HSSFRow wRow = wSheet.createRow(0);
		
		// Create column
		
		HSSFCell wCell = wRow.createCell(0);
		
		// Write something to this cell
		wCell.setCellValue("Stepham Trainer");
		
		// Save the file
		
		wb.write(new FileOutputStream("SimpleHSSF.xls"));
		
		//System.out.println("Excell file created");
		
		// Close the workbook
		
		wb.close();
	}
	
	public void readFromExcel() throws IOException {
		
		// Create an object of file input stream
		
		final FileInputStream fis = new FileInputStream("SimpleHSSF.xls");
			
			//Create a workbook object
			
			HSSFWorkbook wb = new HSSFWorkbook(fis);
			
			// refer to the worksheet
			
			HSSFSheet wSheet = wb.getSheet("FirstHSSF");
			
			HSSFRow gRow = wSheet.getRow(0);
			
			// Get the column
			HSSFCell gCell = gRow.getCell(0);
			
			// Get the value in row/column
			
			String strValue = gCell.getStringCellValue();
			
			// Print the value
			
			System.out.println(strValue);
			
			wb.close();
			
	}

}
