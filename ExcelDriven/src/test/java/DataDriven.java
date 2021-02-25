import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {

	// Once the test case column is idenditied, scan the entire column to identify the desired test case row
	// After you identified the desired test case row, pull all the data on that row and feed into the test
	
	public static void main(String[] args) throws IOException {
		
		DataDriven d = new DataDriven();
		ArrayList<String> data = d.getData("addProfile");
		System.out.println(data.get(0));
		System.out.println(data.get(1));
		System.out.println(data.get(2));
		System.out.println(data.get(3));
		
		System.out.println("Testing git changes to file");
		System.out.println("Testing git changes to file");
	}
	
	@SuppressWarnings("deprecation")
	public ArrayList<String> getData(String testCaseName) throws IOException {
		
		ArrayList<String> alist = new ArrayList<String>();
		
		// file inputStream argument
		FileInputStream fis = new FileInputStream("C:\\Users\\vogonor\\Documents\\dataDemo.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int numberOfSheets = workbook.getNumberOfSheets();
		//System.out.println(numberOfSheets);
		for(int i = 0; i < numberOfSheets; i++) {
			
			if(workbook.getSheetName(i).equalsIgnoreCase("testData")) {
				
				XSSFSheet sheet = workbook.getSheetAt(i);
				
				// Identify test cases column by scanning the entire first row
				Iterator<Row> rows = sheet.iterator(); // sheet is collection of rows
				
				Row firstRow = rows.next();
				Iterator<Cell> ci = firstRow.cellIterator();        //firstRow.iterator(); //row is collection of cells
				int k = 0;
				int mycolumn = 0;
				
				while(ci.hasNext()) {
					Cell value = ci.next();
					System.out.println(value);
					if(value.getStringCellValue().equalsIgnoreCase("TestCases")) {
						
						//desired column
						mycolumn = k;
					}
					k++;
				}
				System.out.println("TestCase column is: "+mycolumn);
				
				// Once column is identified, scan entire testcase
				while(rows.hasNext()) {
					
					Row r = rows.next();
					if(r.getCell(mycolumn).getStringCellValue().equalsIgnoreCase(testCaseName)){
						
						// After you identified the desired test case row, pull all the data on that row and feed into the test
						Iterator<Cell> cv = r.cellIterator();
						while(cv.hasNext()) {
							
							Cell c = cv.next();
							if(c.getCellType()== CellType.STRING) {
								alist.add(c.getStringCellValue());
							}else {
								alist.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}
							
						}
					}
				}
				
			}
			workbook.close();
		}
		return alist;

	}

}
