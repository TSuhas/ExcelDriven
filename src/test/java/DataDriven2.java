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

public class DataDriven2 {

	
	public ArrayList<String> getdata(String testcasename) throws IOException {
		
ArrayList<String> a = new ArrayList<String>();
		
		
		FileInputStream fis = new FileInputStream("C:\\Users\\suhas.tupake\\Desktop\\path\\testdata.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		// to fetch how many  number of sheets are there
		int sheets = workbook.getNumberOfSheets();
		
		for(int i=0; i<sheets; i++)
		{
			
			if(workbook.getSheetName(i).equalsIgnoreCase("testdada"))
			{
				XSSFSheet sheet = workbook.getSheetAt(i);
				// identify the testcases column by scanning entire 1st  row
				
				Iterator<Row> rows = sheet.iterator();           // sheet is collection of rows
				Row firstrow = rows.next();
				Iterator<Cell> ce = firstrow.cellIterator();     //row is collection of cells
				
				int k=0;
				int column = 0;
				while(ce.hasNext())
				{
				    Cell value = ce.next();
				    if(value.getStringCellValue().equalsIgnoreCase("Data2"))
				    {
				    	// desired column
				    	column=k;
				    	
				    }
				    k++;
			    }
				System.out.println(column);
				//once column is identified then scan entire test column to identify "purchase" test case row
			while(rows.hasNext())
			{
				Row r = rows.next();
				if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testcasename))
				{
					// after you grab purchase" test case row, pull all the data of that row and feed it into that test case
					
					Iterator<Cell> cv = r.cellIterator();
					
					while(cv.hasNext()) 
					{
						 Cell c = cv.next();
						 if(c.getCellType()==CellType.STRING) {
							 a.add(cv.next().getStringCellValue());
						 }
						 else {
							 a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							
						 }
					}
					
				}
				
			}
			
			
		}
		
	}
		return a;
	}
	public static void main(String[] args) {
		
	}
}
