import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {

	// Identify test case coulumn by scanning the entire 1st row
	// Once the column is identified then scan entire TC column to identify the purchase test case row
	// after you grab purchase test case roew pull the all the data row and feed
	// into test

	public ArrayList<String> getData(String TestCaseName, String ExcelBook) throws IOException {
		ArrayList<String> a = new ArrayList<String>();
		FileInputStream fis = new FileInputStream("D://work_dsi//MyFiles//Excel//"+ExcelBook+".xlsx");
		XSSFWorkbook Workbook = new XSSFWorkbook(fis);
		int sheetNo = Workbook.getNumberOfSheets();
		for (int i = 0; i < sheetNo; i++) {
			if (Workbook.getSheetName(i).equalsIgnoreCase("testdata")) {
				XSSFSheet sheet = Workbook.getSheetAt(i);
				System.out.println(sheet.getSheetName());

				// Identify test case column by scanning the entire 1st row
				Iterator<Row> rows = sheet.iterator();
				Row firstrow = rows.next();
				Iterator<Cell> ce = firstrow.cellIterator();
//				Cell cell = ce.next();
//				System.out.println(cell.getStringCellValue()); 
				
				int k = 0;
				int coloumn = 0;
				while (ce.hasNext()) {
					Cell value = ce.next();
					if (value.getStringCellValue().equalsIgnoreCase("Testcases")) {
						// Desired Column
						coloumn = k;
					}
					k++;
				}
				System.out.println(coloumn);
				
				
				
// Once the column is identified then scan entire TC column to identify the purchase test case row
				while(rows.hasNext()) {
					Row row = rows.next();
					if(row.getCell(coloumn).getStringCellValue().equalsIgnoreCase(TestCaseName)) {
// after you grab purchase test case row pull the all the data row and feed
						Iterator<Cell> cv = row.cellIterator();
						while(cv.hasNext()) {
							Cell c = cv.next();
							if(c.getCellTypeEnum()==CellType.STRING) {
								a.add(c.getStringCellValue());
							}else {
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}
						
//							System.out.println(cv.next().getStringCellValue());
//							a.add(cv.next().getStringCellValue());
						}
						
					}
				}
				
//				System.out.println(a);
			}
		}
		return a;
	}
	

}
