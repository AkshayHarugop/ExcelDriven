import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Assert;

public class dataDriven1 {

	public ArrayList<String> getData(String WorkBoook, String SheetName, String Header) throws IOException {
		ArrayList<String> a = new ArrayList<String>();
		FileInputStream fis =  new FileInputStream("D://work_dsi//MyFiles//Excel//"+WorkBoook+".xlsx");
		XSSFWorkbook Workbook = new XSSFWorkbook(fis);
		int NoOfSheet = Workbook.getNumberOfSheets();
		for(int i=0;i<NoOfSheet;i++) {
			Assert.assertTrue(Workbook.getSheetName(0).equalsIgnoreCase(SheetName));
			if(Workbook.getSheetName(i).equalsIgnoreCase(SheetName)) {
				XSSFSheet sheet = Workbook.getSheetAt(i);
				Iterator<Row> rows =  sheet.iterator();
				while(rows.hasNext()) {
					Row firstRow = rows.next();
					if(firstRow.getCell(i).getStringCellValue().equalsIgnoreCase(Header)) {
						Iterator<Cell> ce = firstRow.cellIterator();
						while(ce.hasNext()) {
							a.add(ce.next().getStringCellValue());
						}
					}
				}
			}
		}
		return a;
	}
}