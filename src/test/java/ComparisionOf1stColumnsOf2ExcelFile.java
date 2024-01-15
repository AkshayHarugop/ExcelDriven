import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ComparisionOf1stColumnsOf2ExcelFile {

	 public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException {
	        try {
	            // Replace these file paths with your actual file paths
	            String filePath1 = "D://work_dsi//MyFiles//Excel//file1.xlsx";
	            String filePath2 = "D://work_dsi//MyFiles//Excel//file2.xlsx";
	            String reportFilePath = "D://work_dsi//MyFiles//Excel//report_highlighted.xlsx";

	            // Open the Excel files
	            FileInputStream file1 = new FileInputStream(filePath1);
	            FileInputStream file2 = new FileInputStream(filePath2);

	            // Create Workbook instances for both files
	            Workbook workbook1 = WorkbookFactory.create(file1);
	            Workbook workbook2 = WorkbookFactory.create(file2);

	            // Get the first sheet from each workbook
	            Sheet sheet1 = workbook1.getSheetAt(0);
	            Sheet sheet2 = workbook2.getSheetAt(0);

	            // Create a new workbook for the report
	            Workbook reportWorkbook = new XSSFWorkbook();
	            Sheet reportSheet = reportWorkbook.createSheet("Differences");

	            // Compare the 1st columns and highlight differences in the report
	            compareColumnsAndHighlightDifferences(sheet1, sheet2, reportSheet, 0, reportWorkbook);

	            // Save the report to a new Excel file
	            FileOutputStream reportFileOut = new FileOutputStream(reportFilePath);
	            reportWorkbook.write(reportFileOut);

	            // Close all the files
	            file1.close();
	            file2.close();
	            reportFileOut.close();
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }

	    private static void compareColumnsAndHighlightDifferences(Sheet sheet1, Sheet sheet2, Sheet reportSheet, int columnIndex, Workbook workbook) {
	        int lastRow1 = sheet1.getLastRowNum();
	        int lastRow2 = sheet2.getLastRowNum();
	        System.out.println(lastRow1);
	        System.out.println(lastRow2);

	        int lastRow = Math.max(lastRow1, lastRow2);
	        System.out.println(lastRow);

	        int rowNum = 0; // Row number for the report sheet

	        for (int i = 0; i <= lastRow; i++) {
	            Row row1 = sheet1.getRow(i);
	            Row row2 = sheet2.getRow(i);

	            if (row1 != null && row2 != null) {
	                Cell cell1 = row1.getCell(columnIndex);
	                Cell cell2 = row2.getCell(columnIndex);
	                
	                String value1 = cell1 == null ? "" : cell1.toString().trim();
	                String value2 = cell2 == null ? "" : cell2.toString().trim();

	                Row reportRow = reportSheet.createRow(rowNum++);
	                reportRow.createCell(0).setCellValue(value1);
	                reportRow.createCell(1).setCellValue(value2);

	                // Compare values and highlight differences in the report
	                if (!value1.equals(value2)) {
	                    CellStyle highlightStyle = workbook.createCellStyle();
	                    Font font = workbook.createFont();
	                    font.setColor(IndexedColors.RED.getIndex());
	                    highlightStyle.setFont(font);

	                    Cell reportCell1 = reportRow.getCell(0);
	                    Cell reportCell2 = reportRow.getCell(1);

	                    reportCell1.setCellStyle(highlightStyle);
	                    reportCell2.setCellStyle(highlightStyle);
	                }
	            }
	        }
	    }
}
