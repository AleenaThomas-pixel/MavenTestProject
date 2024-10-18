package excelread;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;


public class ExcelRead {

	XSSFSheet sheet;
	
	ExcelRead() throws IOException {
	
	FileInputStream inputfile = new FileInputStream("D:\\Aleena\\Empdetails.Worksheet.xlsx");
	XSSFWorkbook workbook = new   XSSFWorkbook (inputfile);//HSSFWorkbook if file type is .xls
	
	sheet = workbook.getSheet("Sheet1");
	}
	
	public String readExcel(int i,int j) {
		
		Row row = sheet.getRow(i);
		Cell cell = row.getCell(j);
		
		CellType type = cell.getCellType();
		
		switch(type) {
		case NUMERIC:
			return String.valueOf(cell.getNumericCellValue());
		case STRING:
			return cell.getStringCellValue();
		case BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue());
		}
		return cell.getStringCellValue();
	
	}
public static void main(String args[]) throws IOException {
	
	ExcelRead ex = new ExcelRead();
	
	int r = ex.sheet.getLastRowNum();
	ex.sheet.getRow(r).getLastCellNum();
	
	System.out.println("Excel DATA : ");
	System.out.println("All Employee Details : ");
	for (int i=0; i< ex.sheet.getLastRowNum()+1;i++) {
		for (int j=0; j<= ex.sheet.getRow(i).getLastCellNum()-1;j++) {
			
			String s = ex.readExcel(i,j);
			System.out.print(s + "  ");
		}
		System.out.println();
	}
	
}

}
