package fileOperations;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {
	
	static String value;

	public static void all_Data() throws IOException {
		File f = new File("C:\\Users\\sony\\eclipse-workspace\\DataDriven1\\Task_13.xlsx");
		FileInputStream fs = new FileInputStream(f);
		Workbook wb= new XSSFWorkbook(fs);
		Sheet sheet = wb.getSheet("Sheet1");
		int numberOfRows = sheet.getPhysicalNumberOfRows();
		for (int i = 0; i < numberOfRows; i++) {
			Row row = sheet.getRow(i);
			int numberOfCells = row.getPhysicalNumberOfCells();
			for (int j = 0; j < numberOfCells; j++) {
				Cell cell = row.getCell(j);
				CellType cellType = cell.getCellType();
				
				if(cellType.equals(CellType.STRING)) {
					String stringCellValue = cell.getStringCellValue();
					System.out.print(stringCellValue);
				}
				else if (cellType.equals(CellType.NUMERIC)) {
					double numericCellValue = cell.getNumericCellValue();
					int value = (int) numericCellValue;
					System.out.print(value);
				}
				System.out.print("-");
			}
			System.out.println();
		}
		
	}
	
	public static void Write_Data() throws IOException {
		File f = new File("C:\\Users\\sony\\eclipse-workspace\\DataDriven1\\Task_13.xlsx");
		FileInputStream fs = new FileInputStream(f);
		Workbook wb= new XSSFWorkbook(fs);
		Sheet sheet = wb.createSheet("Project1");
		
		Row crRow = sheet.createRow(0);

		Cell crCell = crRow.createCell(0);
		Cell crCell1 = crRow.createCell(1);
		Cell crCell2 = crRow.createCell(2);

		crCell.setCellValue("Name");
		crCell1.setCellValue("Age");
		crCell2.setCellValue("E-mail");
		
		wb.getSheet("Project1").createRow(1).createCell(1).setCellValue("28");

		FileOutputStream fileOutputStream = new FileOutputStream(f);
		wb.write(fileOutputStream);
		wb.close();

	}
	
	public static void main(String[] args) throws IOException {
		//DataDriven.Write_Data();
		DataDriven.all_Data();
	}
	
}
