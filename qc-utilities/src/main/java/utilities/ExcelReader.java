package utilities;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

	//*Main function to call all excel related functions
	public static void main(String[] args) {
		getRowCount();
		getcelldataString(1,0);
		getcelldataNumber(1,1);
	}

	//* Function to fetch the row count from excel in which data is present
	public static void getRowCount()  {

		try {
			//String ProjectPath=System.getProperty("user.dir");
			XSSFWorkbook workbook = new XSSFWorkbook("C:\\Users\\user\\Desktop\\Automation_1\\Automation-testng\\ui-tests\\src\\test\\resources\\test-data\\TestData.xlsx");
			XSSFSheet sheet= workbook.getSheet("Sheet1");
			int rowCount = sheet.getPhysicalNumberOfRows();
			System.out.println("No of rows : "+rowCount);

		} catch (Exception exp ) {
			System.out.println(exp.getMessage());;
			System.out.println(exp.getCause());
			exp.printStackTrace();
		}
	}

	//* Function to fetch the String data from the Cell of excel

	public static void getcelldataString(int rowNum, int colNum)  {

		try {
			//String ProjectPath=System.getProperty("user.dir");
			XSSFWorkbook workbook = new XSSFWorkbook("C:\\Users\\user\\Desktop\\Automation_1\\Automation-testng\\ui-tests\\src\\test\\resources\\test-data\\TestData.xlsx");
			XSSFSheet sheet= workbook.getSheet("Sheet1");
			String CellData= sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
			System.out.println(CellData);

		} catch (Exception exp ) {
			System.out.println(exp.getMessage());;
			System.out.println(exp.getCause());
			exp.printStackTrace();
		}
	}
	//* Function to fetch the  Numeric data from the Cell of excel

		public static void getcelldataNumber(int rowNum, int colNum)  {

			try {
				//String ProjectPath=System.getProperty("user.dir");
				XSSFWorkbook workbook = new XSSFWorkbook("C:\\Users\\user\\Desktop\\Automation_1\\Automation-testng\\ui-tests\\src\\test\\resources\\test-data\\TestData.xlsx");
				XSSFSheet sheet= workbook.getSheet("Sheet1");
				double CellData= sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
				System.out.println(CellData);

			} catch (Exception exp ) {
				System.out.println(exp.getMessage());;
				System.out.println(exp.getCause());
				exp.printStackTrace();
			}
		}
}
