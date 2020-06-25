package Utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class BaseExcel {
	
	
	public final String TESTDATA = "./src/main/java/TestData/TestData.xlsx";

	public final String CONFIGDATA_SHEET = "Config";
	public final String LOGIN_SHEET = "LoginData";
	
	protected static String testCaseID = null;
	
	

	private FileInputStream fis = null;
	private XSSFSheet excelWorkSheet = null;
	private XSSFWorkbook excelWorkBook = null;
	protected HashMap<String, String> allSheetData = null;
	protected static HashMap<String, String> ConfigData = null;
	protected static HashMap<String, String> TestData = null;

	


	
	
	protected XSSFWorkbook getExcelWorkBook(String workBookName) {
		File sourceFile = new File(workBookName);

		try {
			fis = new FileInputStream(sourceFile);
		} catch (FileNotFoundException e) {
			
		}

		try {
			excelWorkBook = new XSSFWorkbook(fis);
		} catch (IOException e) {
			
		}

		return excelWorkBook;
	}
	
	public ArrayList<String> getWebDriverDetails() {
		String currentBrowserName = getBrowserName();
		ArrayList<String> driverDetails = new ArrayList<String>();
		String sheetName = CONFIGDATA_SHEET;

		int totalRowCount = getRowCount(TESTDATA, sheetName);
		DataFormatter dataFormatter = new DataFormatter();
		String cellValue = null;

		excelWorkSheet = excelWorkBook.getSheet(sheetName);

		try {
			for (int i = 0; i < totalRowCount; i++) {
				cellValue = dataFormatter.formatCellValue(excelWorkSheet.getRow(i).getCell(0));
				if (cellValue != null && cellValue.length() > 0) {
					if (cellValue.equalsIgnoreCase(currentBrowserName)) {
						driverDetails.add(currentBrowserName);// adding current
																// browser name
																// in the
																// ArrayList
																				
						break;
					}
				}
			}
		} catch (Exception e) {
			
		}

		return driverDetails;
	}
	
	
	protected HashMap<String, String> getTestData(String excelFileName, String excelSheetName) {
		int totalColumnCount = getColumnCount(excelFileName, excelSheetName);
		

		String cellValue = null;
		String key = null;
		String value = null;
		DataFormatter dataFormatter = new DataFormatter();

		excelWorkSheet = excelWorkBook.getSheet(excelSheetName);
		HashMap<String, String> TestData = new HashMap<String, String>();

		try {
			for (int i = 0; i < totalColumnCount; i++) {
				try {
					cellValue = dataFormatter.formatCellValue(excelWorkSheet.getRow(i).getCell(1));
				} catch (Exception e) {
					
				}

				try {
					if (cellValue != null && cellValue.length() > 0) {
						key = dataFormatter.formatCellValue(excelWorkSheet.getRow(0).getCell(i)).trim();
						value = dataFormatter.formatCellValue(excelWorkSheet.getRow(0).getCell(i)).trim();
						TestData.put(key, value);
					}
				} catch (Exception e) {
					System.out.println(" Data missing from DX reports");
					
				}
			}
		} catch (Exception e) {
			
		}

		return TestData;
	}

	
	protected HashMap<String, String> getSheetDataInMap(String excelFileName, String excelSheetName) {
		int totalRowCount = getRowCount(excelFileName, excelSheetName);
		String cellValue = null;
		String key = null;
		String value = null;
		DataFormatter dataFormatter = new DataFormatter();

		excelWorkSheet = excelWorkBook.getSheet(excelSheetName);
		HashMap<String, String> mapSheetData = new HashMap<String, String>();

		try {
			for (int i = 0; i < totalRowCount; i++) {
				cellValue = dataFormatter.formatCellValue(excelWorkSheet.getRow(i).getCell(0));

				if (cellValue != null && cellValue.length() > 0) {
					key = dataFormatter.formatCellValue(excelWorkSheet.getRow(i).getCell(0));
					value = dataFormatter.formatCellValue(excelWorkSheet.getRow(i).getCell(1));
					mapSheetData.put(key, value);
				}
			}
		} catch (Exception e) {
			
		}

		return mapSheetData;
	}

	
	
	protected String getBrowserName() {
		String sheetName = CONFIGDATA_SHEET;
		String currentBrowserName = null;
		String cellValue = null;
		DataFormatter dataFormatter = new DataFormatter();

		int totalRowCount = getRowCount(TESTDATA, sheetName);

		excelWorkSheet = excelWorkBook.getSheet(sheetName);

		try {
			for (int i = 0; i < totalRowCount; i++) {
				cellValue = dataFormatter.formatCellValue(excelWorkSheet.getRow(i).getCell(0)).trim();
				if (cellValue != null && cellValue.length() > 0) {
					if (cellValue.equalsIgnoreCase("BrowserForCurrentTest")) {
						currentBrowserName = dataFormatter.formatCellValue(excelWorkSheet.getRow(i).getCell(1)).trim();
						break;
					}
				}
			}
		} catch (Exception e) {
			
		}

		return currentBrowserName;
	}
	
	
	private int getRowCount(String workBookName, String sheetName) {
		int rowCount = 0;

		try {
			if (excelWorkBook == null) {
				excelWorkBook = getExcelWorkBook(workBookName);
			}
		} catch (Exception e) {
			
		}

		try {
			rowCount = (excelWorkBook.getSheet(sheetName).getLastRowNum()) + 1;

		} catch (Exception e) {
			
		}

		return rowCount;
	}
	
	private int getColumnCount(String workBookName, String sheetName) {
		int colCount = 0;

		try {
			if (excelWorkBook == null) {
				excelWorkBook = getExcelWorkBook(workBookName);
			}
		} catch (Exception e) {
			
		}

		try {
			colCount = (excelWorkBook.getSheet(sheetName).getRow(0).getLastCellNum());

		} catch (Exception e) {
			
		}

		return colCount;
	}

	protected HashMap<String, String> getAllSheetData(String workBookName, String sheetName) {
		int totalRowCount = getRowCount(workBookName, sheetName);
		String cellValue = null;
		String key = null;
		String value = null;
		DataFormatter dataFormatter = new DataFormatter();

		excelWorkSheet = excelWorkBook.getSheet(sheetName);
		allSheetData = new HashMap<String, String>();

		try {
			for (int i = 0; i < totalRowCount; i++) {
				cellValue = dataFormatter.formatCellValue(excelWorkSheet.getRow(i).getCell(0));

				if (cellValue != null && cellValue.length() > 0) {
					key = dataFormatter.formatCellValue(excelWorkSheet.getRow(i).getCell(0));
					value = dataFormatter.formatCellValue(excelWorkSheet.getRow(i).getCell(1));
					allSheetData.put(key, value);
				}
			}
		} catch (Exception e) {
			
		}

		return allSheetData;
	}
	
	
	
	
	protected void openLoginPage(String moduleName) {
		

		try {
			TestData = getTestData(TESTDATA, getModuleSheetName(moduleName));
			
		} catch (Exception e) {
			
		}

	
	}

	protected String getModuleSheetName(String moduleName) {
		String moduleSheetName = null;

		try {
			switch (moduleName) {
			case "Login":
				moduleSheetName = LOGIN_SHEET;
				break;

			 case "Config":
			 moduleSheetName = CONFIGDATA_SHEET;
			 break;
			
		
			}
		} catch (Exception e) {
			
		}

		return moduleSheetName;
	}

	
	public static void main(String[] args){
		
	}
	
	
}
