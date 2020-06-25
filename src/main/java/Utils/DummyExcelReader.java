package Utils;
import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DummyExcelReader {

	private static String ExcelSheetPath = "./src/main/java/TestData/TestData.xlsx";
	private static FileInputStream fis;
	private static XSSFWorkbook workbook;
	private static Map<String, Map<String, String>> parentMap = new HashMap<String, Map<String, String>>();

	private  void loadExcel() {
		System.out.println("Load Excel Sheet.........");
		File file = new File(ExcelSheetPath);
		try {
			fis = new FileInputStream(file);
			workbook = new XSSFWorkbook(fis);
			fis.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	private  Map<String, Map<String, String>> createExcelDataMap() {
		loadExcel();
		XSSFSheet sheet;
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			sheet = workbook.getSheetAt(i);
			Iterator<Row> rowIterator = sheet.iterator();
			Map<String, String> childMap = new HashMap<String, String>();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				String key = row.getCell(0).getStringCellValue().trim();
				String value = row.getCell(1).getStringCellValue().trim();
				// don't just dump data, check if it has a value before putting into map.
				if (!key.isEmpty()) {
					childMap.put(key, value);
				}
			}
			parentMap.put(sheet.getSheetName(), childMap);
		}
		System.out.println(parentMap); 
		return parentMap;
	}

	// if you want to pass sheetname everytime
	public  String getValue(String sheetName, String key) {
		if (parentMap.isEmpty()) {
			createExcelDataMap();
		}
		return parentMap.get(sheetName).get(key);
	}

	// if u just want to get from login sheet
	public  String getValueFromLoginDataSheet(String key) {
		return getValue("LoginData", key);
	}

	public static void main(String[] args) {
		//System.out.println(getValueFromLoginDataSheet("Browser"));
//		System.out.println(getValue("Level2_Login", "USER_NAME"));
//		
	}


}
