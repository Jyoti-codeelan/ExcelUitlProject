package Utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MasterData {

	
	public static final String ExcelFile_Path = "./src/main/java/TestData/TestData.xlsx";
	
	private static FileInputStream fis;
	private static XSSFWorkbook workbook;
	private static XSSFSheet sheet;
	private static XSSFRow row;
	
	
	public static void loadExcel(){
		File file = new File(ExcelFile_Path);
		try {
			fis = new FileInputStream(file);
			
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheet("ConfigData_Sheet");
			fis.close();
			
		} catch (FileNotFoundException e) {
	
			e.printStackTrace();
		} catch (IOException e) {
			
			e.printStackTrace();
		}
	}
	
	
	public static List<Map<String, String>> readAllData(){
		
		if(sheet==null){
			loadExcel();
		}
		
		List< Map< String, String > > listMap = new ArrayList<>();
		int rows = sheet.getLastRowNum();
		row = sheet.getRow(0);
		
		for(int j=1;j<row.getLastCellNum();j++){
			Map<String , String> myMap = new HashMap<>();
			
			for(int i = 1; i<rows+1; i++){
				Row r = CellUtil.getRow(i, sheet);
				String key = CellUtil.getCell(r, 0).getStringCellValue();
				String value = CellUtil.getCell(r, j).getStringCellValue();
				
				myMap.put(key, value);
			}
			listMap.add(myMap);
		}
		return listMap;
	}
	
	public static void retriveData(List< Map< String, String > > readAllData){
		
		for(Map<String, String> map : readAllData){
			map.get("Status");
		}
		
	}
	
	public static void main(String[] args) {
	

		System.out.println(readAllData());
	}

}
