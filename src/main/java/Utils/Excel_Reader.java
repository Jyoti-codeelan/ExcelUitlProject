package Utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_Reader {
	
	
	public static Workbook book;
	//public static Sheet sheet;
	public static FileInputStream fis;
	public static XSSFWorkbook workbook;
	public static XSSFSheet sheet;

	public static String TESTDATA_SHEET_PATH ="./src/main/java/TestData/TestData.xlsx";
	
	
	public  static void getTestData(){
	
	
	File file = new File(TESTDATA_SHEET_PATH);

    try {
        fis = new FileInputStream(file);
        workbook = new XSSFWorkbook(fis);
        sheet = workbook.getSheet("LoginData");

    } catch (FileNotFoundException e) {
        e.printStackTrace();
    } catch (IOException e) {
        e.printStackTrace();
    }


    try {
        fis.close();
    } catch (IOException e) {
        e.printStackTrace();
    }   
    //return data;
}

	
	public static  Map<String,Map<String,String>> getDataMap() { 
        if(sheet==null){
        	getTestData();
        }

        Map<String, Map<String,String>> parentMap = new HashMap<String, Map<String,String>>();
        Map<String, String> childMap = new HashMap<String, String>();

        Iterator<Row> rowIterator = sheet.iterator();

        while( rowIterator.hasNext() )
        {
            Row row = rowIterator.next();
            childMap.put(row.getCell(0).getStringCellValue(), row.getCell(1).getStringCellValue());
        }

        parentMap.put("MASTERDATA", childMap);

        return parentMap;


    }
	
	
	
	public static  String getValue(String key) {
		
        Map<String,String> mapValue = getDataMap().get("MASTERDATA");
        String retValue = mapValue.get(key);

        return retValue;
    }
	
	public static void main(String []args){
		System.out.println(getValue("Browser"));
		System.out.println(getValue("Level1_Login"));
	}
}















