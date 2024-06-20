package utilities;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataReader
{
	
	public static HashMap<String, String> storeValues = new HashMap<>();

	@SuppressWarnings("incomplete-switch")
	public static List<HashMap<String, String>> data(String filepath, String sheetName) {
		
		List<HashMap<String, String>> mydata = new ArrayList<>();
		
		try {
			FileInputStream fs = new FileInputStream(filepath);
			try (XSSFWorkbook workbook = new XSSFWorkbook(fs)) {
				XSSFSheet sheet = workbook.getSheet(sheetName);
				Row HeaderRow = sheet.getRow(0);
				for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) 
					{
					Row currentRow = sheet.getRow(i);
					HashMap<String, String> currentHash = new HashMap<String, String>();
					for (int j = 0; j < currentRow.getPhysicalNumberOfCells(); j++) 
						{
						Cell currentCell = currentRow.getCell(j);
						switch (currentCell.getCellType()) 
							{
								case STRING:
									currentHash.put(HeaderRow.getCell(j).getStringCellValue(),currentCell.getStringCellValue());
								break;
							}
						}
					mydata.add(currentHash);
					}
			}
			fs.close();
			} catch (Exception e) {
			e.printStackTrace();
		}
		return mydata;
	}
	
//	public static void main(String[] args) {
//		List<HashMap<String,String>> d = data(System.getProperty("user.dir") + "\\TestData\\ExcelData.xlsx", "Recipient_Email");
//		
//		System.out.println(d.get(1));
//	}
}

/*
package utilities;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataReader {

	public static HashMap<String,String> storeValues = new HashMap<>();
	
	@SuppressWarnings("incomplete-switch")
	public static List<HashMap<String,String>> data(String fileName, String sheetName){
		List<HashMap<String, String>> myData = new ArrayList<>();
		
		try {
			FileInputStream file = new FileInputStream(fileName);
			
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheet(sheetName);
			XSSFRow headerRow = sheet.getRow(0);
			Row currentRow;
			Cell currentCell;
			
			for(int i=1; i<sheet.getPhysicalNumberOfRows();i++) {
				currentRow = sheet.getRow(i);
				
				HashMap<String, String> currentMap = new HashMap<String, String>();
				
				for(int j=0; j<currentRow.getPhysicalNumberOfCells();j++) {
					currentCell = currentRow.getCell(j);
					switch(currentCell.getCellType()) {
						case STRING:
							currentMap.put(headerRow.getCell(j).getStringCellValue(), currentCell.getStringCellValue());
						break;
					
					}
				}
				
				myData.add(currentMap);
			}
			workbook.close();
			file.close();
		}
		catch(Exception e) {
			
		}
		
		return myData;
	}
}*/

