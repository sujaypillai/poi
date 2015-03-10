package com.sop;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map.Entry;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        System.out.println( "Hello World!" );
        FileInputStream fis;
		try {
			fis = new FileInputStream("108158.xlsx");
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet ws = wb.getSheet("Charter Status Info");
			
			XSSFRow row1 = ws.getRow(0);
			XSSFCell cell1 = row1.getCell(0);
			XSSFCell cell2 = row1.getCell(1);
			
			//System.out.println("Cell 1 :" + cell1.getStringCellValue() + ">> " + cell2.getStringCellValue());
			
			HashMap<String, Integer> labelMap = createLabelMaps(ws);
			//System.out.println(labelMap);
			for(Entry<String, Integer> key : labelMap.entrySet()){
				System.out.println(key.getKey() + ">> " + key.getValue());
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }

	private static HashMap<String, Integer> createLabelMaps(XSSFSheet ws) {
		HashMap<String, Integer> labelMap = new HashMap<String, Integer>(0);
		for(int i=0;i<=ws.getLastRowNum();i++){
			XSSFRow row = ws.getRow(i);
			if(row!=null){
				XSSFCell cell = row.getCell(0);
				if(cell.getStringCellValue().equalsIgnoreCase("Trip Number"))
					labelMap.put(cell.getStringCellValue(), i);
				else if(cell.getStringCellValue().equalsIgnoreCase("Customer Name"))
					labelMap.put(cell.getStringCellValue(), i);
				else if(cell.getStringCellValue().equalsIgnoreCase("ROUTING"))
					labelMap.put(cell.getStringCellValue(), i);
				else if(cell.getStringCellValue().equalsIgnoreCase("BROKER INFORMATION"))
					labelMap.put(cell.getStringCellValue(), i);
				else if(cell.getStringCellValue().equalsIgnoreCase("Legs"))
					labelMap.put(cell.getStringCellValue(), i);
				else if(cell.getStringCellValue().equalsIgnoreCase("Itinerary"))
					labelMap.put(cell.getStringCellValue(), i);
				else if(cell.getStringCellValue().equalsIgnoreCase("Credit Card Authorization"))
					labelMap.put(cell.getStringCellValue(), i);
				else if(cell.getStringCellValue().equalsIgnoreCase("Catering"))
					labelMap.put(cell.getStringCellValue(), i);
				else if(cell.getStringCellValue().equalsIgnoreCase("Additional Charges"))
					labelMap.put(cell.getStringCellValue(), i);
				else if(cell.getStringCellValue().equalsIgnoreCase("SHRED CREDIT CARD FORM"))
					labelMap.put(cell.getStringCellValue(), i);
				else if(cell.getStringCellValue().equalsIgnoreCase("COMPILE FINAL INVOICE FOR ACCOUNTING"))
					labelMap.put(cell.getStringCellValue(), i);
				else if(cell.getStringCellValue().equalsIgnoreCase("NOTES"))
					labelMap.put(cell.getStringCellValue(), i);
				else if(cell.getStringCellValue().equalsIgnoreCase("CONTAINS"))
					labelMap.put(cell.getStringCellValue(), i);
				else if(cell.getStringCellValue().equalsIgnoreCase("CHARTER STATUS"))
					labelMap.put(cell.getStringCellValue(), i);
				
			}
		}
		return labelMap;
	}
}
