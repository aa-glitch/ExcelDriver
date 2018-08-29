import java.io.*;
import java.util.*;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;


import javax.swing.text.html.HTMLDocument.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		FileInputStream fis =  new FileInputStream("C://work//DemoExcel.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		int sheetcount = workbook.getNumberOfSheets();
		for(int i =0;i<=sheetcount;i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase("DemoExcel"));
			{
				  XSSFSheet sheet = workbook.getSheetAt(i);	
				  java.util.Iterator<Row> rows = sheet.iterator();
				 Row firstrow = rows.next();
				 java.util.Iterator<Cell> Cells = firstrow.cellIterator();
				 int k = 0;
				 int column = 0;
				 while (Cells.hasNext()) {
					 
					 Cell value = Cells.next();
					If (value.getStringCellValue().equalsIgnoreCase("Purchase"))
					{
						
						column = k;
					}
					k++;
					 
				 }
				 
				 System.out.println(column); 
				 
				 while(rows.hasNext()) {
					 
					java.util.Iterator<Row>= rows.next() ;
					 
				 }
				 
			}
			
			
			
		}
		
		
	}

}
