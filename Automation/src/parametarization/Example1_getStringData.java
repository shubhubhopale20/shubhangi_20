package parametarization;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Example1_getStringData {
public static void main(String[] args) throws EncryptedDocumentException, IOException {
	FileInputStream file=new FileInputStream("D:\\Excel1.xlsx");
	
	//1.getStringValue() -to read string data
	
	//String data = WorkbookFactory.create(file).getSheet("Sheet1").getRow(0).getCell(0).getStringCellValue();
	//System.out.println(data);
	
	//2.getnumericcallvalue()- to read numeric data
	
	//double data1 = WorkbookFactory.create(file).getSheet("Sheet1").getRow(0).getCell(1).getNumericCellValue();
	//System.out.println(data1);
	
	//int data2=(int)data1; //explicit casting
	//System.out.println(data2);
	
	//3.getbooleancellvalue- to read boolean data
	
	//boolean data3 = WorkbookFactory.create(file).getSheet("Sheet1").getRow(0).getCell(2).getBooleanCellValue();
	//System.out.println(data3);
	
   // String data4 = WorkbookFactory.create(file).getSheet("Sheet1").getRow(0).getCell(3).getStringCellValue();
   // System.out.println(data4);
	
	//4.getLastRowNum- returns numberof rows from 0th index
	//int rowsize = WorkbookFactory.create(file).getSheet("Sheet1").getLastRowNum();
	//System.out.println(rowsize);
	//System.out.println(rowsize+1);
	
	//5.getlastcellnum-returns no of values present in cell/column
	short cellSize = WorkbookFactory.create(file).getSheet("Sheet1").getRow(0).getLastCellNum();
	System.out.println(cellSize);
}
}
