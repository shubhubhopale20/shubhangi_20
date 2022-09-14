package parametarization;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Example3_print1cell {
public static void main(String[] args) throws EncryptedDocumentException, IOException {
	FileInputStream file=new FileInputStream("D:\\Excel1.xlsx");
	Sheet sh = WorkbookFactory.create(file).getSheet("Sheet3");
	
	int rowSize=sh.getLastRowNum();
	
	for(int i=0;i<=rowSize;i++)
	{
		double data = sh.getRow(i).getCell(0).getNumericCellValue();
		
		System.out.println(data);
	}
	
	 
}
}
