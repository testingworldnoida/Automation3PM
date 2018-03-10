package testing;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class B {

	public static void main(String[] args) throws IOException {
		
		File f = new File("E:\\FR121.xlsx");
		FileOutputStream fs = new FileOutputStream(f);
		
		XSSFWorkbook wk = new XSSFWorkbook();
		XSSFSheet s1 = wk.createSheet("Data");
		XSSFRow f1 = s1.createRow(0);
		XSSFCell c1 = f1.createCell(0);
		c1.setCellValue("Hello World");
		wk.write(fs);
		wk.close();
		
		
//		FileInputStream fs = new FileInputStream(f);
//		
//		XSSFWorkbook wk = new XSSFWorkbook(fs);
//		XSSFSheet s1 = wk.getSheet("Sheet1");
//		int r = s1.getPhysicalNumberOfRows();
//		
//		for(int i=0;i<r;i++)
//		{
//		  XSSFRow r1 = s1.getRow(i);
//		  int c = r1.getPhysicalNumberOfCells();
//		  for(int j=0;j<c;j++)
//		  {
//		    XSSFCell c1 = r1.getCell(j);
//		    System.out.print(c1.getStringCellValue()+"    ");
//		  }
//		  System.out.println();
//		}
	}
	
}
