package pack1;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.Test;


public class Excel 
{
	@Test
	public void openapp() throws Exception
	{
FileInputStream fis= new FileInputStream("C:\\Users\\Satheesh VUSIKELA\\Desktop\\vusikela.xlsx");
Workbook wb=WorkbookFactory.create(fis);
		Sheet s=wb.getSheet("Sheet1");
		int rr=s.getLastRowNum();
		System.out.println(rr);
		for(int i=0;i<=rr;i++)
		{
			Row r=s.getRow(i);
			int cc=r.getLastCellNum();
			//System.out.println(cc);
			for(int j=0;j<cc;j++)
			{
				Cell c=r.getCell(j);
				String a=c.getStringCellValue();
				System.out.println(a);
				
			}
			
		}
		
		
		
		
	}
	
	
	
	
	
	

}
