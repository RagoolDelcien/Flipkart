package generic;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Write_To_Excel implements Constants {

	
	public void Write_To(String Sheet,String value,int row, int cell)
	{
		FileInputStream fis;
		try {
			fis = new FileInputStream(Path_Of_Excel);
			Workbook wb = WorkbookFactory.create(fis);
			Cell c = wb.getSheet(Sheet).createRow(row).createCell(cell);
			c.setCellValue(value);
			FileOutputStream fos=new FileOutputStream(Path_Of_Excel);
			wb.write(fos);
		} 
		catch (Exception e) 
		{
			
		}
		
		
		
	}
	
	
	
	
	
	
}
