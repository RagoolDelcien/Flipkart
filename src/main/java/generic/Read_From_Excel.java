package generic;

import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;



public class Read_From_Excel implements Constants {
	String value;
	
	public String get_from(String Sheet, int row,int cell)
	{
		try {
			FileInputStream fis=new FileInputStream(Path_Of_Excel);
			Workbook wb = WorkbookFactory.create(fis);
			Cell c = wb.getSheet(Sheet).getRow(row).getCell(cell);
			 value = c.toString();
		} 
		catch (Exception e)
		{
			
			
		}
		
		
		return value;
	}
	
	
	
}
