package generic;

import java.util.List;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;

public class Select_Generic 
{
	Select s;
	WebElement ele;
	
public Select_Generic(WebElement ele)

{
	s=new Select(ele);	
}



	public void deselectIndex(int index)
	{
		
		
		s.deselectByIndex(index);
		
	}
	public void deselectValue(String value)
	{
		
		
		s.deselectByValue(value);
		
	}
	public void deselectVisibleText(String Text)
	{
	
		
		s.deselectByVisibleText(Text);
		
	}
	public void selectIndex(int index)
	{
	
		
		s.selectByIndex(index);
		
	}
	public void selectValue(String value)
	{
		
		
		s.deselectByValue(value);
		
	}
	public void selectVisibleText(String Text)
	{
		
		
		s.deselectByVisibleText(Text);
		
	}
	
	public boolean Multiple()
	{
		s.isMultiple();
		
		return s.isMultiple();
			
	}
	
	public List<WebElement> GetAll()
	{
		
	List<WebElement> Options = s.getOptions();
		return Options;
	}
	
	public WebElement GetFirstSelected()
	{
		WebElement FirstSelected = s.getFirstSelectedOption();
		return FirstSelected;
	}
	
	public List<WebElement> GetAllSelected()
	{
		 List<WebElement> AllSelected = s.getAllSelectedOptions();
		return AllSelected;
	}
	
	
	
	public void DeselectAll()
	{
		s.deselectAll();
	}
	
}
