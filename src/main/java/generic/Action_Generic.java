package generic;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;

public class Action_Generic {

	
	Actions a1;
	
	Action_Generic(WebDriver driver)
	
	{
		 
		this. a1=new Actions(driver);
	}
	
	
	public void RightClick()
	{
		a1.contextClick().perform();
	}
	
	
	public void MouseHover(WebElement ele)
	{
		a1.moveToElement(ele).perform();
		
	}
	
	public void DragDrop(WebElement source,WebElement target)
	{
		a1.dragAndDrop(source, target).perform();
		
	}
	
	public void doubleClick()
	{
		a1.doubleClick().perform();
		
	}
	
	
	
}
