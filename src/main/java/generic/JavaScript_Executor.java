package generic;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

public class JavaScript_Executor {
	JavascriptExecutor js;
	WebDriver driver;
	public JavaScript_Executor(WebDriver driver)
	{
		this.driver=driver;
	}
	
	
	
	public void sendvalue(String id, String value)
	{
		
		 js=(JavascriptExecutor) driver;
		
		js.executeScript("document.getElementById("+id+") value="+value+"");
	
	}
	
	
	public void scroll_To(WebElement ele)
	{
		int y = ele.getLocation().y;
		 js=(JavascriptExecutor) driver;
		
		js.executeScript("window.scrollBy(0,"+y+")");
	
	}
	
	
	
	
}
