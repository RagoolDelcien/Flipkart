package generic;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

public abstract class Base_Page {

	WebDriver driver;
	
	public Base_Page(WebDriver driver)
	{
		this.driver=driver;
		PageFactory.initElements(driver, this);
	}
	
	
	public void verify_title(String title)
	{
		WebDriverWait wait=new WebDriverWait(driver, 10);
		try
		{
		wait.until(ExpectedConditions.titleContains(title));
		}
		catch(Exception e)
		{
			System.out.println("page is not displayed");
			Assert.fail();
		}
	}
	
	
	public void verify_element(WebElement element)
	{
		WebDriverWait wait=new WebDriverWait(driver, 10);
		
		try
		{
		wait.until(ExpectedConditions.visibilityOf(element));
		}
		catch(Exception e)
		{
			System.out.println("element is not visible");
			Assert.fail();
		}
		
		
	}
	
	
	
	
	
	
	
	
}
