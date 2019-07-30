package scripts;

import java.awt.AWTException;
import java.awt.Robot;

import generic.Base_Page;
import generic.Select_Generic;

import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;

public class Home_Page extends Base_Page {

	
	
	@FindBy(xpath="//input[@title='Search for products, brands and more']")
	private WebElement search;
	
	public Home_Page(WebDriver driver)
	{
		super(driver);
	}
	
	
	public void SearchForProduct(String Product) throws AWTException
	{
		search.sendKeys(Product);
		search.sendKeys(Keys.ENTER);
		
	}
	
	
	
	
	
}
