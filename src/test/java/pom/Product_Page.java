package pom;

import generic.Base_Page;

import java.util.List;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;

public class Product_Page extends Base_Page
{

	@FindBy(xpath="(//div[@class='_1HmYoV _35HD7C'])[2]/div[@class='bhgxx2 col-12-12']/div[1]/div/div[1]/div")
	private List<WebElement>  brands;
	
	WebDriver driver;
	
	public Product_Page(WebDriver driver)
	{ 
		super(driver);
		this.driver=driver;
	
	}
	
	public void sortBrands(String name) throws InterruptedException
	{
		System.out.println(brands);
		
//		for(WebElement brand:brands)
//		{
//			Thread.sleep(2000);
//		
//		
//		if(brand.getText().contains(name))
//		{
//			brand.click();
//			
//	
//		}
//		}
		
	}
	
}
