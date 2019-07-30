package generic;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Parameters;

public class Before_Methods implements Constants{

	public WebDriver driver;
	
	static
	{
		System.setProperty(Firefox_Key, Firefox_Value);
	}
	
	@Parameters({"browser"})
	@BeforeMethod
	public void Open_App(String browse)
	{
		
		if(browse.equals("firefox"))
		{
			driver=new FirefoxDriver();
			//driver.get("https://www.gmail.com");
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		}
		
		else
		{
			driver=new ChromeDriver();
			//driver.get("https://www.gmail.com");
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		}
		
		
	}
	
	
	@AfterMethod
	public void Close_App()
	{
		driver.quit();
	}
	
	
	
	
}
