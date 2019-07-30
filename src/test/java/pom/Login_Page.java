package pom;

import generic.Base_Page;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;

public class Login_Page extends Base_Page {

	
	@FindBy(xpath="//input[@class='_2zrpKA _1dBPDZ']")
	private WebElement userName;
	@FindBy(xpath="//input[@type='password']")
	private WebElement pssword;
	@FindBy(xpath="//button[@class='_2AkmmA _1LctnI _7UHT_c']")
	private WebElement login;
	
	
	
	public Login_Page(WebDriver driver)
	{
		super(driver);
	}
	
	
	public void setUsername(String userId)
	{
		userName.sendKeys(userId);
	}
	
	
	public void setPassword(String Password)
	{
		pssword.sendKeys(Password);
	}
	
	public void Login()
	{
		login.click();
	}
	
	
	
	
	
	
	
}
