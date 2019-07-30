package generic;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;

import org.openqa.selenium.Dimension;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebDriver;

public class BrowserOption {

	public void minimize() throws AWTException, InterruptedException
	{
		Robot r=new Robot();
		r.keyPress(KeyEvent.VK_ALT);
		r.keyPress(KeyEvent.VK_SPACE);
		r.keyPress(KeyEvent.VK_N);
		Thread.sleep(1000);
		r.keyRelease(KeyEvent.VK_ALT);
		r.keyRelease(KeyEvent.VK_SPACE);
		r.keyRelease(KeyEvent.VK_N);
		
	}
	
	public void maximize(WebDriver driver)
	{
		driver.manage().window().maximize();
	}
	
	
	public void closeWindow() throws AWTException
	{
		Robot r=new Robot();
		
		r.keyPress(KeyEvent.VK_ALT);
		r.keyPress(KeyEvent.VK_SPACE);
		r.keyPress(KeyEvent.VK_C);
		r.keyRelease(KeyEvent.VK_ALT);
		r.keyRelease(KeyEvent.VK_SPACE);
		r.keyRelease(KeyEvent.VK_C);
	}

	public void Resize(WebDriver driver,int width,int height)
	{
		Dimension D=new Dimension(width, height);
		driver.manage().window().setSize(D);
	}
	
	
	public void dragBrowser(WebDriver driver,int x, int y)
	{
		Point p=new Point(x, y);
		
		driver.manage().window().setPosition(p);
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}
