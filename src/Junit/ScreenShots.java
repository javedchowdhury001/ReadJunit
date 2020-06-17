package Junit;

import java.io.File;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class ScreenShots {

	

	public static void main(String[] args) {
		System.setProperty("webdriver.chrome.driver","D:\\QA\\Selenium Automation\\1st Class\\chromedriver_win32\\chromedriver.exe" );
		WebDriver driver=new ChromeDriver();
		driver.get("https://www.google.com");
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		
	File screenShot=  ((TakesScreenshot)driver)
			.getScreenshotAs(OutputType.FILE);
	try {
		String ScreenLocation= "D:\\QA\\Selenium Automation\\Screenshot.png";
		
		FileUtils.copyFile(screenShot , new File(ScreenLocation));
		System.out.println("Screenshot Done ");
		
	
	}catch (Exception e) {
		System.out.println(e.getMessage());
	}
				
	

	}

}
