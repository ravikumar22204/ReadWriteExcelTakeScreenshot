package libraries;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

public class Utilities {
	
	public static void captureScreenshot(WebDriver driver) throws IOException {
		
		TakesScreenshot ts= (TakesScreenshot)driver;
		File srce= ts.getScreenshotAs(OutputType.FILE);
		//FileUtils.copyFile(srce, new File("./Screenshots/"+screenshotName+".png"));
		//FileUtils.copyFile(srce, new File("C://RK//SeleniumFiles//sc//"+screenshotName+".png"));
		FileUtils.copyFile(srce, new File("C://RK//SeleniumFiles//sc//CreateCard//"+timestamp()+".png"));
		System.out.println("Screenshot captured");
		
	}
	
	public static  String timestamp() {
	    return new SimpleDateFormat("yyyy-MM-dd HH-mm-ss").format(new Date());
	}

}
