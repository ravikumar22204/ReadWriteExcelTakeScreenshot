/*WriteUp:
 * Login in to CMSe Application
 * Read the input from Excel Sheet
 * Create card 
 * Write the card number in Excel sheet
 * Update card Status to Active
 * Perform Add Funds
 * Take screenprints for all the above
 * Logout of the Application
 */

package read_writeExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import libraries.Utilities;


public class Demo_Q3 {
	
	public static WebDriver driver;
	public static XSSFWorkbook wb;
	public static XSSFSheet sh1,sh2;
	public static File src;
	public static String usename, pwd;
	public static int l=1;
	
	public static void main(String[] args) throws Exception {
		login();
		for(int x=1; x<=1; x++,l++) {
		createcard();
		search();
		negativeFileStatus();
		addFunds();
		deleteCard();
		}
		logout();
	}
	
	public static void readFileHandling() {
		
		try {
			  // Specify the path of file
			   src=new File("C:\\RK\\SeleniumFiles\\InputFile_Testdata.xlsx");
			 
			   // load file
			   FileInputStream fis=new FileInputStream(src);
			 
			   // Load workbook
			   wb=new XSSFWorkbook(fis);
			   
			   // Load sheet- Here we are loading first sheetonly
			   sh1= wb.getSheet("testdata1");
			   sh2=wb.getSheet("output");
			
			} catch (Exception e) {
			 
			   System.out.println(e.getMessage());
			 }
		}
	
	
	public static void login() throws IOException {
		
		readFileHandling();
		
		System.setProperty("webdriver.chrome.driver", "C:\\RK\\SeleniumFiles\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("http://vlcapweb03.fisdev.local:12143/main/cmse/Home");
		usename= sh1.getRow(1).getCell(0).getStringCellValue();
		pwd= sh1.getRow(1).getCell(1).getStringCellValue();
		driver.findElement(By.id("userid")).sendKeys(usename);
		driver.findElement(By.id("password")).sendKeys(pwd);
		Utilities.captureScreenshot(driver);
		driver.findElement(By.id("loginButton")).click();
		Utilities.captureScreenshot(driver);
	}
	
	
	public static void createcard() throws IOException {
		
			
			driver.findElement(By.linkText("Cardholder")).click();
			driver.findElement(By.id("inst")).clear();
			driver.findElement(By.id("inst")).sendKeys("00548");
			driver.findElement(By.linkText("Open New Card")).click();
			
			
			try {
				
				  if(driver.findElement(By.id("nextButton")).isDisplayed()) {
					  
					   for(int  i=1,j=3; i<26; i++,j=j+2) {
						   
						   	String selectprefix= (driver.findElement(By.xpath("//*[@id=\"main-column\"]/div/table/tbody/tr["+j+"]/td[1]"))).getText();
							String myprefix = sh1.getRow(1).getCell(2).getStringCellValue();
						 
							 if(selectprefix.equals(myprefix)) {
							
							 driver.findElement(By.xpath("//*[@id=\"main-column\"]/div/table/tbody/tr["+j+"]/td[1]")).click();
							 Utilities.captureScreenshot(driver);
							 break;	 
						 	}
							 
							 else {
								 
								 if(i==25) {
									 driver.findElement(By.id("nextButton")).click();
									 i=1; j=1;
									 continue;
								 }
							}
					   }				
				  }
			  }
			  catch(Exception e){
					 driver.findElement(By.linkText("Return to Search")).click();
					 System.out.println("Prefix Not found");
					 return;
					}
			Utilities.captureScreenshot(driver);
			try {
			  	if(driver.findElement(By.linkText("Skip Product Selection")).isDisplayed())
			  	{
			  		driver.findElement(By.linkText("Skip Product Selection")).click();	
			  	}
			  }	
			  catch(Exception p) {
			 	   System.out.println("Product ID not enabled for this prefix");
			  	}	
			
			driver.findElement(By.linkText("Generate Card Number")).click();
			driver.findElement(By.id("NameOne")).sendKeys(sh1.getRow(1).getCell(3).getStringCellValue());
			driver.findElement(By.id("Address1")).sendKeys(sh1.getRow(1).getCell(4).getStringCellValue());
			driver.findElement(By.id("City")).sendKeys(sh1.getRow(1).getCell(5).getStringCellValue());
			Select State= new Select(driver.findElement(By.id("State")));
			State.selectByVisibleText(sh1.getRow(1).getCell(6).getStringCellValue());
			driver.findElement(By.id("ZipCode")).sendKeys(sh1.getRow(1).getCell(7).getStringCellValue());
			driver.findElement(By.id("Email")).sendKeys(sh1.getRow(1).getCell(8).getStringCellValue());
			driver.findElement(By.id("PrimSSN1")).sendKeys(sh1.getRow(1).getCell(9).getStringCellValue());
			driver.findElement(By.id("Mon")).sendKeys(sh1.getRow(1).getCell(10).getStringCellValue());
			driver.findElement(By.id("GeneratePin")).click();
			driver.findElement(By.id("continueEdit")).click();
			driver.findElement(By.id("submit")).click();
			driver.findElement(By.id("submitReview")).click();
			Utilities.captureScreenshot(driver);
			String cardNumber = (driver.findElement(By.xpath("//*[@id=\"mainForm\"]/table[1]/tbody/tr[3]/td[1]")).getText());
			System.out.println(cardNumber);
			sh2.getRow(l).createCell(0).setCellValue(cardNumber);
			FileOutputStream fout = new FileOutputStream(src);
			wb.write(fout);
			driver.findElement(By.id("returnToSearch")).click();
			Utilities.captureScreenshot(driver);
			
		}
	
	public static void search() {
		
		driver.findElement(By.id("maskedCardNumber")).sendKeys(sh2.getRow(l).getCell(0).getStringCellValue());
		driver.findElement(By.id("searchSubmit")).click();
		}
	
	public static void negativeFileStatus() throws IOException {
		
		driver.findElement(By.linkText("Neg File/Status")).click();
		Utilities.captureScreenshot(driver);
		driver.findElement(By.id("Active")).click();
		Select reason= new Select(driver.findElement(By.id("cardStatusReason")));
		reason.selectByIndex(1);
		Select negativeFileUpdateCode= new Select(driver.findElement(By.id("negativeFileUpdateCode")));
		negativeFileUpdateCode.selectByIndex(4);
		driver.findElement(By.id("Submit")).click();
		Utilities.captureScreenshot(driver);
		}
	
	public static void addFunds() throws IOException {
		
		driver.findElement(By.linkText("Prepaid Debit")).click();
		driver.findElement(By.linkText("Funding Options")).click();
		driver.findElement(By.linkText("Add Funds")).click();
		Utilities.captureScreenshot(driver);
		driver.findElement(By.id("amountToAdd")).clear();
		driver.findElement(By.id("amountToAdd")).sendKeys("200.00");
		driver.findElement(By.id("Submit")).click();
		Utilities.captureScreenshot(driver);
		}
	
	public static void logout() throws IOException {
		driver.findElement(By.linkText("Log Out")).click();
		Utilities.captureScreenshot(driver);
	}
	
	public static void deleteCard() throws Exception{
		
		driver.findElement(By.linkText("Delete Card")).click();
		driver.findElement(By.id("submitButton")).click();
		driver.findElement(By.id("okId")).click();
	}
	
}
