import java.util.List;
import java.util.concurrent.TimeUnit;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//import org.apache.poi.XSSF.usermodel.*;
//import org.apache.poi.XSSF.usermodel.XSSFRow;
//import org.apache.poi.XSSF.usermodel.XSSFSheet;
//import org.apache.poi.XSSF.usermodel.XSSFWorkbook;

import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
//import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;
public class Search {

	public static void main(String[] args) throws Exception {
		
		//String product="Handbag";
		//String product1="skirt";
		//String department= "Women's Fashion";
		//String sortCondition= "Price: Low to High";
		//String resultValue= searchProduct(String product);
		
		//searchProduct(product, driver);
		//sortProduct(department,sortCondition, driver);
		//searchProduct(product1, driver);
		String [][] data;
		data=excelRead();
		for(int i=1; i<data.length;i++)
		{
			WebDriver driver = new FirefoxDriver();
			launch(driver);
			searchProduct(data[i][0],driver);
			sortProduct(data[i][1],data[i][2],driver);
			closeApp(driver);
		}
		
		
			}
	public static String [][] excelRead() throws Exception
	{
		File excel = new File("C:\\workspace\\Amazon\\TestData.xlsx");
		
		FileInputStream fis = new FileInputStream(excel);
		XSSFWorkbook wb = new XSSFWorkbook(fis); 
		XSSFSheet ws = wb.getSheet("Sheet1");
		int rowNum=ws.getLastRowNum() +1;
		int colNum= ws.getRow(0).getLastCellNum();
		String [][] data = new String [rowNum][colNum];
		for(int i=0;i<rowNum;i++)
		{
			XSSFRow row = ws.getRow(i);
			for(int j=0;j<colNum;j++)
			{
				XSSFCell cell = row.getCell(j);
				String value = celltoString(cell);
				data [i][j]= value;
				
			}
		}
		
		return data;
	}
	public static String celltoString(XSSFCell cell)
	{
		int type;
		Object result;
		type=cell.getCellType();
		switch(type)
		{
		case 0:
			result=cell.getNumericCellValue();
			break;
		case 1:
			result=cell.getStringCellValue();
			break;
			default:
				throw new RuntimeException("No support for this return type");
		}
		return result.toString();
	}
	public static void launch(WebDriver driver)
	{
		String appUrl = "http://www.amazon.com/";
		driver.get(appUrl);
        driver.manage().timeouts().implicitlyWait(2, TimeUnit.SECONDS);
        driver.manage().window().maximize();
        String pageTitle = driver.getTitle();
        String expectedpageTitle = "Amazon.com: Online Shopping for Electronics, Apparel, Computers, Books, DVDs &amp; more";
        //assertEquals("Current page title", expectedpageTitle, pageTitle);
	}
	
	public static void closeApp(WebDriver driver) {
		try{
		driver.close();
		}
		catch(Exception e)
		{
			
		}
	}
	public static void searchProduct(String prod, WebDriver driver) {
		
	 	   //  WebDriver driver = new ChromeDriver();
	           
	           

	// launch the firefox browser and open the application url

	           
	           driver.findElement(By.id("twotabsearchtextbox")).clear();
	           driver.findElement(By.id("twotabsearchtextbox")).sendKeys(prod);
	           driver.findElement(By.cssSelector("#nav-search > form > div.nav-right > div > input")).click();
	           WebElement result = driver.findElement(By.id("s-result-count"));
	           
	           if(result.isDisplayed())
	           {
	        	   String resultText= result.getText();
	        	   String [] splitResult= resultText.split(" ");
	        	   String count = splitResult[2];
	        	   System.out.println(resultText);
	        	   System.out.println(count);

	       			
	        	   //void sort(driver);
	        	   	

	           }
	           else
	           {
	        	   System.out.println("Fail");
	           }
	           
	        	   
	        		  

			

	}
	public static void sortProduct(String department, String sortCondition, WebDriver driver) throws InterruptedException {
		WebElement dept=driver.findElement(By.cssSelector("#s-result-info-bar-content > div.a-column.a-span4.a-text-right.a-spacing-none.a-span-last > div > div > span > span > a > span"));
		if(dept.isDisplayed())
		{
				driver.findElement(By.cssSelector("#s-result-info-bar-content > div.a-column.a-span4.a-text-right.a-spacing-none.a-span-last > div > div > span > span > a > span")).click();
				Dimension popOver= driver.findElement(By.cssSelector("#a-popover-content-1")).getSize();
				System.out.println(popOver);
	   	 		driver.findElement(By.linkText(department)).click();
	   		
	   	}
	   	WebElement sortBy = driver.findElement(By.id("sort"));
	   	if(sortBy.isDisplayed())
	   	{	
	   		Thread.sleep(5000);
	   		sortBy.click();
	   		new Select(sortBy).selectByVisibleText(sortCondition);
	   		//selenium.keyPressNative("10");
	   		sortBy.sendKeys(Keys.ENTER);
	   		JavascriptExecutor jse = (JavascriptExecutor)driver;
	   		jse.executeScript("scroll(0, 250);");
	   		Thread.sleep(10000);
	   		//WebElement ulList =driver.findElement(By.id("s-results-list-atf"));
	   		//List<WebElement> values = ulList.findElements(By.tagName("li"));
	   		//driver.findElement(By.cssSelector("#rot-B00OGD0UCA > div > a > div.s-card.s-card-group-rot-B00OGD0UCA.s-hidden.s-active > img")).click();
	   		//WebElement anchor= values.get(0).findElement(By.className("a-link-normal a-text-normal"));
	   		//anchor.click();
	   		//driver.findElement(By.cssSelector("#rot-B0043D2MCE > div > a > div.s-card.s-card-group-rot-B0043D2MCE.s-active > img")).click();
	   		//ulList.findElement(By.tagName("img")).click();
	   		//driver.findElement(By.cssSelector("#result_0 > div > div.a-row.a-spacing-base > div > div > a > img")).click();
	   	}
	   	
	}

}
